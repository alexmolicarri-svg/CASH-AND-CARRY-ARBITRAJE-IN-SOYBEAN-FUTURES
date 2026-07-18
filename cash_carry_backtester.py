"""
=============================================================================
ARBITRAJE CASH & CARRY EN FUTUROS AGRÍCOLAS (CBOT) — BACKTESTER HISTÓRICO
Procesador por Lotes para Análisis Empírico Retrospectivo
=============================================================================

Propósito   : Reconstrucción histórica del dataset de arbitraje Cash &
              Carry para el trabajo de investigación. Complementa al
              recolector en tiempo real (cash_carry_research.py) con un
              flujo de procesamiento por lotes: inicio → proceso → fin →
              exportación consolidada a Excel multi-hoja.

-----------------------------------------------------------------------------
RELACIÓN CON EL RECOLECTOR EN TIEMPO REAL
-----------------------------------------------------------------------------
La matemática es EXACTAMENTE la misma que en el daemon en vivo:

  1. Forward teórico con carry implícito:
         F_teorico(T) = S * exp[(r + c_neto) * T]
     donde c_neto se estima cada día desde el calendar spread del par
     de contratos cercanos:
         c_neto = ln(F_far / F_near) / (T_far - T_near)

  2. Evaluación asimétrica de las dos piernas cruzando el spread:
         Cash & Carry     : compra spot al ASK, vende futuro al BID
         Reverse C&C      : vende spot al BID, compra futuro al ASK

  3. Neteo contra costes de transacción institucionales (tx_cost_pct).

-----------------------------------------------------------------------------
TRATAMIENTO DEL BID-ASK EN DATOS HISTÓRICOS (decisión metodológica)
-----------------------------------------------------------------------------
Los datos históricos gratuitos (cierres diarios) no contienen el order
book. Este script resuelve el problema en dos niveles, priorizando
siempre el dato real:

  NIVEL 1 — BID_ASK bars de IBKR (preferente):
      reqHistoricalData con whatToShow='BID_ASK' devuelve, para cada
      barra diaria, el promedio de bid (campo open) y el promedio de
      ask (campo close) del periodo. Es microestructura real observada,
      no una aproximación. Las filas construidas así llevan
      spread_method='IBKR_BID_ASK'.

  NIVEL 2 — TRADES + half-spread sintético (fallback declarado):
      Si la suscripción no permite BID_ASK para un contrato, se usan
      cierres de TRADES y se aplica un half-spread simétrico de
      `assumed_half_spread_ticks` ticks por lado (tick de soja CBOT =
      0.25 ¢/bu). Las filas llevan spread_method='SYNTHETIC_SPREAD' y
      la columna assumed_half_spread_usd documenta el supuesto exacto,
      permitiendo análisis de sensibilidad posterior o exclusión.

Ambos métodos NUNCA se mezclan de forma silenciosa: cada observación
declara su procedencia, preservando la trazabilidad exigible en una
revisión académica.

-----------------------------------------------------------------------------
TASA LIBRE DE RIESGO HISTÓRICA
-----------------------------------------------------------------------------
Treasury.gov publica el CSV completo del año con la curva diaria. El
script lo descarga UNA vez, construye una serie indexada por fecha
(tramo 6 Mo) y asigna a cada observación la tasa vigente en SU fecha
(merge_asof hacia atrás: la última tasa publicada ≤ fecha de la
observación, replicando la información disponible en ese momento).
=============================================================================
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from io import StringIO
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import requests
from ib_insync import IB, Future, util

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("cash_carry_backtester")
util.logToConsole(logging.ERROR)

# Tick mínimo de la soja en CBOT: 1/4 de centavo por bushel (en USD/bu)
ZS_TICK_USD = 0.0025

# Carpeta donde vive este script — el Excel se guarda AQUÍ siempre,
# sin importar el directorio de trabajo desde el que se invoque python.
# Esto evita que el archivo aparezca "perdido" en C:\Users\<usuario>\
# cuando el script se ejecuta con una ruta absoluta desde otra carpeta.
SCRIPT_DIR = Path(__file__).resolve().parent


# =========================================================================
# CONFIGURACIÓN DEL ESTUDIO
# =========================================================================

@dataclass
class BacktestConfig:
    """
    Parametrización completa del estudio retrospectivo. Centralizar la
    configuración en un dataclass permite reportar en el paper, de forma
    literal, todos los supuestos con los que se generó cada dataset.
    """

    symbol: str = "ZS"
    exchange: str = "CBOT"
    duration: str = "6 M"            # ventana histórica hacia atrás desde hoy
    bar_size: str = "1 day"
    tx_cost_pct: float = 0.001       # costes institucionales (0.10%)
    assumed_half_spread_ticks: int = 2   # fallback: 2 ticks por lado
    # Los futuros de granos de CBOT (ZS, ZC, ZW) cotizan en CENTAVOS por
    # bushel (p.ej. 1087.5 = $10.875/bu). price_scale convierte a USD/bu
    # dividiendo por 100. Para instrumentos ya en USD, poner 1.0.
    price_scale: float = 100.0
    # Índices de la cadena (ordenada por vencimiento) para spot/near/far/
    # target. Por defecto se usan los tres primeros contactos consecutivos.
    # ADVERTENCIA METODOLÓGICA: si spot y target pertenecen a cosechas
    # distintas (old crop vs new crop en soja: Jul vs Nov), el modelo de
    # cost-of-carry NO aplica limpiamente y el "spread" capturará prima
    # estacional, no arbitraje. Ver PDF de resultados.
    idx_near: int = 0
    idx_far: int = 1
    idx_target: int = 2
    host: str = "127.0.0.1"
    port: int = 7496
    client_id: int = 11              # distinto del daemon en vivo para no colisionar
    output_xlsx: str = field(
        default_factory=lambda: str(
            SCRIPT_DIR / f"cash_carry_backtest_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
    )

    @property
    def assumed_half_spread_usd(self) -> float:
        return self.assumed_half_spread_ticks * ZS_TICK_USD


# =========================================================================
# COMPONENTE 1 · TASA LIBRE DE RIESGO HISTÓRICA (serie diaria)
# =========================================================================

class HistoricalRiskFreeRateProvider:
    """
    Descarga la curva diaria del Tesoro (año actual y, si la ventana lo
    requiere, el anterior) y expone una serie fecha → tasa 6M decimal.

    La asignación posterior por fecha se hace con merge_asof(direction=
    'backward'), que replica la información realmente disponible en cada
    fecha de observación (no hay look-ahead bias).
    """

    DEFAULT_RATE = 0.045

    @classmethod
    def build_series(cls, start_year: int) -> tuple[pd.Series, str]:
        frames: list[pd.DataFrame] = []
        current_year = datetime.now().year
        for year in range(start_year, current_year + 1):
            df = cls._fetch_year(year)
            if df is not None:
                frames.append(df)

        if not frames:
            logger.warning(
                "No se pudo obtener ninguna tasa histórica de Treasury.gov. "
                "Se usará la tasa constante por defecto %.2f%%.",
                cls.DEFAULT_RATE * 100,
            )
            return pd.Series(dtype=float), "DEFAULT_CONSTANT_RATE"

        rates = pd.concat(frames).sort_index()
        rates = rates[~rates.index.duplicated(keep="last")]
        logger.info(
            "Serie de tasas construida: %d días (%s → %s).",
            len(rates), rates.index.min().date(), rates.index.max().date(),
        )
        return rates["rate_6m"], "TREASURY_GOV_6MO_DAILY"

    @staticmethod
    def _fetch_year(year: int) -> Optional[pd.DataFrame]:
        url = (
            "https://home.treasury.gov/resource-center/data-chart-center/"
            f"interest-rates/daily-treasury-rates.csv/{year}/all"
            f"?type=daily_treasury_yield_curve&field_tdr_date_value={year}"
            "&page&_format=csv"
        )
        try:
            resp = requests.get(url, timeout=20)
            resp.raise_for_status()
            df = pd.read_csv(StringIO(resp.text))
            date_col = df.columns[0]
            target_col = next(
                (c for c in df.columns if "6 Mo" in c or "6MO" in c.upper()),
                None,
            )
            if target_col is None:
                raise ValueError(f"Columna '6 Mo' ausente (año {year})")
            out = pd.DataFrame({
                "date": pd.to_datetime(df[date_col]),
                "rate_6m": pd.to_numeric(df[target_col], errors="coerce") / 100.0,
            }).dropna().set_index("date")
            logger.info("Treasury.gov %d: %d observaciones.", year, len(out))
            return out
        except Exception as exc:
            logger.warning("Fallo Treasury.gov (año %d): %s", year, exc)
            return None


# =========================================================================
# COMPONENTE 2 · ESTIMADOR DE CARRY IMPLÍCITO (idéntico al daemon)
# =========================================================================

class ImpliedCarryEstimator:
    """
    c_neto = ln(F_far_mid / F_near_mid) / (T_far - T_near)

    Misma formulación que el recolector en vivo. Se aplica fila a fila
    sobre los mids históricos de los dos contratos del par de estimación.
    """

    @staticmethod
    def estimate(near_mid: float, near_T: float, far_mid: float, far_T: float) -> float:
        if near_mid <= 0 or far_mid <= 0:
            raise ValueError("Precios mid no positivos.")
        delta_T = far_T - near_T
        if delta_T <= 1e-6:
            raise ValueError("Diferencia de vencimientos insuficiente.")
        return float(np.log(far_mid / near_mid) / delta_T)


# =========================================================================
# COMPONENTE 3 · MOTOR DE ANÁLISIS (idéntico al daemon)
# =========================================================================

class CashCarryAnalyzer:
    """
    Evaluación asimétrica de ambas piernas del arbitraje cruzando el
    spread, neteada contra costes de transacción. Matemática idéntica,
    línea a línea, a la clase homónima del recolector en tiempo real.
    """

    def __init__(self, tx_cost_pct: float):
        self.tx_cost_pct = tx_cost_pct

    def evaluate(
        self,
        spot_bid: float,
        spot_ask: float,
        future_bid: float,
        future_ask: float,
        risk_free_rate: float,
        implied_carry_rate: float,
        time_to_maturity: float,
    ) -> dict:
        total_rate = risk_free_rate + implied_carry_rate

        fwd_theoretical_ask_leg = spot_ask * np.exp(total_rate * time_to_maturity)
        cash_carry_gross = future_bid - fwd_theoretical_ask_leg
        cash_carry_net = cash_carry_gross - (fwd_theoretical_ask_leg * self.tx_cost_pct)

        fwd_theoretical_bid_leg = spot_bid * np.exp(total_rate * time_to_maturity)
        reverse_gross = fwd_theoretical_bid_leg - future_ask
        reverse_net = reverse_gross - (fwd_theoretical_bid_leg * self.tx_cost_pct)

        if cash_carry_net > 0 and cash_carry_net >= reverse_net:
            signal = "CASH_AND_CARRY_ARBITRAGE"
        elif reverse_net > 0:
            signal = "REVERSE_CASH_AND_CARRY_ARBITRAGE"
        else:
            signal = "EQUILIBRIUM_NO_ARBITRAGE"

        return {
            "theoretical_fwd_ask_leg": fwd_theoretical_ask_leg,
            "theoretical_fwd_bid_leg": fwd_theoretical_bid_leg,
            "cash_carry_gross_spread": cash_carry_gross,
            "cash_carry_net_spread": cash_carry_net,
            "reverse_carry_gross_spread": reverse_gross,
            "reverse_carry_net_spread": reverse_net,
            "signal": signal,
        }


# =========================================================================
# COMPONENTE 4 · DATOS HISTÓRICOS DESDE IBKR
# =========================================================================

class IBKRHistoricalDataGateway:
    """
    Conexión a TWS/IB Gateway y descarga de barras históricas diarias.
    Prioriza BID_ASK bars (microestructura real); si el permiso de datos
    lo impide para un contrato, cae a TRADES y lo marca explícitamente.
    """

    def __init__(self, config: BacktestConfig):
        self.ib = IB()
        self.config = config

    def connect(self, max_retries: int = 5, base_delay: float = 2.0) -> None:
        for attempt in range(1, max_retries + 1):
            try:
                self.ib.connect(
                    self.config.host, self.config.port,
                    clientId=self.config.client_id, timeout=30,
                )
                logger.info("Conectado a IBKR (intento %d).", attempt)
                return
            except Exception as exc:
                logger.warning(
                    "Intento de conexión %d/%d fallido: %s",
                    attempt, max_retries, exc,
                )
                if attempt < max_retries:
                    time.sleep(base_delay * attempt)
        raise ConnectionError(
            f"No se pudo conectar a IBKR tras {max_retries} intentos."
        )

    def disconnect(self) -> None:
        if self.ib.isConnected():
            self.ib.disconnect()
            logger.info("Desconectado de IBKR.")

    def fetch_chain(self) -> list:
        """Cadena de futuros viva, ordenada por vencimiento ascendente."""
        generic = Future(symbol=self.config.symbol, exchange=self.config.exchange)
        details = self.ib.reqContractDetails(generic)
        if not details:
            raise ValueError(
                f"No se encontraron contratos para "
                f"{self.config.symbol}/{self.config.exchange}."
            )
        contracts = [
            cd.contract for cd in details
            if cd.contract.lastTradeDateOrContractMonth
        ]
        if len(contracts) < 3:
            raise ValueError(
                f"Se necesitan >= 3 contratos; encontrados {len(contracts)}."
            )
        return sorted(contracts, key=lambda c: c.lastTradeDateOrContractMonth)

    def fetch_daily_bid_ask(self, contract) -> tuple[Optional[pd.DataFrame], str]:
        """
        Devuelve un DataFrame indexado por fecha con columnas [bid, ask]
        y la etiqueta del método usado.

        NIVEL 1: whatToShow='BID_ASK' → open=avg bid, close=avg ask.
        NIVEL 2: whatToShow='TRADES' → close ± half-spread sintético.
        """
        self.ib.qualifyContracts(contract)

        # ── Nivel 1: BID_ASK real ────────────────────────────────
        try:
            bars = self.ib.reqHistoricalData(
                contract,
                endDateTime="",
                durationStr=self.config.duration,
                barSizeSetting=self.config.bar_size,
                whatToShow="BID_ASK",
                useRTH=True,
                formatDate=1,
            )
            if bars:
                df = util.df(bars)[["date", "open", "close"]].rename(
                    columns={"open": "bid", "close": "ask"}
                )
                df["date"] = pd.to_datetime(df["date"])
                df = df.set_index("date").sort_index()
                # Convertir centavos → USD/bushel (ver price_scale en config)
                df["bid"] = df["bid"] / self.config.price_scale
                df["ask"] = df["ask"] / self.config.price_scale
                df = df[(df["bid"] > 0) & (df["ask"] > 0)]
                if not df.empty:
                    logger.info(
                        "%s: %d barras BID_ASK reales.",
                        contract.localSymbol, len(df),
                    )
                    return df, "IBKR_BID_ASK"
        except Exception as exc:
            logger.warning(
                "%s: BID_ASK no disponible (%s). Probando TRADES...",
                contract.localSymbol, exc,
            )

        # ── Nivel 2: TRADES + spread sintético declarado ─────────
        try:
            bars = self.ib.reqHistoricalData(
                contract,
                endDateTime="",
                durationStr=self.config.duration,
                barSizeSetting=self.config.bar_size,
                whatToShow="TRADES",
                useRTH=True,
                formatDate=1,
            )
            if bars:
                df = util.df(bars)[["date", "close"]]
                df["date"] = pd.to_datetime(df["date"])
                df = df.set_index("date").sort_index()
                # Convertir centavos → USD/bushel ANTES de aplicar el
                # half-spread (que ya está expresado en USD/bu).
                df["close"] = df["close"] / self.config.price_scale
                half = self.config.assumed_half_spread_usd
                out = pd.DataFrame({
                    "bid": df["close"] - half,
                    "ask": df["close"] + half,
                })
                out = out[(out["bid"] > 0) & (out["ask"] > 0)]
                if not out.empty:
                    logger.info(
                        "%s: %d barras TRADES + half-spread sintético de "
                        "%.4f USD/bu por lado.",
                        contract.localSymbol, len(out), half,
                    )
                    return out, "SYNTHETIC_SPREAD"
        except Exception as exc:
            logger.error("%s: fallo también en TRADES: %s", contract.localSymbol, exc)

        return None, "NO_DATA"

    @staticmethod
    def time_to_maturity_on(contract, as_of: pd.Timestamp) -> float:
        """T en años desde la fecha de observación hasta el vencimiento."""
        raw = contract.lastTradeDateOrContractMonth
        if len(raw) == 6:
            exp = datetime.strptime(raw + "15", "%Y%m%d")
        else:
            exp = datetime.strptime(raw, "%Y%m%d")
        days = (exp - as_of.to_pydatetime().replace(tzinfo=None)).total_seconds() / 86400.0
        return max(days, 0.0) / 365.0


# =========================================================================
# COMPONENTE 5 · ORQUESTADOR DEL BACKTEST
# =========================================================================

class CashCarryBacktester:
    """
    Flujo por lotes: (1) conectar, (2) resolver cadena, (3) descargar
    históricos de los tres contratos, (4) alinear por fecha, (5) aplicar
    la matemática día a día, (6) consolidar en DataFrame, (7) exportar
    Excel multi-hoja. Con inicio y fin definidos — sin bucle infinito.
    """

    def __init__(self, config: BacktestConfig):
        self.config = config
        self.gateway = IBKRHistoricalDataGateway(config)
        self.analyzer = CashCarryAnalyzer(tx_cost_pct=config.tx_cost_pct)

    # ------------------------------------------------------------------
    def run(self) -> pd.DataFrame:
        self.gateway.connect()
        try:
            chain = self.gateway.fetch_chain()
            near = chain[self.config.idx_near]
            far = chain[self.config.idx_far]
            target = chain[self.config.idx_target]
            logger.info(
                "Contratos del estudio | near=%s far=%s target=%s",
                near.localSymbol, far.localSymbol, target.localSymbol,
            )

            near_px, near_method = self.gateway.fetch_daily_bid_ask(near)
            far_px, far_method = self.gateway.fetch_daily_bid_ask(far)
            tgt_px, tgt_method = self.gateway.fetch_daily_bid_ask(target)

            if any(px is None for px in (near_px, far_px, tgt_px)):
                raise ValueError(
                    "No se pudieron obtener históricos de los tres contratos. "
                    "Verifica permisos de datos históricos en la cuenta."
                )

            df = self._build_dataset(
                near, far, target,
                near_px, far_px, tgt_px,
                near_method, far_method, tgt_method,
            )
            logger.info("Dataset construido: %d observaciones diarias.", len(df))
            return df
        finally:
            self.gateway.disconnect()

    # ------------------------------------------------------------------
    def _build_dataset(
        self, near, far, target,
        near_px, far_px, tgt_px,
        near_method, far_method, tgt_method,
    ) -> pd.DataFrame:
        # 1. Alinear los tres históricos por fecha (inner join: solo días
        #    en que los TRES contratos tienen dato — sin interpolación).
        merged = (
            near_px.rename(columns={"bid": "near_bid", "ask": "near_ask"})
            .join(far_px.rename(columns={"bid": "far_bid", "ask": "far_ask"}), how="inner")
            .join(tgt_px.rename(columns={"bid": "target_bid", "ask": "target_ask"}), how="inner")
        )
        if merged.empty:
            raise ValueError("Sin fechas comunes entre los tres contratos.")

        # 2. Serie histórica de tasas — asignación sin look-ahead.
        start_year = merged.index.min().year
        rate_series, rate_source = HistoricalRiskFreeRateProvider.build_series(start_year)
        if rate_series.empty:
            merged["risk_free_rate"] = HistoricalRiskFreeRateProvider.DEFAULT_RATE
        else:
            rates_df = rate_series.rename("risk_free_rate").reset_index()
            rates_df.columns = ["date", "risk_free_rate"]
            obs_df = merged.reset_index().rename(columns={"index": "date"})
            # Normalizar AMBOS lados a la misma resolución de datetime64.
            # Las barras de IBKR llegan en [s] y el CSV de Treasury en
            # [us]/[ns] según la versión de pandas; merge_asof exige que
            # coincidan exactamente para poder comparar las claves.
            obs_df["date"] = pd.to_datetime(obs_df["date"]).astype("datetime64[ns]")
            rates_df["date"] = pd.to_datetime(rates_df["date"]).astype("datetime64[ns]")
            merged = pd.merge_asof(
                obs_df.sort_values("date"),
                rates_df.sort_values("date"),
                on="date",
                direction="backward",
            ).set_index("date")
            merged["risk_free_rate"] = merged["risk_free_rate"].fillna(
                HistoricalRiskFreeRateProvider.DEFAULT_RATE
            )

        # 3. Aplicar la matemática fila a fila.
        records: list[dict] = []
        for date, row in merged.iterrows():
            try:
                t_near = self.gateway.time_to_maturity_on(near, date)
                t_far = self.gateway.time_to_maturity_on(far, date)
                t_tgt = self.gateway.time_to_maturity_on(target, date)

                near_mid = (row["near_bid"] + row["near_ask"]) / 2.0
                far_mid = (row["far_bid"] + row["far_ask"]) / 2.0

                implied_carry = ImpliedCarryEstimator.estimate(
                    near_mid, t_near, far_mid, t_far
                )
                result = self.analyzer.evaluate(
                    spot_bid=row["near_bid"],
                    spot_ask=row["near_ask"],
                    future_bid=row["target_bid"],
                    future_ask=row["target_ask"],
                    risk_free_rate=row["risk_free_rate"],
                    implied_carry_rate=implied_carry,
                    time_to_maturity=t_tgt,
                )
                records.append({
                    "date": date,
                    "spot_proxy_symbol": near.localSymbol,
                    "near_future_symbol": near.localSymbol,
                    "far_future_symbol": far.localSymbol,
                    "target_future_symbol": target.localSymbol,
                    "spot_bid": row["near_bid"],
                    "spot_ask": row["near_ask"],
                    "target_future_bid": row["target_bid"],
                    "target_future_ask": row["target_ask"],
                    "risk_free_rate": row["risk_free_rate"],
                    "risk_free_rate_source": rate_source,
                    "implied_carry_rate": implied_carry,
                    "carry_estimation_pair": f"{near.localSymbol}/{far.localSymbol}",
                    "time_to_maturity_years": t_tgt,
                    "spread_method_near": near_method,
                    "spread_method_far": far_method,
                    "spread_method_target": tgt_method,
                    "assumed_half_spread_usd": self.config.assumed_half_spread_usd,
                    "transaction_cost_pct": self.config.tx_cost_pct,
                    **result,
                })
            except ValueError as exc:
                logger.warning("Fila %s descartada: %s", date.date(), exc)

        return pd.DataFrame.from_records(records).set_index("date")

    # ------------------------------------------------------------------
    def export_excel(self, df: pd.DataFrame) -> str:
        """
        Exporta el dataset a Excel con dos hojas:
          - 'Raw Data'      : todas las observaciones.
          - 'Summary Stats' : estadística descriptiva y métricas de señal.
        """
        summary = self._build_summary(df)
        with pd.ExcelWriter(self.config.output_xlsx, engine="openpyxl") as writer:
            df.reset_index().to_excel(writer, sheet_name="Raw Data", index=False)
            summary.to_excel(writer, sheet_name="Summary Stats", index=False)
        logger.info("Excel exportado: %s", self.config.output_xlsx)
        return self.config.output_xlsx

    # ------------------------------------------------------------------
    def _build_summary(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Resumen estadístico del estudio: distribución de los spreads
        netos, frecuencia de señales y beneficio teórico agregado.
        """
        n = len(df)
        best_net = df[["cash_carry_net_spread", "reverse_carry_net_spread"]].max(axis=1)

        signal_counts = df["signal"].value_counts()
        pct_contango = signal_counts.get("CASH_AND_CARRY_ARBITRAGE", 0) / n * 100 if n else 0.0
        pct_backward = signal_counts.get("REVERSE_CASH_AND_CARRY_ARBITRAGE", 0) / n * 100 if n else 0.0
        pct_equilib = signal_counts.get("EQUILIBRIUM_NO_ARBITRAGE", 0) / n * 100 if n else 0.0

        # Beneficio teórico: suma de spreads netos POSITIVOS (solo días
        # con oportunidad ejecutable), expresado por bushel y por
        # contrato estándar de 5,000 bushels.
        exploitable = best_net[best_net > 0]
        total_profit_bu = exploitable.sum()
        contract_size = 5_000

        rows = [
            ("Observaciones totales", n),
            ("Periodo inicio", df.index.min().date() if n else "—"),
            ("Periodo fin", df.index.max().date() if n else "—"),
            ("", ""),
            ("Media spread neto C&C (USD/bu)", round(df["cash_carry_net_spread"].mean(), 4)),
            ("Desv. estándar spread neto C&C", round(df["cash_carry_net_spread"].std(), 4)),
            ("Media spread neto Reverse (USD/bu)", round(df["reverse_carry_net_spread"].mean(), 4)),
            ("Desv. estándar spread neto Reverse", round(df["reverse_carry_net_spread"].std(), 4)),
            ("Media carry implícito anualizado", f"{df['implied_carry_rate'].mean() * 100:.2f}%"),
            ("Desv. estándar carry implícito", f"{df['implied_carry_rate'].std() * 100:.2f}%"),
            ("", ""),
            ("% señales Cash & Carry (contango)", f"{pct_contango:.1f}%"),
            ("% señales Reverse C&C (backwardation)", f"{pct_backward:.1f}%"),
            ("% señales Equilibrio", f"{pct_equilib:.1f}%"),
            ("", ""),
            ("Días con oportunidad ejecutable (net > 0)", int((best_net > 0).sum())),
            ("Beneficio teórico total (USD/bushel)", round(total_profit_bu, 4)),
            (f"Beneficio teórico total (USD/contrato de {contract_size:,} bu)",
             round(total_profit_bu * contract_size, 2)),
            ("", ""),
            ("Coste transacción asumido", f"{self.config.tx_cost_pct * 100:.2f}%"),
            ("Half-spread sintético (si aplica, USD/bu)",
             self.config.assumed_half_spread_usd),
            ("Métodos de spread usados",
             ", ".join(sorted(set(
                 df["spread_method_near"].iloc[0:1].tolist()
                 + df["spread_method_far"].iloc[0:1].tolist()
                 + df["spread_method_target"].iloc[0:1].tolist()
             ))) if n else "—"),
        ]
        return pd.DataFrame(rows, columns=["Métrica", "Valor"])


# =========================================================================
# PUNTO DE ENTRADA
# =========================================================================

if __name__ == "__main__":
    config = BacktestConfig(
        symbol="ZS",
        duration="6 M",              # ventana del estudio
        tx_cost_pct=0.001,
        assumed_half_spread_ticks=2, # fallback si no hay BID_ASK bars
        port=7496,
    )

    backtester = CashCarryBacktester(config)
    dataset = backtester.run()

    if dataset.empty:
        logger.error("Dataset vacío: no se exporta Excel.")
    else:
        path = backtester.export_excel(dataset)
        print(f"\n{'=' * 60}")
        print("  BACKTEST COMPLETADO")
        print(f"{'=' * 60}")
        print(f"  Observaciones : {len(dataset)}")
        print(f"  Archivo Excel : {path}")
        print(f"{'=' * 60}\n")
