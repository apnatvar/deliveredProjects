"""

- user inputs:
  Symbol, Timeframe (daily/intraday), CSV (optional), Start, End, Cash, Commission
- Has a Strategy Prompt box
- Has a Strategy Code box
- Displays Metrics and Trades and a final interactive chart (plotted via backtest.py)
- If CSV is provided, it is loaded to a DataFrame; else it tries `backtesting.test.<SYMBOL>`
  (e.g., GOOG). If not found, falls back to GOOG.

TODO :
- currently only working with GOOG, trying to get a custom ticker working
- add a code validator to perform a "pre-check" if the code will run or fail
- backtesting.py is the backend, so there are more options to add such as optimisation for a given metric
- other features as I learn and explore the backtesting.py libary
"""

from __future__ import annotations

import sys
import types
import traceback
import pandas as pd
from backtesting import Backtest
from backtesting import Strategy as BTStrategy
import backtesting.test
from dataclasses import dataclass
from typing import Optional, Tuple

from PyQt6.QtCore import (
    Qt,
    QDate,
    QSettings,
    QThread,
    pyqtSignal,
    QRegularExpression,
)
from PyQt6.QtGui import (
    QAction,
    QColor,
    QFont,
    QSyntaxHighlighter,
    QTextCharFormat,
    QTextOption,
)
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSplitter,
    QStatusBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QToolBar,
    QVBoxLayout,
    QWidget,
    QDateEdit,
    QComboBox,
)


# --------------------------- Config model --------------------------- #
@dataclass
class RunConfig:
    symbol: str
    timeframe: str
    csv_path: Optional[str]
    start_date_iso: str
    end_date_iso: str
    cash: float
    commission: float
    strategy_code: str


# --------------------------- Strategy loader ------------------------ #


# TODO add the libraries so that they are imported implicityly and need not be mentioned in the strategy dialogbox
class RunError(Exception):
    pass


def load_strategy_from_source(source: str):
    """
    Execute user-supplied code and return the Strategy class.
    The code must define: class Strategy(backtesting.Strategy): ...
    """
    module = types.ModuleType("user_strategy")
    try:
        exec(source, module.__dict__)  # noqa: S102 (intentional: local execution)
    except Exception as exc:  # bubble error with traceback
        raise RunError(f"Strategy code error:\n{traceback.format_exc()}") from exc

    Strategy = getattr(module, "Strategy", None)
    if Strategy is None or not isinstance(Strategy, type):
        raise RunError(
            "No class named 'Strategy' found.\n"
            "Please define: `class Strategy(backtesting.Strategy): ...`"
        )
    # Optional: sanity check that it subclasses BTStrategy
    if not issubclass(Strategy, BTStrategy):
        raise RunError("`Strategy` must subclass backtesting.Strategy.")
    return Strategy


# --------------------------- Data loading --------------------------- #
def load_data_frame(cfg: RunConfig) -> pd.DataFrame:
    """
    Load OHLCV dataframe with DateTimeIndex and columns: Open, High, Low, Close, Volume.
    - If csv_path is provided, read CSV and normalize.
    - Else, attempt to load from backtesting.test by symbol (uppercased).
    - Finally, fallback to bt_test.GOOG.
    """
    print(f"{cfg.csv_path=}")
    if cfg.csv_path:
        df = pd.read_csv(cfg.csv_path)
        # Identify datetime column
        for dt_col in ("Date", "Datetime", "date", "timestamp", "Timestamp"):
            if dt_col in df.columns:
                df[dt_col] = pd.to_datetime(df[dt_col])
                df = df.set_index(dt_col)
                break
        if not isinstance(df.index, pd.DatetimeIndex):
            raise RunError(
                "CSV must include a Date/Datetime column to use as the index."
            )

        # Normalize columns to expected names
        colmap_lower = {c.lower(): c for c in df.columns}

        def pick(*names):
            for n in names:
                if n in df.columns:
                    return n
                if n.lower() in colmap_lower:
                    return colmap_lower[n.lower()]
            return None

        req = {
            "Open": pick("Open"),
            "High": pick("High"),
            "Low": pick("Low"),
            "Close": pick("Close", "Adj Close", "AdjClose"),
            "Volume": pick("Volume", "Vol"),
        }
        missing = [k for k, v in req.items() if v is None]
        if missing:
            raise RunError(f"CSV missing required columns: {', '.join(missing)}")

        df = df[
            [req["Open"], req["High"], req["Low"], req["Close"], req["Volume"]]
        ].copy()
        df.columns = ["Open", "High", "Low", "Close", "Volume"]

    else:
        # Try dynamic dataset from backtesting.test
        sym = (cfg.symbol or "GOOG").strip()
        df = getattr(backtesting.test, sym)
        print(f"{df=}")
        return df

    # Date filter (if possible)
    try:
        df = df.loc[cfg.start_date_iso : cfg.end_date_iso]
    except Exception:
        pass
    try:
        # Ensure required columns exist
        expected_cols = {"Open", "High", "Low", "Close", "Volume"}
        if expected_cols.issubset(set(df.columns)):
            return df
    except Exception:
        print("Error generating test dataframe")
    return None


# --------------------------- Backtest runner ------------------------ #
def runBackTest(cfg: RunConfig) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Run a backtest with Backtesting.py.
    Returns:
      metrics_df: a 2-col DataFrame [Metric, Value]
      trades_df:  a DataFrame built from output._trades (if available)
    """
    # Compile strategy
    Strategy = load_strategy_from_source(cfg.strategy_code)

    # Load data
    data = load_data_frame(cfg)
    print(f"{data=}")
    # TODO if data is empty switch to using bt_test
    # Backtesting
    # print(backtesting.test.GOOG)
    bt = Backtest(
        data,
        # backtesting.test.GOOG,
        Strategy,
        cash=cfg.cash,
        commission=cfg.commission,  # e.g., 0.001 means 0.1%
        exclusive_orders=True,
        # finalize_trades=True,
    )

    output = bt.run()
    bt.plot(open_browser=True)

    # Build metrics DataFrame for display
    metrics_items = []
    for key in [
        "Equity Final [$]",
        "Equity Peak [$]",
        "Return [%]",
        "Buy & Hold Return [%]",
        "# Trades",
        "Win Rate [%]",
        "Best Trade [%]",
        "Worst Trade [%]",
        "Avg. Trade [%]",
        "Max. Drawdown [%]",
        "Sharpe Ratio",
        "Sortino Ratio",
    ]:
        if key in output:
            metrics_items.append((key, output[key]))
    metrics_df = pd.DataFrame(metrics_items, columns=["Metric", "Value"])

    # Trades DataFrame (if present)
    if hasattr(output, "_trades") and isinstance(output._trades, pd.DataFrame):
        trades_df = output._trades.copy()
    else:
        trades_df = pd.DataFrame(
            columns=["EntryTime", "ExitTime", "EntryPrice", "ExitPrice", "PnL"]
        )

    return metrics_df, trades_df


# --------------------------- UI helpers ----------------------------- #
class PythonHighlighter(QSyntaxHighlighter):
    KEYWORDS = [
        "and",
        "as",
        "assert",
        "break",
        "class",
        "continue",
        "def",
        "del",
        "elif",
        "else",
        "except",
        "False",
        "finally",
        "for",
        "from",
        "global",
        "if",
        "import",
        "in",
        "is",
        "lambda",
        "None",
        "nonlocal",
        "not",
        "or",
        "pass",
        "raise",
        "return",
        "True",
        "try",
        "while",
        "with",
        "yield",
    ]

    def __init__(self, document):
        super().__init__(document)
        self._build_rules()

    def _build_rules(self):
        def fmt(rgb, weight=QFont.Weight.Normal):
            f = QTextCharFormat()
            f.setForeground(QColor(*rgb))
            f.setFontWeight(weight)
            return f

        self.rules = []
        kw_fmt = fmt((197, 134, 192), QFont.Weight.DemiBold)
        for kw in self.KEYWORDS:
            self.rules.append((QRegularExpression(rf"\b{kw}\b"), kw_fmt))

        self.rules.append(
            (QRegularExpression(r"#[^\n]*"), fmt((106, 153, 85)))
        )  # comments
        self.rules.append(
            (QRegularExpression(r"\bdef\s+\w+"), fmt((220, 220, 170)))
        )  # def
        self.rules.append(
            (QRegularExpression(r"\bclass\s+\w+"), fmt((220, 220, 170)))
        )  # class
        self.rules.append(
            (QRegularExpression(r"\bself\b"), fmt((86, 156, 214)))
        )  # self
        self.rules.append(
            (QRegularExpression(r"\b\d+(\.\d+)?\b"), fmt((181, 206, 168)))
        )  # numbers
        self.rules.append(
            (QRegularExpression(r"\"[^\"]*\"|'[^']*'"), fmt((206, 145, 120)))
        )  # strings

    def highlightBlock(self, text: str) -> None:
        for regex, qfmt in self.rules:
            it = regex.globalMatch(text)
            while it.hasNext():
                m = it.next()
                self.setFormat(m.capturedStart(), m.capturedLength(), qfmt)


def df_to_table(table: QTableWidget, df: pd.DataFrame) -> None:
    table.clear()
    # Set headers
    table.setColumnCount(len(df.columns))
    table.setRowCount(len(df))
    table.setHorizontalHeaderLabels([str(c) for c in df.columns])
    # Fill
    for r in range(len(df)):
        for c in range(len(df.columns)):
            val = df.iat[r, c]
            item = QTableWidgetItem("" if pd.isna(val) else str(val))
            table.setItem(r, c, item)
    table.resizeColumnsToContents()
    table.horizontalHeader().setStretchLastSection(True)


# --------------------------- Worker thread -------------------------- #
class RunWorker(QThread):
    done = pyqtSignal(pd.DataFrame, pd.DataFrame)  # metrics_df, trades_df
    errored = pyqtSignal(str)

    def __init__(self, cfg: RunConfig):
        super().__init__()
        self.cfg = cfg

    def run(self):
        try:
            metrics_df, trades_df = runBackTest(self.cfg)
            self.done.emit(metrics_df, trades_df)
        except Exception as exc:
            self.errored.emit(str(exc))


# --------------------------- Main Window ---------------------------- #
class MainWindow(QMainWindow):
    ORG = "PasTick"
    APP = "BacktestingWorkbench"

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Backtesting Workbench — PyQt6")
        self.resize(1260, 840)

        self.settings = QSettings(self.ORG, self.APP)

        self._build_menu_toolbar()
        self._build_statusbar()
        self._build_central()
        self._load_state()

    # ----- UI construction -----
    def _build_menu_toolbar(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu("&File")
        self.act_quit = QAction("Quit", self)
        file_menu.addAction(self.act_quit)
        self.act_quit.triggered.connect(self.close)

        tb = QToolBar("Main")
        tb.setMovable(False)
        self.addToolBar(tb)

        self.act_run = QAction("Run Backtest", self)
        self.act_run.setShortcut("F5")
        tb.addAction(self.act_run)
        self.act_run.triggered.connect(self._on_run)

        self.act_browse_csv = QAction("Choose CSV…", self)
        tb.addAction(self.act_browse_csv)
        self.act_browse_csv.triggered.connect(self._browse_csv)

    def _build_statusbar(self):
        sb = QStatusBar(self)
        self.setStatusBar(sb)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.statusBar().addPermanentWidget(self.progress, 0)

    def _build_central(self):
        root = QSplitter(Qt.Orientation.Horizontal, self)
        self.setCentralWidget(root)

        # Left: form
        left = QWidget()
        left_lay = QVBoxLayout(left)
        left_lay.setContentsMargins(10, 10, 10, 10)
        left_lay.setSpacing(8)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.inp_symbol = QLineEdit()
        self.cmb_timeframe = QComboBox()
        self.cmb_timeframe.addItems(["Daily", "Intraday"])

        # Data Source: CSV only with browse
        self.cmb_datasource = QComboBox()
        self.cmb_datasource.addItems(["CSV (local)"])
        self.inp_csv = QLineEdit()
        self.btn_csv = QPushButton("Browse…")
        self.btn_csv.clicked.connect(self._browse_csv)
        csv_row = QWidget()
        csv_lay = QHBoxLayout(csv_row)
        csv_lay.setContentsMargins(0, 0, 0, 0)
        csv_lay.addWidget(self.inp_csv, 1)
        csv_lay.addWidget(self.btn_csv)

        self.date_start = QDateEdit()
        self.date_start.setCalendarPopup(True)
        self.date_end = QDateEdit()
        self.date_end.setCalendarPopup(True)

        self.inp_cash = QLineEdit()
        self.inp_commission = QLineEdit()

        form.addRow("Symbol:", self.inp_symbol)
        form.addRow("Timeframe:", self.cmb_timeframe)
        form.addRow("Data Source:", self.cmb_datasource)
        form.addRow("CSV Path:", csv_row)
        form.addRow("Start:", self.date_start)
        form.addRow("End:", self.date_end)
        form.addRow("Cash:", self.inp_cash)
        form.addRow("Commission (fraction):", self.inp_commission)

        left_lay.addLayout(form)

        # Strategy prompt (text only; NOT passed to runBackTest)
        left_lay.addWidget(QLabel("Strategy Prompt (for your API):"))
        self.prompt_edit = QPlainTextEdit()
        self.prompt_edit.setPlaceholderText(
            "Describe the strategy in plain English… (not sent to runBackTest)"
        )
        left_lay.addWidget(self.prompt_edit, 1)

        # Run button
        self.btn_run = QPushButton("Run Backtest")
        self.btn_run.clicked.connect(self._on_run)
        left_lay.addWidget(self.btn_run)

        root.addWidget(left)

        # Right: code + results tabs
        right = QSplitter(Qt.Orientation.Vertical)
        root.addWidget(right)
        root.setStretchFactor(0, 0)
        root.setStretchFactor(1, 1)

        # Strategy code
        code_wrap = QWidget()
        code_lay = QVBoxLayout(code_wrap)
        code_lay.setContentsMargins(8, 8, 8, 8)
        code_lay.addWidget(
            QLabel(
                "Strategy Code (must define class `Strategy(backtesting.Strategy)`):"
            )
        )

        self.code_edit = QPlainTextEdit()
        self.code_edit.setWordWrapMode(QTextOption.WrapMode.NoWrap)
        self.code_edit.setFont(QFont("Consolas", 11))
        PythonHighlighter(self.code_edit.document())
        self.code_edit.setPlainText(DEFAULT_STRATEGY_CODE.strip())
        code_lay.addWidget(self.code_edit, 1)

        right.addWidget(code_wrap)

        # Tabs: Metrics, Trades, Logs
        self.tabs = QTabWidget()
        self.tbl_metrics = QTableWidget(0, 2)
        self.tbl_metrics.setHorizontalHeaderLabels(["Metric", "Value"])
        self.tbl_metrics.horizontalHeader().setStretchLastSection(True)

        self.tbl_trades = QTableWidget(0, 5)
        self.tbl_trades.setHorizontalHeaderLabels(
            ["EntryTime", "ExitTime", "EntryPrice", "ExitPrice", "PnL"]
        )
        self.tbl_trades.horizontalHeader().setStretchLastSection(True)

        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setFont(QFont("Consolas", 10))

        tab_metrics = QWidget()
        lay_m = QVBoxLayout(tab_metrics)
        lay_m.addWidget(self.tbl_metrics)

        tab_trades = QWidget()
        lay_t = QVBoxLayout(tab_trades)
        lay_t.addWidget(self.tbl_trades)

        self.tabs.addTab(tab_metrics, "Metrics")
        self.tabs.addTab(tab_trades, "Trades")
        self.tabs.addTab(self.txt_log, "Logs")

        right.addWidget(self.tabs)

    # ----- state load/save -----
    def _load_state(self):
        s = self.settings
        self.inp_symbol.setText(s.value("symbol", "GOOG"))
        self.cmb_timeframe.setCurrentText(s.value("timeframe", "Daily"))
        self.inp_csv.setText(s.value("csv_path", ""))

        self.date_start.setDate(
            QDate.fromString(
                s.value(
                    "start", QDate.currentDate().addYears(-1).toString("yyyy-MM-dd")
                ),
                "yyyy-MM-dd",
            )
        )
        self.date_end.setDate(
            QDate.fromString(
                s.value("end", QDate.currentDate().toString("yyyy-MM-dd")), "yyyy-MM-dd"
            )
        )

        self.inp_cash.setText(s.value("cash", "100000"))
        self.inp_commission.setText(s.value("commission", "0.0005"))  # 5 bps default

    def _save_state(self):
        s = self.settings
        s.setValue("symbol", self.inp_symbol.text().strip())
        s.setValue("timeframe", self.cmb_timeframe.currentText())
        s.setValue("csv_path", self.inp_csv.text().strip())
        s.setValue("start", self.date_start.date().toString("yyyy-MM-dd"))
        s.setValue("end", self.date_end.date().toString("yyyy-MM-dd"))
        s.setValue("cash", self.inp_cash.text().strip())
        s.setValue("commission", self.inp_commission.text().strip())

    # ----- handlers -----
    def _browse_csv(self):
        fn, _ = QFileDialog.getOpenFileName(
            self, "Select OHLCV CSV", "", "CSV Files (*.csv);"
        )
        if fn:
            self.inp_csv.setText(fn)

    def _on_run(self):
        try:
            cfg = RunConfig(
                symbol=self.inp_symbol.text().strip() or "GOOG",
                timeframe=self.cmb_timeframe.currentText(),
                csv_path=(self.inp_csv.text().strip() or None),
                start_date_iso=self.date_start.date().toString("yyyy-MM-dd"),
                end_date_iso=self.date_end.date().toString("yyyy-MM-dd"),
                cash=float(self.inp_cash.text().strip() or "100000"),
                commission=float(self.inp_commission.text().strip() or "0.0"),
                strategy_code=self.code_edit.toPlainText(),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Invalid Input", str(exc))
            return

        self._save_state()
        self.txt_log.clear()
        self.progress.setRange(0, 0)  # indeterminate

        self.worker = RunWorker(cfg)
        self.worker.done.connect(self._on_done)
        self.worker.errored.connect(self._on_error)
        self.worker.start()

    def _on_done(self, metrics_df: pd.DataFrame, trades_df: pd.DataFrame):
        self.progress.setRange(0, 100)
        self.progress.setValue(100)
        df_to_table(self.tbl_metrics, metrics_df)
        df_to_table(self.tbl_trades, trades_df)
        self.tabs.setCurrentIndex(0)

    def _on_error(self, message: str):
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.txt_log.setPlainText(message)
        self.tabs.setCurrentIndex(2)

    # ensure settings persist
    def closeEvent(self, event):
        try:
            self._save_state()
        finally:
            super().closeEvent(event)


# --------------------------- Defaults ------------------------------- #
DEFAULT_STRATEGY_CODE = r"""
# Example Strategy
# IMPORTANT: Keep the class name exactly `Strategy` and subclass backtesting.Strategy

from backtesting import Strategy
from backtesting.lib import crossover
from backtesting.test import SMA

class Strategy(Strategy):
    def init(self):
        self.sma_fast = self.I(SMA, self.data.Close, 10)
        self.sma_slow = self.I(SMA, self.data.Close, 20)

    def next(self):
        if crossover(self.sma_fast, self.sma_slow):
            self.position.close()
            self.buy()
        elif crossover(self.sma_slow, self.sma_fast):
            self.position.close()
            self.sell()
""".strip()


# --------------------------- Entrypoint ----------------------------- #
def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Backtesting Workbench")
    app.setOrganizationName("PasTick")
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
