"""Build the valuation workbook.

Only raw data (EDGAR values, prices, regression output) and assumptions are
written as values. Everything downstream is a live Excel formula, so changing
any blue input on the Assumptions sheet flows through to the Summary.

Key scalars are Excel *named ranges* (WACC, TerminalGrowth, SharePrice, ...)
so formulas read like the finance: =G16*(1+TerminalGrowth)/(WACC-TerminalGrowth).
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from edgar.extract import Financials
from model.assumptions import Assumptions
from model.beta import BetaResult
from excel.styles import (
    BLUE, FMT_FACTOR, FMT_INT, FMT_MM, FMT_MULT, FMT_NUM, FMT_NUM4, FMT_PCT, FMT_PER_SHARE, FMT_SHARES_MM,
    GREY, KEY_FILL, THIN_TOP, calc, header_row, inp, label, link, section, title, widths,
)

SHEETS = ["Summary", "Historicals", "Assumptions", "DCF", "Beta Regression"]
RAW_ROWS = ["Revenue", "Operating Income", "Net Income", "D&A", "Operating Cash Flow", "CapEx",
            "Pretax Income", "Income Tax", "Cash", "Debt"]
PROJECTION_YEARS = 5


@dataclass
class ModelData:
    fin: Financials
    assumptions: Assumptions
    beta: BetaResult
    price: float
    price_date: str
    as_of: date


def _if_both(a: str, b: str, expr: str) -> str:
    """Formula that returns "" unless both referenced cells hold numbers (missing data stays blank)."""
    return f'=IF(COUNT({a},{b})=2,{expr},"")'


class ValuationWorkbook:
    def __init__(self, data: ModelData):
        self.d = data
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.ws = {name: self.wb.create_sheet(name) for name in SHEETS}
        self.hist_rows: dict[str, int] = {}
        self.hist_latest_col = ""

    # ------------------------------------------------------------------ helpers
    def define(self, name: str, sheet: str, ref: str) -> None:
        """Create a workbook-level named range pointing at a single cell."""
        col = "".join(ch for ch in ref if ch.isalpha())
        row = "".join(ch for ch in ref if ch.isdigit())
        self.wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!${col}${row}")

    def hist(self, item: str, col: str | None = None) -> str:
        """Cross-sheet reference to a Historicals cell (latest fiscal year by default)."""
        return f"Historicals!{col or self.hist_latest_col}{self.hist_rows[item]}"

    def build(self, path: Path | str) -> Path:
        # Order matters: later sheets reference rows recorded by earlier ones.
        self._historicals()
        self._assumptions()
        self._beta()
        self._dcf()
        self._summary()
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(path)
        return path

    # ------------------------------------------------------------------ Historicals
    def _historicals(self):
        ws, fin = self.ws["Historicals"], self.d.fin
        years = list(fin.table.columns)
        cols = [get_column_letter(2 + i) for i in range(len(years))]
        src_col = get_column_letter(2 + len(years))
        self.hist_latest_col = cols[-1]

        title(ws, "A1", f"{fin.entity_name} – Historical Financials",
              f"$ millions · fiscal years ending ~{fin.year_ends[years[-1]]:%B %d} · "
              "source: SEC EDGAR XBRL companyfacts (10-K, FY)")
        header_row(ws, 4, 1, ["($ millions)"] + [f"FY{y}" for y in years] + ["Source XBRL tag (us-gaap)"])

        r = 5
        for item in RAW_ROWS:
            label(ws, f"A{r}", item)
            for c, y in zip(cols, years):
                v = fin.table.loc[item, y]
                if pd.notna(v):              # missing stays a truly empty cell (see Data Notes)
                    inp(ws, f"{c}{r}", float(v), FMT_MM)
            label(ws, f"{src_col}{r}", fin.sources.get(item, ""), italic=True, color=GREY)
            self.hist_rows[item] = r
            r += 1

        r += 1
        section(ws, r, 1, len(years) + 2, "Derived metrics (formulas)")
        r += 1
        H = self.hist_rows

        def derived(name: str, fmt: str, make, first_col: int = 0):
            nonlocal r
            label(ws, f"A{r}", name)
            for i, c in enumerate(cols):
                if i >= first_col:
                    calc(ws, f"{c}{r}", make(i, c), fmt)
            H[name] = r
            r += 1

        def ratio(num: str, den: str):
            return lambda i, c: _if_both(f"{c}{H[num]}", f"{c}{H[den]}", f'IFERROR({c}{H[num]}/{c}{H[den]},"")')

        derived("EBITDA", FMT_MM, lambda i, c: _if_both(f"{c}{H['Operating Income']}", f"{c}{H['D&A']}",
                                                         f"{c}{H['Operating Income']}+{c}{H['D&A']}"))
        derived("Free Cash Flow", FMT_MM, lambda i, c: _if_both(f"{c}{H['Operating Cash Flow']}", f"{c}{H['CapEx']}",
                                                                 f"{c}{H['Operating Cash Flow']}-{c}{H['CapEx']}"))
        derived("Revenue Growth %", FMT_PCT, lambda i, c: _if_both(
            f"{cols[i - 1]}{H['Revenue']}", f"{c}{H['Revenue']}",
            f'IFERROR({c}{H["Revenue"]}/{cols[i - 1]}{H["Revenue"]}-1,"")'), first_col=1)
        derived("EBITDA Margin %", FMT_PCT, ratio("EBITDA", "Revenue"))
        derived("FCF Margin %", FMT_PCT, ratio("Free Cash Flow", "Revenue"))
        derived("D&A % of Revenue", FMT_PCT, ratio("D&A", "Revenue"))
        derived("CapEx % of Revenue", FMT_PCT, ratio("CapEx", "Revenue"))
        derived("Effective Tax Rate %", FMT_PCT, ratio("Income Tax", "Pretax Income"))

        widths(ws, A=24, **{c: 12 for c in cols}, **{src_col: 70})
        ws.freeze_panes = "B5"

    # ------------------------------------------------------------------ Assumptions
    def _assumptions(self):
        ws, a, fin = self.ws["Assumptions"], self.d.assumptions, self.d.fin
        notes_col = get_column_letter(2 + PROJECTION_YEARS)
        title(ws, "A1", "Assumptions",
              "Blue = input (edit these) · Black = formula · Green = link to another sheet")
        header_row(ws, 4, 1, ["Operating assumptions"] + [f"Year {k}" for k in range(1, PROJECTION_YEARS + 1)]
                   + ["Basis for default"])

        label(ws, "A5", "Revenue growth %")
        for k, g in enumerate(a.revenue_growth):
            inp(ws, f"{get_column_letter(2 + k)}5", g, FMT_PCT)
        label(ws, f"{notes_col}5", a.basis["revenue_growth"], italic=True, color=GREY)
        self.growth_row = 5

        def scalar(row: int, text: str, value, fmt: str, name: str, basis: str, is_link: bool = False):
            label(ws, f"A{row}", text)
            (link if is_link else inp)(ws, f"B{row}", value, fmt)
            label(ws, f"{notes_col}{row}", basis, italic=True, color=GREY)
            self.define(name, "Assumptions", f"B{row}")

        scalar(6, "EBITDA margin %", a.ebitda_margin, FMT_PCT, "EBITDAMargin", a.basis["ebitda_margin"])
        scalar(7, "D&A % of revenue", a.da_pct, FMT_PCT, "DAPct", a.basis["da_pct"])
        scalar(8, "CapEx % of revenue", a.capex_pct, FMT_PCT, "CapExPct", a.basis["capex_pct"])
        scalar(9, "Tax rate", a.tax_rate, FMT_PCT, "TaxRate", a.basis["tax_rate"])

        section(ws, 11, 1, 2 + PROJECTION_YEARS, "Cost of capital")
        scalar(12, "Risk-free rate", a.risk_free, FMT_PCT, "RiskFree", a.basis["risk_free"])
        scalar(13, "Equity risk premium", a.equity_risk_premium, FMT_PCT, "EquityRiskPremium",
               a.basis["equity_risk_premium"])
        scalar(14, "Beta", a.beta, FMT_NUM, "EquityBeta", a.basis["beta"])
        scalar(15, "Pre-tax cost of debt", a.cost_of_debt, FMT_PCT, "CostOfDebt", a.basis["cost_of_debt"])
        scalar(16, "Debt materiality threshold", a.debt_materiality, FMT_PCT, "DebtThreshold",
               a.basis["debt_materiality"])
        scalar(17, "Terminal growth rate", a.terminal_growth, FMT_PCT, "TerminalGrowth", a.basis["terminal_growth"])

        section(ws, 19, 1, 2 + PROJECTION_YEARS, "Market data & balance sheet")
        scalar(20, "Current share price", self.d.price, FMT_PER_SHARE, "SharePrice",
               f"yfinance close, {self.d.price_date}")
        scalar(21, "Shares outstanding (millions)", fin.shares_outstanding, FMT_SHARES_MM, "SharesOut",
               f"{fin.sources.get('Shares Outstanding', '')}, as of {fin.shares_as_of}")
        scalar(22, "Debt ($ millions)", f"={self.hist('Debt')}", FMT_MM, "BookDebt",
               f"Latest fiscal year (FY{fin.table.columns[-1]}) balance sheet, from Historicals", is_link=True)
        scalar(23, "Cash ($ millions)", f"={self.hist('Cash')}", FMT_MM, "Cash",
               f"Latest fiscal year (FY{fin.table.columns[-1]}) balance sheet, from Historicals", is_link=True)

        widths(ws, A=30, **{get_column_letter(2 + k): 11 for k in range(PROJECTION_YEARS)}, **{notes_col: 80})
        ws.freeze_panes = "A5"

    # ------------------------------------------------------------------ DCF
    def _dcf(self):
        ws, fin = self.ws["DCF"], self.d.fin
        base_fy = int(fin.table.columns[-1])
        cols = [get_column_letter(2 + k) for k in range(PROJECTION_YEARS + 1)]   # B = base year, C..G = Years 1-5
        proj = cols[1:]
        first, last = proj[0], proj[-1]

        title(ws, "A1", "Discounted Cash Flow (DCF)",
              "$ millions except per share · unlevered free cash flow · end-of-year discounting")
        header_row(ws, 4, 1, ["Fiscal year", f"FY{base_fy}A"] + [f"FY{base_fy + k}E" for k in range(1, 6)])

        label(ws, "A5", "Period (t)", italic=True)
        calc(ws, "B5", 0, FMT_INT)
        for prev, c in zip(cols, proj):
            calc(ws, f"{c}5", f"={prev}5+1", FMT_INT)

        # Base-year actuals (green links) for context.
        link(ws, "B6", f"={self.hist('Revenue')}", FMT_MM)
        link(ws, "B7", f"={self.hist('Revenue Growth %')}", FMT_PCT)
        link(ws, "B8", f"={self.hist('EBITDA')}", FMT_MM)
        link(ws, "B9", f"={self.hist('EBITDA Margin %')}", FMT_PCT)
        link(ws, "B10", f"={self.hist('D&A')}", FMT_MM)
        link(ws, "B15", f"={self.hist('CapEx')}", FMT_MM)

        rows = [
            (6, "Revenue", FMT_MM, lambda p, c, k: f"={p}6*(1+Assumptions!{get_column_letter(1 + k)}${self.growth_row})"),
            (7, "  % growth", FMT_PCT, lambda p, c, k: f"=IFERROR({c}6/{p}6-1,0)"),
            (8, "EBITDA", FMT_MM, lambda p, c, k: f"={c}6*EBITDAMargin"),
            (9, "  % margin", FMT_PCT, lambda p, c, k: f"=IFERROR({c}8/{c}6,0)"),
            (10, "Less: D&A", FMT_MM, lambda p, c, k: f"={c}6*DAPct"),
            (11, "EBIT", FMT_MM, lambda p, c, k: f"={c}8-{c}10"),
            (12, "Less: taxes on EBIT", FMT_MM, lambda p, c, k: f"=MAX({c}11,0)*TaxRate"),
            (13, "NOPAT", FMT_MM, lambda p, c, k: f"={c}11-{c}12"),
            (14, "Plus: D&A", FMT_MM, lambda p, c, k: f"={c}10"),
            (15, "Less: CapEx", FMT_MM, lambda p, c, k: f"={c}6*CapExPct"),
            (16, "Unlevered free cash flow", FMT_MM, lambda p, c, k: f"={c}13+{c}14-{c}15"),
            (17, "  % of revenue", FMT_PCT, lambda p, c, k: f"=IFERROR({c}16/{c}6,0)"),
            (18, "Discount factor", FMT_FACTOR, lambda p, c, k: f"=1/(1+WACC)^{c}$5"),
            (19, "PV of unlevered FCF", FMT_MM, lambda p, c, k: f"={c}16*{c}18"),
        ]
        for row, text, fmt, make in rows:
            bold = row in (16, 19)
            label(ws, f"A{row}", text, bold=bold, italic=text.startswith("  "))
            for k, (p, c) in enumerate(zip(cols, proj), start=1):
                cell = calc(ws, f"{c}{row}", make(p, c, k), fmt, bold=bold)
                if row == 16:
                    cell.border = THIN_TOP

        # --- WACC (CAPM) ---
        section(ws, 21, 1, 7, "Cost of capital (WACC)")
        wacc = [
            (22, "Risk-free rate", "=RiskFree", FMT_PCT, link),
            (23, "Beta", "=EquityBeta", FMT_NUM, link),
            (24, "Equity risk premium", "=EquityRiskPremium", FMT_PCT, link),
            (25, "Cost of equity (CAPM: rf + β × ERP)", "=B22+B23*B24", FMT_PCT, calc),
            (26, "Pre-tax cost of debt", "=CostOfDebt", FMT_PCT, link),
            (27, "Tax rate", "=TaxRate", FMT_PCT, link),
            (28, "After-tax cost of debt", "=B26*(1-B27)", FMT_PCT, calc),
            (29, "Market value of equity (price × shares)", "=SharePrice*SharesOut", FMT_MM, calc),
            (30, "Book value of debt", "=BookDebt", FMT_MM, link),
            (31, "Debt / total capital", "=IFERROR(B30/(B29+B30),0)", FMT_PCT, calc),
            (32, "Materiality threshold", "=DebtThreshold", FMT_PCT, link),
            (33, "Debt weight used (0 if immaterial)", "=IF(B31>=B32,B31,0)", FMT_PCT, calc),
            (34, "Equity weight", "=1-B33", FMT_PCT, calc),
            (35, "WACC", "=B34*B25+B33*B28", FMT_PCT, calc),
        ]
        for row, text, formula, fmt, kind in wacc:
            label(ws, f"A{row}", text, bold=row == 35)
            kind(ws, f"B{row}", formula, fmt, bold=row == 35)
        ws["B35"].fill = KEY_FILL
        self.define("WACC", "DCF", "B35")

        # --- Terminal value, EV, equity value per share ---
        section(ws, 37, 1, 7, "Enterprise value → equity value per share")
        valuation = [
            (38, "Sum of PV of FCF (Years 1–5)", f"=SUM({first}19:{last}19)", FMT_MM, calc),
            (39, "Terminal growth rate", "=TerminalGrowth", FMT_PCT, link),
            (40, "Terminal value (Gordon growth, end of Year 5)", f"={last}16*(1+B39)/(WACC-B39)", FMT_MM, calc),
            (41, "PV of terminal value", f"=B40*{last}18", FMT_MM, calc),
            (42, "Enterprise value", "=B38+B41", FMT_MM, calc),
            (43, "  % of EV from terminal value", "=IFERROR(B41/B42,0)", FMT_PCT, calc),
            (44, "  Implied terminal EV / EBITDA", f'=IF({last}8>0,B40/{last}8,"n/m")', FMT_MULT, calc),
            (45, "Less: debt", "=BookDebt", FMT_MM, link),
            (46, "Plus: cash", "=Cash", FMT_MM, link),
            (47, "Equity value", "=B42-B45+B46", FMT_MM, calc),
            (48, "Shares outstanding (millions)", "=SharesOut", FMT_SHARES_MM, link),
            (49, "DCF value per share", "=B47/B48", FMT_PER_SHARE, calc),
            (50, "Check: WACC > terminal growth", '=IF(WACC>TerminalGrowth,"OK","ERROR – WACC must exceed g")',
             None, calc),
        ]
        for row, text, formula, fmt, kind in valuation:
            bold = row in (42, 47, 49)
            label(ws, f"A{row}", text, bold=bold, italic=text.startswith("  "))
            kind(ws, f"B{row}", formula, fmt, bold=bold)
        ws["B49"].fill = KEY_FILL
        self.define("EnterpriseValue", "DCF", "B42")
        self.define("PctEVFromTV", "DCF", "B43")
        self.define("EquityValue", "DCF", "B47")
        self.define("DCFValue", "DCF", "B49")

        widths(ws, A=42, **{c: 13 for c in cols})
        ws.freeze_panes = "B5"

    # ------------------------------------------------------------------ Beta Regression
    def _beta(self):
        ws, b, ticker = self.ws["Beta Regression"], self.d.beta, self.d.fin.ticker
        ret = b.returns
        n = len(ret)
        first, last = 5, 4 + n
        y, x = f"$J${first}:$J${last}", f"$K${first}:$K${last}"

        title(ws, "A1", "Beta Regression",
              f"OLS: weekly {ticker} return = α + β × weekly S&P 500 return · "
              f"{ret.index[0]:%b %d, %Y} – {ret.index[-1]:%b %d, %Y}")

        header_row(ws, 4, 1, ["Statistic", "statsmodels (Python)", "Excel check"])
        stats = [
            (5, "Beta (slope)", b.beta, f"=SLOPE({y},{x})", FMT_NUM4),
            (6, "Alpha (weekly intercept)", b.alpha, f"=INTERCEPT({y},{x})", FMT_NUM4),
            (7, "R-squared", b.r_squared, f"=RSQ({y},{x})", FMT_NUM4),
            (8, "Standard error of beta", b.std_err, f"=INDEX(LINEST({y},{x},TRUE,TRUE),2,1)", FMT_NUM4),
            (9, "Observations (weeks)", b.n_obs, f"=COUNT({y})", FMT_INT),
        ]
        for row, text, value, check, fmt in stats:
            label(ws, f"A{row}", text)
            inp(ws, f"B{row}", value, fmt)
            calc(ws, f"C{row}", check, fmt)
        derived = [
            (10, "95% CI – low (β − 1.96 × SE)", "=B5-1.96*B8"),
            (11, "95% CI – high (β + 1.96 × SE)", "=B5+1.96*B8"),
            (12, "Blume-adjusted beta (0.67β + 0.33)", "=0.67*B5+0.33"),
        ]
        for row, text, formula in derived:
            label(ws, f"A{row}", text)
            calc(ws, f"B{row}", formula, FMT_NUM4)
        label(ws, "A13", "Assumptions!Beta uses the raw beta (B5). Overwrite it there to use the adjusted beta instead.",
              italic=True, color=GREY)

        # Chart data: weekly returns, plus a two-point fit line driven by the stats above.
        header_row(ws, 4, 9, ["Week of", f"{ticker} return", "S&P 500 return"])
        for i, (dt, row) in enumerate(ret.iterrows()):
            r = first + i
            inp(ws, f"I{r}", dt.to_pydatetime(), "yyyy-mm-dd")
            inp(ws, f"J{r}", float(row["stock"]), "0.00%")
            inp(ws, f"K{r}", float(row["market"]), "0.00%")

        header_row(ws, 4, 13, ["Fit line: S&P 500", f"Fitted {ticker}"])
        calc(ws, "M5", f"=MIN({x})", "0.00%")
        calc(ws, "M6", f"=MAX({x})", "0.00%")
        for r in (5, 6):
            calc(ws, f"N{r}", f"=$B$6+$B$5*M{r}", "0.00%")

        chart = ScatterChart()
        chart.title = f"{ticker} vs S&P 500 weekly returns (β = {b.beta:.2f}, R² = {b.r_squared:.2f})"
        chart.style = 13
        chart.x_axis.title = "S&P 500 weekly return"
        chart.y_axis.title = f"{ticker} weekly return"
        chart.x_axis.number_format = chart.y_axis.number_format = "0%"
        chart.x_axis.delete = chart.y_axis.delete = False     # newer Excel hides axes otherwise
        chart.legend.position = "b"

        points = Series(Reference(ws, min_col=10, min_row=first, max_row=last),
                        Reference(ws, min_col=11, min_row=first, max_row=last), title="Weekly returns")
        points.marker.symbol = "circle"
        points.marker.size = 5
        points.graphicalProperties.line.noFill = True
        fit = Series(Reference(ws, min_col=14, min_row=5, max_row=6),
                     Reference(ws, min_col=13, min_row=5, max_row=6), title="OLS fit")
        fit.marker.symbol = "none"
        fit.graphicalProperties.line.solidFill = "C00000"
        fit.graphicalProperties.line.width = 28575
        chart.series.extend([points, fit])
        chart.width, chart.height = 17, 10
        ws.add_chart(chart, "A15")

        widths(ws, A=36, B=20, C=14, D=3, I=12, J=13, K=15, L=3, M=17, N=14)
        ws.freeze_panes = "A5"

    # ------------------------------------------------------------------ Summary
    def _summary(self):
        ws, fin = self.ws["Summary"], self.d.fin
        ws.sheet_view.showGridLines = False
        title(ws, "B2", f"{fin.entity_name} ({fin.ticker})",
              f"Equity valuation · prepared {self.d.as_of:%B %d, %Y} · $ millions except per share")

        header_row(ws, 5, 2, ["Valuation summary", "Per share", "vs. current"])
        label(ws, "B6", "Current share price")
        link(ws, "C6", "=SharePrice", FMT_PER_SHARE)
        label(ws, "B7", "DCF value per share")
        link(ws, "C7", "=DCFValue", FMT_PER_SHARE)
        calc(ws, "D7", '=IFERROR(C7/$C$6-1,"")', FMT_PCT)
        label(ws, "B8", "Price target", bold=True)
        calc(ws, "C8", "=C7", FMT_PER_SHARE, bold=True)
        calc(ws, "D8", '=IFERROR(C8/$C$6-1,"")', FMT_PCT, bold=True)
        for c in ("B8", "C8", "D8"):
            ws[c].fill = KEY_FILL
            ws[c].border = THIN_TOP

        header_row(ws, 10, 2, ["Key outputs", "Value", ""])
        outputs = [
            (11, "Market capitalization", "=SharePrice*SharesOut", FMT_MM, calc),
            (12, "Enterprise value (DCF)", "=EnterpriseValue", FMT_MM, link),
            (13, "Equity value (DCF)", "=EquityValue", FMT_MM, link),
            (14, "WACC", "=WACC", FMT_PCT, link),
            (15, "Terminal growth rate", "=TerminalGrowth", FMT_PCT, link),
            (16, "Beta", "=EquityBeta", FMT_NUM, link),
            (17, "% of EV from terminal value", "=PctEVFromTV", FMT_PCT, link),
            (18, "Model check", "=DCF!B50", None, link),
        ]
        for row, text, formula, fmt, kind in outputs:
            label(ws, f"B{row}", text)
            kind(ws, f"C{row}", formula, fmt)

        label(ws, "B20", "Change any blue cell on the Assumptions sheet and every number here updates.",
              italic=True, color=BLUE)
        widths(ws, A=2, B=36, C=16, D=14)
        ws.freeze_panes = "A6"


def write_workbook(data: ModelData, path: Path | str) -> Path:
    return ValuationWorkbook(data).build(path)
