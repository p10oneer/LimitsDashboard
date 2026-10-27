"""Investment-committee view of upcoming bond redemptions."""

import math
from datetime import date
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

from portfolio import add_months, filter_by_horizon, parse_portfolio

HORIZONS = {
    "6 месяцев": 6,
    "12 месяцев": 12,
    "24 месяца": 24,
    "Все": None,
}

CURRENCY_COLORS = {
    "KZT": "#3dba7a",
    "USD": "#5b8def",
    "EUR": "#f0a202",
    "UZS": "#c084fc",
    "RUB": "#ef6f6c",
    "GBP": "#2ec4b6",
    "CNY": "#e07a3d",
    "CHF": "#adb5bd",
    "AED": "#e9c46a",
    "KGS": "#90be6d",
}
_EXTRA_CURRENCY_COLORS = ("#f4a261", "#577590", "#b5179e", "#43aa8b", "#ffd166", "#6d597a")
# Labels shrink with the column, but never below this size. Past that the chart scrolls.
_MIN_FONT = 10
_MAX_FONT = 14
_CHAR_EM = 0.58
_LABEL_GAP_PX = 12
_WRAP_AT = 12
_AXIS_MARGIN_PX = 150
_VIEWPORT_PX = 1200


def main() -> None:
    st.set_page_config(page_title="Погашение облигаций", layout="wide")
    st.title("Погашение облигаций")
    st.caption("Когда освобождается лимит: дата погашения и номинал, пересчитанный в тенге по курсу из шапки файла.")

    with st.sidebar:
        st.header("Портфель")
        upload = st.file_uploader("Файл портфеля (.xlsx)", type=["xlsx"])
        horizon_label = st.radio("Горизонт", list(HORIZONS), index=0)

    if upload is None:
        st.info("Загрузите файл структуры портфеля. На графике останутся только облигации с датой погашения.")
        return

    try:
        portfolio = load_portfolio(upload.getvalue(), _PARSER_VERSION)
    except Exception as exc:
        st.error(f"Не удалось прочитать файл: {exc}")
        return

    report_date = portfolio.report_date or date.today()
    if portfolio.report_date is None:
        st.warning("В файле не найдена дата отчёта рядом с «за». Горизонт посчитан от сегодняшнего дня.")

    months = HORIZONS[horizon_label]
    window_end = add_months(report_date, months) if months is not None else None
    in_window = filter_by_horizon(portfolio.bonds, report_date, months)
    matured = _already_matured(portfolio.bonds, report_date)

    fund = portfolio.fund_name or "Портфель"
    if window_end is None:
        end_label = _format_date(in_window["maturity"].max()) if not in_window.empty else _format_date(report_date)
        window_text = f"все погашения с {_format_date(report_date)} по {end_label}"
    else:
        window_text = f"с {_format_date(report_date)} по {_format_date(window_end)}"
    st.subheader(f"{fund} — {window_text}")

    _metrics(in_window)
    missing_rates = _missing_rates(in_window)
    if missing_rates:
        st.warning("В шапке файла нет курса для: " + ", ".join(missing_rates) + ". Эти выпуски не попали на график.")
    _timeline(in_window.dropna(subset=["nominal_kzt"]) if not in_window.empty else in_window, report_date)
    if not matured.empty:
        st.caption(
            f"Не показаны {len(matured)} выпусков с датой погашения раньше даты отчёта."
        )
    _table(in_window, report_date)
    _exceptions(portfolio.exceptions)


# Bump when the parser output changes, so an old cached table is not reused.
_PARSER_VERSION = 3


@st.cache_data(show_spinner="Читаем портфель…")
def load_portfolio(payload: bytes, parser_version: int):
    del parser_version
    return parse_portfolio(BytesIO(payload))


def _already_matured(bonds: pd.DataFrame, report_date: date) -> pd.DataFrame:
    if bonds.empty:
        return bonds
    maturity = pd.to_datetime(bonds["maturity"])
    return bonds.loc[maturity < pd.Timestamp(report_date)].copy()


def _metrics(bonds: pd.DataFrame) -> None:
    issues, issuers, total = st.columns(3)
    issues.metric("Выпусков к погашению", 0 if bonds.empty else int(bonds["isin"].nunique()))
    issuers.metric("Эмитентов", 0 if bonds.empty else int(bonds["issuer"].nunique()))
    total.metric(
        "Номинал, KZT",
        _compact_amount(0 if bonds.empty or bonds["nominal_kzt"].isna().all() else bonds["nominal_kzt"].sum(min_count=1)),
    )


def _missing_rates(bonds: pd.DataFrame) -> list[str]:
    if bonds.empty or "nominal_kzt" not in bonds.columns:
        return []
    missing = bonds.loc[bonds["nominal_kzt"].isna(), "currency"].dropna().astype(str)
    return sorted(set(missing))


def _timeline(bonds: pd.DataFrame, report_date: date) -> None:
    if bonds.empty:
        st.info("В выбранном горизонте нет погашений. Выберите более длинный горизонт, чтобы увидеть все выпуски.")
        return
    zoom = st.slider(
        "Масштаб",
        min_value=0.7,
        max_value=2.0,
        value=1.0,
        step=0.05,
        help="Влево — больше выпусков на экране, подписи мельче. Вправо — крупнее, график прокручивается. Колёсико мыши на графике делает то же по датам. Подписи не уменьшаются ниже читаемого размера и не наезжают друг на друга.",
    )
    figure = build_timeline_figure(bonds, report_date, zoom)
    _show_scrollable_chart(figure, zoom)


def _show_scrollable_chart(figure, zoom: float) -> None:
    """Keep each label inside its own column. Zooming out shrinks type only down to a readable floor."""
    chart_height = int(figure.layout.height)
    figure_html = figure.to_html(
        full_html=False,
        include_plotlyjs=True,
        config={
            "responsive": False,
            "scrollZoom": True,
            "displaylogo": False,
            "modeBarButtonsToRemove": [
                "lasso2d",
                "select2d",
                "autoScale2d",
                "hoverClosestCartesian",
                "hoverCompareCartesian",
                "toggleSpikelines",
            ],
        },
    )
    page = f"""
    <div id="chart-host" style="width:100%; overflow-x:auto; overflow-y:hidden; background:#0e1117;">
      {figure_html}
    </div>
    <script>
    (function() {{
      var host = document.getElementById("chart-host");
      var zoom = {zoom};
      var minFont = {_MIN_FONT};
      var maxFont = {_MAX_FONT};
      var em = {_CHAR_EM};
      var gap = {_LABEL_GAP_PX};
      var margin = {_AXIS_MARGIN_PX};
      function metrics(container, visible, longest) {{
        var room = Math.max(container - margin, 160);
        var minSlot = Math.ceil(longest * minFont * em + gap);
        var desired = Math.max(minSlot, (room / visible) * zoom);
        var font = Math.floor((desired - gap) / (longest * em));
        font = Math.max(minFont, Math.min(maxFont, font));
        var textSlot = Math.ceil(longest * font * em + gap);
        var slot = textSlot;
        if (zoom > 1) slot = Math.max(textSlot, Math.floor((room / visible) * zoom));
        else if (textSlot * visible <= room) slot = Math.floor(room / visible);
        var width = Math.ceil(slot * visible + margin);
        if (width < container && font <= maxFont) width = container;
        return {{font: font, width: width}};
      }}
      function apply() {{
        var gd = host.querySelector(".plotly-graph-div");
        if (!gd || !gd._fullLayout || !gd.layout.meta) return;
        var meta = gd.layout.meta;
        var longest = Math.max(meta.longest, 1);
        var total = Math.max(meta.count, 1);
        var xa = gd._fullLayout.xaxis;
        var r0 = xa.range ? xa.range[0] : -0.5;
        var r1 = xa.range ? xa.range[1] : total - 0.5;
        var start = Math.max(0, Math.min(total - 1, Math.round(r0 + 0.5)));
        var end = Math.max(start + 1, Math.min(total, Math.round(r1 + 0.5)));
        var visible = end - start;
        var fit = metrics(host.clientWidth || {_VIEWPORT_PX}, visible, longest);
        var range0 = start - 0.5;
        var range1 = end - 0.5;
        var widthNow = gd._fullLayout.width || 0;
        if (Math.abs(r0 - range0) < 0.05 && Math.abs(r1 - range1) < 0.05 && Math.abs(widthNow - fit.width) < 2) return;
        var annotations = (gd.layout.annotations || []).map(function(item) {{
          var copy = Object.assign({{}}, item);
          copy.font = {{size: fit.font, color: "#fafafa"}};
          return copy;
        }});
        window.__chartLock = true;
        Plotly.relayout(gd, {{
          width: fit.width,
          "xaxis.range": [range0, range1],
          "xaxis.tickfont.size": fit.font,
          "yaxis.tickfont.size": Math.max(12, fit.font),
          "yaxis.title.font.size": Math.max(12, fit.font),
          annotations: annotations
        }}).then(function() {{
          host.scrollLeft = 0;
          window.__chartLock = false;
        }});
      }}
      function bind(gd) {{
        if (gd.__scaleBound) return;
        gd.__scaleBound = true;
        gd.on("plotly_relayout", function(event) {{
          if (window.__chartLock) return;
          var keys = Object.keys(event || {{}});
          var zoomed = keys.some(function(key) {{
            return key.indexOf("xaxis.range") === 0 || key === "xaxis.autorange";
          }});
          if (zoomed) apply();
        }});
        apply();
      }}
      function boot() {{
        var gd = host.querySelector(".plotly-graph-div");
        if (!gd || !gd._fullLayout) {{
          setTimeout(boot, 50);
          return;
        }}
        bind(gd);
      }}
      boot();
    }})();
    </script>
    """
    components.html(page, height=chart_height + 28, scrolling=False)


def build_timeline_figure(bonds: pd.DataFrame, report_date: date, zoom: float = 1.0):
    plot = _chart_frame(bonds, report_date)
    font, _slot, width = _chart_width(plot, zoom)
    label_lines = _label_lines(plot)
    currencies = _currency_order(plot["currency"])
    figure = px.bar(
        plot,
        x="key",
        y="nominal_scaled",
        color="currency",
        color_discrete_map=_currency_color_map(currencies),
        category_orders={"key": plot["key"].tolist(), "currency": currencies},
        labels={"currency": "Валюта"},
        custom_data=["issuer", "ticker", "isin", "maturity_text", "nominal_text", "currency", "days_left", "face_text"],
    )
    figure.update_traces(
        hovertemplate=(
            "%{customdata[0]}<br>"
            "Выпуск: %{customdata[1]}<br>"
            "ISIN: %{customdata[2]}<br>"
            "Валюта: %{customdata[5]}<br>"
            "Погашение: %{customdata[3]} · через %{customdata[6]} дн.<br>"
            "Номинал: %{customdata[7]} %{customdata[5]}<br>"
            "В тенге: %{customdata[4]}"
            "<extra></extra>"
        ),
    )
    highest = float(plot["nominal_scaled"].max())
    for row in plot.itertuples(index=False):
        figure.add_annotation(
            x=row.key,
            y=float(row.nominal_scaled),
            text=f"{row.short_ticker}<br>{row.bar_text}",
            showarrow=False,
            yanchor="bottom",
            yshift=3,
            font=dict(size=font, color="#fafafa"),
            align="center",
        )
    figure.update_xaxes(
        tickmode="array",
        tickvals=plot["key"].tolist(),
        ticktext=plot["maturity_text"].tolist(),
        tickangle=0,
        title=None,
        tickfont=dict(size=font, color="#fafafa"),
        fixedrange=False,
        automargin=True,
    )
    figure.update_yaxes(
        title=plot["axis_title"].iloc[0],
        title_font=dict(size=max(12, font), color="#fafafa"),
        tickformat=".2f",
        tickfont=dict(size=max(12, font), color="#fafafa"),
        rangemode="tozero",
        range=[0, highest * (1.18 + 0.08 * label_lines)],
        fixedrange=True,
        automargin=True,
    )
    figure.update_layout(
        title=None,
        width=width,
        height=640,
        bargap=0.42,
        autosize=False,
        dragmode="zoom",
        margin=dict(l=72, r=24, t=max(56, 16 + label_lines * (font + 2)), b=16),
        paper_bgcolor="#0e1117",
        plot_bgcolor="#0e1117",
        font=dict(color="#fafafa", size=font),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            x=0,
            font=dict(size=13, color="#fafafa"),
            title=dict(text="Валюта", font=dict(size=13, color="#fafafa")),
        ),
        hoverlabel=dict(font_size=13),
        meta={"longest": _longest_label(plot), "count": int(len(plot))},
    )
    return figure


def _currency_order(currencies: pd.Series) -> list[str]:
    present = list(dict.fromkeys(currencies.tolist()))
    preferred = [code for code in CURRENCY_COLORS if code in present]
    rest = sorted(code for code in present if code not in preferred)
    return preferred + rest


def _currency_color_map(currencies: list[str]) -> dict[str, str]:
    extras = iter(_EXTRA_CURRENCY_COLORS)
    mapping = {}
    for code in currencies:
        mapping[code] = CURRENCY_COLORS.get(code) or next(extras, "#8d99ae")
    return mapping


def _chart_frame(bonds: pd.DataFrame, report_date: date) -> pd.DataFrame:
    plot = bonds.sort_values(["maturity", "issuer", "isin"]).reset_index(drop=True).copy()
    plot["maturity_text"] = plot["maturity"].map(_format_date)
    plot["nominal_text"] = plot["nominal_kzt"].map(_format_kzt)
    plot["face_text"] = plot["nominal_amount"].map(_format_kzt)
    scale, axis_title = _nominal_axis(plot["nominal_kzt"])
    plot["nominal_scaled"] = plot["nominal_kzt"] / scale
    plot["axis_title"] = axis_title
    plot["days_left"] = [
        (pd.Timestamp(day) - pd.Timestamp(report_date)).days for day in plot["maturity"]
    ]
    plot["currency"] = plot["currency"].fillna("—").astype(str).str.strip().replace({"": "—"})
    plot["ticker"] = plot["ticker"].fillna("—").astype(str)
    plot["short_ticker"] = plot["ticker"].map(lambda ticker: _wrap_ticker(_short_ticker(ticker)))
    plot["key"] = [f"{index:02d}-{isin}" for index, isin in enumerate(plot["isin"])]
    plot["bar_text"] = plot["nominal_kzt"].map(_compact_amount)
    return plot


def _longest_label(plot: pd.DataFrame) -> int:
    return max(
        _longest_line(plot["maturity_text"]),
        _longest_line(plot["short_ticker"]),
        _longest_line(plot["bar_text"]),
        1,
    )


def _longest_line(values) -> int:
    longest = 1
    for value in values.astype(str):
        for line in value.split("<br>"):
            longest = max(longest, len(line))
    return longest


def _label_lines(plot: pd.DataFrame) -> int:
    """Ticker lines plus the amount line sitting above each bar."""
    ticker_lines = max(str(value).count("<br>") + 1 for value in plot["short_ticker"])
    return ticker_lines + 1


def _chart_width(plot: pd.DataFrame, zoom: float) -> tuple[int, int, int]:
    """Font size, column width, and figure width for this zoom.

    Zoom 1 fills the screen when the labels still fit. A larger zoom widens
    each column. The font never drops below the readable floor; the figure
    gets wider and scrolls instead of letting labels overlap.
    """
    longest = _longest_label(plot)
    count = max(len(plot), 1)
    room = _VIEWPORT_PX - _AXIS_MARGIN_PX
    min_slot = math.ceil(longest * _MIN_FONT * _CHAR_EM + _LABEL_GAP_PX)
    desired = max(min_slot, (room / count) * zoom)
    font = int((desired - _LABEL_GAP_PX) / (longest * _CHAR_EM))
    font = max(_MIN_FONT, min(_MAX_FONT, font))
    text_slot = math.ceil(longest * font * _CHAR_EM + _LABEL_GAP_PX)
    if zoom > 1:
        slot = max(text_slot, int((room / count) * zoom))
    elif text_slot * count <= room:
        slot = int(room / count)
    else:
        slot = text_slot
    width = int(slot * count + _AXIS_MARGIN_PX)
    return font, slot, max(width, _VIEWPORT_PX if slot * count <= room else width)


def _nominal_axis(values: pd.Series) -> tuple[float, str]:
    highest = float(values.max()) if values.notna().any() else 0.0
    if highest >= 1_000_000_000:
        return 1_000_000_000, "Номинал, млрд KZT"
    if highest >= 1_000_000:
        return 1_000_000, "Номинал, млн KZT"
    return 1.0, "Номинал, KZT"


def _short_ticker(ticker: str) -> str:
    if not ticker or ticker == "—":
        return "—"
    return ticker.strip()


def _wrap_ticker(ticker: str, width: int = _WRAP_AT) -> str:
    """Break a long ticker on spaces so the column stays narrow and the full name remains."""
    if len(ticker) <= width:
        return ticker
    lines: list[str] = []
    current = ""
    for word in ticker.split():
        if not current:
            current = word
        elif len(current) + 1 + len(word) <= width:
            current = f"{current} {word}"
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return "<br>".join(lines)


def _compact_amount(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    if abs(value) >= 1_000_000_000:
        return f"{value / 1_000_000_000:.2f} млрд".replace(".", ",")
    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:.1f} млн".replace(".", ",")
    return _format_kzt(value)


def _table(bonds: pd.DataFrame, report_date: date) -> None:
    st.subheader("Выпуски в выбранном горизонте")
    if bonds.empty:
        st.caption("Таблица пуста.")
        return
    view = pd.DataFrame(
        {
            "Эмитент": bonds["issuer"],
            "Тикер": bonds["ticker"],
            "ISIN": bonds["isin"],
            "Тип": bonds["security_type"],
            "Валюта": bonds["currency"],
            "Дата погашения": pd.to_datetime(bonds["maturity"]),
            "Дней до погашения": [
                (pd.Timestamp(day) - pd.Timestamp(report_date)).days for day in bonds["maturity"]
            ],
            "Количество": bonds["quantity"],
            "Купон, %": bonds["coupon"],
            "Ближайший купон": pd.to_datetime(bonds["next_coupon"]),
            "Количество по номиналу": bonds["nominal_amount"],
            "Номинал, KZT": bonds["nominal_kzt"],
            "Стоимость, KZT": bonds["value_kzt"],
        }
    )
    st.dataframe(
        view,
        hide_index=True,
        width="stretch",
        column_config={
            "Дата погашения": st.column_config.DateColumn(format="DD.MM.YYYY"),
            "Ближайший купон": st.column_config.DateColumn(format="DD.MM.YYYY"),
            "Количество": st.column_config.NumberColumn(format="%.0f"),
            "Купон, %": st.column_config.NumberColumn(format="%.3f"),
            "Количество по номиналу": st.column_config.NumberColumn(format="%,.0f"),
            "Номинал, KZT": st.column_config.NumberColumn(format="%,.0f"),
            "Стоимость, KZT": st.column_config.NumberColumn(format="%,.0f"),
        },
    )


def _exceptions(exceptions: pd.DataFrame) -> None:
    with st.expander(f"Облигации без даты погашения ({len(exceptions)})", expanded=False):
        if exceptions.empty:
            st.caption("Таких строк нет.")
            return
        view = exceptions.rename(
            columns={
                "ticker": "Тикер",
                "issuer": "Эмитент",
                "isin": "ISIN",
                "security_type": "Тип",
                "currency": "Валюта",
                "quantity": "Количество",
                "value_kzt": "Стоимость, KZT",
            }
        )
        st.dataframe(view, hide_index=True, width="stretch")
        st.caption("Эти строки имеют тип облигации и эмитента, но без даты погашения их нельзя поставить на шкалу.")


def _format_date(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    if isinstance(value, pd.Timestamp):
        value = value.date()
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return "—"
    return parsed.strftime("%d.%m.%Y")


def _format_kzt(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"{value:,.0f}".replace(",", " ")


def _running_inside_streamlit() -> bool:
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
    except Exception:
        return False
    return get_script_run_ctx() is not None


if __name__ == "__main__" or _running_inside_streamlit():
    main()
