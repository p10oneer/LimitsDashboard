"""Investment-committee view of upcoming bond redemptions."""

import hashlib
import html
import math
import re
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import streamlit.components.v1 as components

from portfolio import add_months, cash_events, filter_by_horizon, parse_portfolio

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
_PORTFOLIO_PATTERNS = ("", "/", "x", "-", "+", ".", "\\", "|")
_PORTFOLIO_SYMBOLS = ("square", "diamond", "triangle-up", "cross", "star", "hexagon", "x", "circle")
ROOT = Path(__file__).resolve().parent
LOGO = ROOT / "assets" / "brand" / "logo-green.svg"
MARK = ROOT / "assets" / "brand" / "mark-green.svg"


def _inject_css() -> None:
    css_path = ROOT / "assets" / "brand" / "app.css"
    css = css_path.read_text(encoding="utf-8") if css_path.is_file() else ""
    if css:
        st.markdown(f"<style>\n{css}\n</style>", unsafe_allow_html=True)


def _block_title(text: str, color: str) -> None:
    st.markdown(
        (
            f'<h2 class="block-title" style="color:#000000;border-bottom:3px solid {color}">'
            f"{html.escape(text)}</h2>"
        ),
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="Погашение облигаций | BCC Invest",
        page_icon=str(MARK) if MARK.exists() else None,
        layout="wide",
    )
    _inject_css()
    if LOGO.exists():
        st.logo(str(LOGO), icon_image=str(MARK) if MARK.exists() else str(LOGO), size="medium")

    if "portfolios" not in st.session_state:
        st.session_state.portfolios = []

    with st.sidebar:
        st.header("Портфели")
        uploads = st.file_uploader(
            "Файлы портфелей (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            help="Можно добавить несколько файлов. Структура каждого такая же, как у портфеля-прототипа.",
        )
        _remember_uploads(uploads)

    if not st.session_state.portfolios:
        st.info("Загрузите один или несколько файлов структуры портфеля. На графике останутся только облигации с датой погашения.")
        return

    loaded = list(st.session_state.portfolios)
    cash_frames: list[pd.DataFrame] = []
    report_dates: list[date] = []
    date_warnings: list[str] = []
    for order, item in enumerate(loaded):
        portfolio = item["portfolio"]
        report_date = portfolio.report_date or date.today()
        report_dates.append(report_date)
        if portfolio.report_date is None:
            date_warnings.append(
                f"В «{item['name']}» не найдена дата отчёта рядом с «за». Горизонт посчитан от сегодняшнего дня."
            )
        held = portfolio.bonds.copy()
        if not held.empty:
            held["portfolio"] = item["name"]
            held["label"] = item["label"] or item["name"]
            held["portfolio_order"] = order
            held["report_date"] = report_date
            cash_frames.append(held)

    labels = [item["label"] or item["name"] for item in loaded]
    anchor = min(report_dates)
    with st.container(border=True, key="redemption-block"):
        _block_title("Погашения облигаций", "#0B534F")
        st.caption("Когда освобождается лимит: дата погашения и номинал, пересчитанный в тенге по курсу из шапки файла.")
        horizon_label = st.radio(
            "Горизонт погашений",
            list(HORIZONS),
            index=0,
            horizontal=True,
            key="redemption_horizon",
        )
        months = HORIZONS[horizon_label]
        frames: list[pd.DataFrame] = []
        matured_count = 0
        missing_rates: set[str] = set()
        for order, item in enumerate(loaded):
            portfolio = item["portfolio"]
            report_date = report_dates[order]
            window = filter_by_horizon(portfolio.bonds, report_date, months)
            matured_count += len(_already_matured(portfolio.bonds, report_date))
            missing_rates.update(_missing_rates(window))
            if window.empty:
                continue
            part = window.copy()
            part["portfolio"] = item["name"]
            part["label"] = item["label"] or item["name"]
            part["portfolio_order"] = order
            part["report_date"] = report_date
            frames.append(part)
        in_window = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        window_end = add_months(anchor, months) if months is not None else None
        if window_end is None:
            end_label = _format_date(in_window["maturity"].max()) if not in_window.empty else _format_date(anchor)
            window_text = f"все погашения с {_format_date(anchor)} по {end_label}"
        elif len(set(report_dates)) == 1:
            window_text = f"с {_format_date(anchor)} по {_format_date(window_end)}"
        else:
            window_text = f"{horizon_label.lower()} от даты отчёта каждого портфеля"
        st.caption(window_text)
        for message in date_warnings:
            st.warning(message)
        _metrics(in_window)
        if missing_rates:
            st.warning("В шапке файла нет курса для: " + ", ".join(sorted(missing_rates)) + ". Эти выпуски не попали на график.")
        _timeline(in_window, anchor, labels)
        if matured_count:
            st.caption(f"Не показаны {matured_count} выпусков с датой погашения раньше даты отчёта своего портфеля.")
        _table(in_window, anchor)
    _cashflow(
        pd.concat(cash_frames, ignore_index=True) if cash_frames else pd.DataFrame(),
        labels,
    )


# Bump when the parser output changes, so an old cached table is not reused.
_PARSER_VERSION = 4


@st.cache_data(show_spinner="Читаем портфель…")
def load_portfolio(payload: bytes, parser_version: int):
    del parser_version
    return parse_portfolio(BytesIO(payload))


def _remember_uploads(uploads) -> None:
    """Keep the loaded portfolios in step with the files still in the uploader."""
    by_key = {item["key"]: item for item in st.session_state.portfolios}
    kept = []
    seen: set[str] = set()
    for file in uploads or []:
        payload = file.getvalue()
        key = f"{file.name}:{hashlib.sha256(payload).hexdigest()[:12]}"
        if key in seen:
            continue
        seen.add(key)
        if key in by_key:
            kept.append(by_key[key])
            continue
        try:
            parsed = load_portfolio(payload, _PARSER_VERSION)
        except Exception as exc:
            st.error(f"Не удалось прочитать {file.name}: {exc}")
            continue
        name = parsed.fund_name or file.name
        kept.append(
            {
                "key": key,
                "filename": file.name,
                "name": name,
                "label": _unique_label(_default_label(name, len(kept) + 1), {item["label"] for item in kept}),
                "visible": True,
                "portfolio": parsed,
            }
        )
    st.session_state.portfolios = kept


def _portfolio_name(name: str) -> str:
    text = re.sub(r"\s+", " ", (name or "").replace("\xa0", " ").replace("\n", " ")).strip()
    quoted = [part.strip() for part in re.findall(r'"([^"]+)"', text) + re.findall(r"«([^»]+)»", text)]
    quoted = [part for part in quoted if part]
    if quoted:
        return quoted[0]
    return text


def _default_label(name: str, index: int) -> str:
    return _portfolio_name(name) or f"П{index}"


def _unique_label(label: str, taken: set[str] | None = None) -> str:
    used = set(taken) if taken is not None else {item["label"] for item in st.session_state.portfolios}
    if label not in used:
        return label
    number = 2
    while f"{label} {number}" in used:
        number += 1
    return f"{label} {number}"


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


def _timeline(bonds: pd.DataFrame, report_date: date, portfolio_labels: list[str]) -> None:
    if bonds.empty or "nominal_kzt" not in bonds.columns:
        st.info("В выбранном горизонте нет погашений. Выберите более длинный горизонт, чтобы увидеть все выпуски.")
        return
    bonds = bonds.dropna(subset=["nominal_kzt"])
    if bonds.empty:
        st.info("В выбранном горизонте нет выпусков с курсом для пересчёта номинала в тенге.")
        return
    bonds, labels = _chart_filters(bonds, portfolio_labels, "redemption")
    if not labels:
        st.info("На графике ничего не выбрано. Отметьте хотя бы один портфель и одну валюту.")
        return
    if bonds.empty and len(labels) <= 1:
        st.info("В выбранном горизонте нет погашений для отмеченных портфелей и валют.")
        return
    if len(labels) > 1:
        st.caption("Каждая строка — отдельный портфель со своей шкалой. Цвет — валюта. Одинаковый выпуск стоит в одной колонке, суммы не складываются.")
    figure = build_timeline_figure(bonds, report_date, 1.0, labels)
    _show_scrollable_chart(figure, 1.0, int(figure.layout.meta["bars"]))


def _chart_filters(
    bonds: pd.DataFrame,
    portfolio_labels: list[str],
    key_prefix: str,
) -> tuple[pd.DataFrame, list[str]]:
    """Portfolios and currencies chosen for one chart. Each chart keeps its own selection."""
    labels = list(dict.fromkeys(portfolio_labels)) or list(dict.fromkeys(bonds["label"].astype(str)))
    currencies = _currency_order(bonds["currency"])
    picked_labels = labels
    picked_currencies = currencies
    show_portfolios = len(labels) > 1
    show_currencies = len(currencies) > 1
    if show_portfolios or show_currencies:
        slots = st.columns(2 if show_portfolios and show_currencies else 1)
        slot = 0
        if show_portfolios:
            portfolio_key = f"{key_prefix}-portfolios"
            _sync_multiselect(portfolio_key, labels)
            picked_labels = slots[slot].multiselect("Портфели на графике", labels, key=portfolio_key)
            slot += 1
        if show_currencies:
            currency_key = f"{key_prefix}-currencies"
            _sync_multiselect(currency_key, currencies)
            picked_currencies = slots[slot].multiselect("Валюты на графике", currencies, key=currency_key)
    chosen = bonds.loc[bonds["label"].astype(str).isin(picked_labels) & bonds["currency"].isin(picked_currencies)]
    if not picked_currencies:
        return chosen.iloc[0:0], []
    return chosen, [label for label in labels if label in picked_labels]


def _sync_multiselect(widget_key: str, options: list[str]) -> None:
    """Keep a multiselect in sync when the available choices change.

    A choice the user removed stays off. A choice that appears later starts on.
    """
    seen_key = f"{widget_key}:seen"
    seen = list(st.session_state.get(seen_key, []))
    if widget_key not in st.session_state:
        chosen = list(options)
    else:
        chosen = list(st.session_state[widget_key])
        for option in options:
            if option not in seen:
                chosen.append(option)
    st.session_state[widget_key] = [option for option in options if option in chosen]
    st.session_state[seen_key] = list(dict.fromkeys([*seen, *options]))


def _show_scrollable_chart(figure, zoom: float, bars: int) -> None:
    """Keep each label inside its own column. Zooming out shrinks type only down to a readable floor."""
    page, height = _chart_page(figure, zoom, bars)
    components.html(page, height=height, scrolling=False)


def _chart_page(figure, zoom: float, bars: int) -> tuple[str, int]:
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
    <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&amp;display=swap">
    <div id="chart-frame" style="position:relative; width:100%; background:#FFFFFF; font-family:Inter, Arial, sans-serif;">
      <div id="chart-detail" style="display:none; position:absolute; z-index:30; width:380px; overflow:auto; box-sizing:border-box; padding:10px 12px; background:#FFFFFF; color:#21252B; border:1px solid #E1E4E8; border-radius:8px; font:13px/1.45 sans-serif; box-shadow:0 8px 24px rgba(0,0,0,0.45);"></div>
      <div id="chart-host" style="width:100%; overflow-x:auto; overflow-y:hidden; background:#FFFFFF;">
        {figure_html}
      </div>
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
      var bars = Math.max({bars}, 1);
      function metrics(container, visible, longest) {{
        var room = Math.max(container - margin, 160);
        var minSlot = Math.ceil(longest * minFont * em + gap) * bars;
        var desired = Math.max(minSlot, (room / visible) * zoom);
        var font = Math.floor((desired / bars - gap) / (longest * em));
        font = Math.max(minFont, Math.min(maxFont, font));
        var textSlot = Math.ceil(longest * font * em + gap) * bars;
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
          copy.font = {{size: fit.font, color: "#21252B"}};
          return copy;
        }});
        var updates = {{
          width: fit.width,
          height: gd._fullLayout.height,
          annotations: annotations
        }};
        Object.keys(gd._fullLayout).forEach(function(key) {{
          if (/^xaxis\\d*$/.test(key)) {{
            updates[key + ".range"] = [range0, range1];
            updates[key + ".tickfont.size"] = fit.font;
          }}
          if (/^yaxis\\d*$/.test(key)) {{
            updates[key + ".tickfont.size"] = Math.max(11, fit.font);
            updates[key + ".title.font.size"] = 13;
          }}
        }});
        window.__chartLock = true;
        Plotly.relayout(gd, updates).then(function() {{
          host.scrollLeft = 0;
          window.__chartLock = false;
        }});
      }}
      function bind(gd) {{
        if (gd.__scaleBound) return;
        gd.__scaleBound = true;
        var detail = document.getElementById("chart-detail");
        var frame = document.getElementById("chart-frame");
        var hideTimer = null;
        if (detail && gd.layout.meta && gd.layout.meta.scroll_detail) {{
          function place(box) {{
            var bounds = frame.getBoundingClientRect();
            var rect = box.getBoundingClientRect();
            var width = 380;
            var left = rect.right - bounds.left + 12;
            if (left + width > bounds.width - 8) left = Math.max(8, rect.left - bounds.left - width - 12);
            var top = Math.max(8, rect.top - bounds.top);
            if (top + 160 > bounds.height) top = Math.max(8, bounds.height - 168);
            detail.style.left = left + "px";
            detail.style.top = top + "px";
            detail.style.maxHeight = Math.max(140, bounds.height - top - 8) + "px";
          }}
          function showFromLabel() {{
            var box = host.querySelector(".hovertext");
            if (!box) return;
            var lines = [];
            box.querySelectorAll("tspan").forEach(function(item) {{
              var value = (item.textContent || "").trim();
              if (value) lines.push(value);
            }});
            if (!lines.length) return;
            clearTimeout(hideTimer);
            place(box);
            detail.textContent = "";
            lines.forEach(function(line, index) {{
              if (index) detail.appendChild(document.createElement("br"));
              detail.appendChild(document.createTextNode(line));
            }});
            detail.style.display = "block";
            box.style.opacity = "0";
          }}
          function scheduleHide() {{
            clearTimeout(hideTimer);
            hideTimer = setTimeout(function() {{ detail.style.display = "none"; }}, 400);
          }}
          new MutationObserver(showFromLabel).observe(host, {{subtree: true, childList: true, characterData: true}});
          detail.addEventListener("mouseenter", function() {{ clearTimeout(hideTimer); }});
          detail.addEventListener("mouseleave", scheduleHide);
          host.addEventListener("mouseleave", scheduleHide);
        }}
        gd.on("plotly_relayout", function(event) {{
          if (window.__chartLock) return;
          var keys = Object.keys(event || {{}});
          var zoomed = keys.some(function(key) {{
            return /^xaxis\\d*\\.range/.test(key) || /^xaxis\\d*\\.autorange$/.test(key);
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
    return page, chart_height + 28


def _build_portfolio_rows(plot: pd.DataFrame, labels: list[str], zoom: float):
    """One row per portfolio, with the same bond in the same column."""
    font, _slot, width = _chart_width(plot if not plot.empty else _empty_plot_frame(), zoom, bars=1)
    currencies = _currency_order(plot["currency"]) if not plot.empty else []
    colors = _currency_color_map(currencies)
    if plot.empty:
        keys: list[str] = []
        date_by_key = pd.Series(dtype=str)
    else:
        keys = list(dict.fromkeys(plot.sort_values(["maturity", "isin"])["key"].tolist()))
        date_by_key = plot.drop_duplicates("key").set_index("key")["maturity_text"]
    row_count = len(labels)
    figure = make_subplots(
        rows=row_count,
        cols=1,
        shared_xaxes=True,
        shared_yaxes=False,
        vertical_spacing=min(0.14, 0.72 / max(row_count - 1, 1)),
    )
    for currency in currencies:
        figure.add_trace(
            go.Scatter(
                x=[],
                y=[],
                mode="markers",
                name=currency,
                legendgroup="currency",
                legendgrouptitle=dict(text="Валюта", font=dict(color="#21252B", size=13)),
                marker=dict(size=12, color=colors[currency], symbol="square"),
                hoverinfo="skip",
                showlegend=True,
            ),
            row=1,
            col=1,
        )
    hover = (
        "%{customdata[8]}<br>"
        "%{customdata[0]}<br>"
        "Выпуск: %{customdata[1]}<br>"
        "ISIN: %{customdata[2]}<br>"
        "Валюта: %{customdata[5]}<br>"
        "Погашение: %{customdata[3]} · через %{customdata[6]} дн.<br>"
        "Номинал: %{customdata[7]} %{customdata[5]}<br>"
        "В тенге: %{customdata[4]}"
        "<extra></extra>"
    )
    for index, label in enumerate(labels, start=1):
        part = plot.loc[plot["label"] == label] if not plot.empty else plot
        if not part.empty:
            figure.add_trace(
                go.Bar(
                    x=part["key"],
                    y=part["nominal_scaled"],
                    marker=dict(color=[colors[currency] for currency in part["currency"]]),
                    text=part["bar_text"],
                    textposition="outside",
                    textfont=dict(size=font, color="#21252B"),
                    cliponaxis=True,
                    customdata=part[
                        ["issuer", "ticker", "isin", "maturity_text", "nominal_text", "currency", "days_left", "face_text", "portfolio"]
                    ].to_numpy(),
                    hovertemplate=hover,
                    showlegend=False,
                ),
                row=index,
                col=1,
            )
            for row in part.itertuples(index=False):
                figure.add_annotation(
                    x=row.key,
                    y=float(row.nominal_scaled),
                    text=row.short_ticker,
                    showarrow=False,
                    yanchor="bottom",
                    yshift=font + 6,
                    font=dict(size=font, color="#21252B"),
                    align="center",
                    xref="x" if index == 1 else f"x{index}",
                    yref="y" if index == 1 else f"y{index}",
                )
        else:
            figure.add_annotation(
                x=0.5,
                y=0.5,
                xref="x domain" if index == 1 else f"x{index} domain",
                yref="y domain" if index == 1 else f"y{index} domain",
                text="нет погашений в этом горизонте",
                showarrow=False,
                font=dict(size=13, color="#545D6D"),
            )
        figure.add_annotation(
            x=0,
            y=1,
            xref="x domain" if index == 1 else f"x{index} domain",
            yref="y domain" if index == 1 else f"y{index} domain",
            text=label,
            showarrow=False,
            xanchor="left",
            yanchor="bottom",
            yshift=4,
            font=dict(size=14, color="#21252B"),
        )
        row_max = float(part["nominal_scaled"].max()) if not part.empty else 1.0
        if row_max >= 0.1:
            tickformat = ".2f"
        elif row_max >= 0.01:
            tickformat = ".3f"
        else:
            tickformat = ".4f"
        figure.update_yaxes(
            title_text="",
            tickformat=tickformat,
            tickfont=dict(size=max(11, font - 1), color="#21252B"),
            rangemode="tozero",
            range=[0, max(row_max, 1e-9) * 2.05],
            fixedrange=True,
            automargin=True,
            row=index,
            col=1,
        )
        figure.update_xaxes(
            type="category",
            categoryorder="array",
            categoryarray=keys,
            showticklabels=True,
            tickmode="array",
            tickvals=keys,
            ticktext=[date_by_key.loc[key] for key in keys],
            tickangle=0,
            tickfont=dict(size=font, color="#21252B"),
            fixedrange=False,
            automargin=True,
            row=index,
            col=1,
        )
    figure.update_layout(
        title=None,
        width=width,
        height=96 + row_count * 280,
        barmode="group",
        bargap=0.35,
        autosize=False,
        dragmode="zoom",
        margin=dict(l=64, r=24, t=72, b=16),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Inter, Arial, sans-serif", color="#21252B", size=font),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.01,
            x=0,
            font=dict(size=13, color="#21252B"),
            itemsizing="constant",
            itemclick=False,
            itemdoubleclick=False,
        ),
        hoverlabel=dict(font_size=13),
        meta={"longest": _longest_label(plot, False) if not plot.empty else 8, "count": max(len(keys), 1), "bars": 1},
    )
    figure.update_xaxes(gridcolor="#E1E4E8", zerolinecolor="#E1E4E8")
    figure.update_yaxes(gridcolor="#E1E4E8", zerolinecolor="#E1E4E8")
    return figure


def _empty_plot_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "key": ["—"],
            "maturity_text": ["—"],
            "short_ticker": ["—"],
            "bar_text": ["—"],
            "label": [""],
        }
    )


def build_timeline_figure(
    bonds: pd.DataFrame,
    report_date: date,
    zoom: float = 1.0,
    portfolios: list[str] | None = None,
):
    plot = _chart_frame(bonds, report_date) if not bonds.empty else bonds
    if portfolios:
        labels = list(portfolios)
    elif plot.empty:
        labels = []
    elif "portfolio_order" in plot.columns:
        labels = list(dict.fromkeys(plot.sort_values("portfolio_order")["label"].tolist()))
    else:
        labels = list(dict.fromkeys(plot["label"].tolist()))
    if len(labels) > 1:
        return _build_portfolio_rows(plot, labels, zoom)
    several = len(labels) > 1
    bars = len(labels) if several else 1
    font, _slot, width = _chart_width(plot, zoom, bars)
    label_lines = _label_lines(plot, several)
    currencies = _currency_order(plot["currency"])
    colors = _currency_color_map(currencies)
    keys = list(dict.fromkeys(plot["key"].tolist()))
    date_by_key = plot.drop_duplicates("key").set_index("key")["maturity_text"]
    figure = go.Figure()
    for currency in currencies:
        figure.add_trace(
            go.Scatter(
                x=[],
                y=[],
                mode="markers",
                name=currency,
                legendgroup="currency",
                legendgrouptitle=dict(text="Валюта", font=dict(color="#21252B", size=13)),
                marker=dict(size=12, color=colors[currency], symbol="square"),
                hoverinfo="skip",
                showlegend=True,
            )
        )
    hover = (
        "%{customdata[8]}<br>"
        "%{customdata[0]}<br>"
        "Выпуск: %{customdata[1]}<br>"
        "ISIN: %{customdata[2]}<br>"
        "Валюта: %{customdata[5]}<br>"
        "Погашение: %{customdata[3]} · через %{customdata[6]} дн.<br>"
        "Номинал: %{customdata[7]} %{customdata[5]}<br>"
        "В тенге: %{customdata[4]}"
        "<extra></extra>"
    )
    for index, label in enumerate(labels):
        part = plot.loc[plot["label"] == label]
        pattern = _PORTFOLIO_PATTERNS[index % len(_PORTFOLIO_PATTERNS)]
        text = part["bar_text"] + ("<br>" + part["label"] if several else "")
        figure.add_trace(
            go.Bar(
                x=part["key"],
                y=part["nominal_scaled"],
                name=label,
                legendgroup="portfolio",
                legendgrouptitle=dict(text="Портфель", font=dict(color="#21252B", size=13)),
                offsetgroup=label,
                marker=dict(
                    color=[colors[currency] for currency in part["currency"]],
                    pattern=dict(
                        shape=pattern,
                        fgcolor="rgba(255,255,255,0.9)",
                        solidity=0 if pattern == "" else 0.45,
                        size=8,
                    ),
                ),
                text=text,
                textposition="outside",
                textfont=dict(size=font, color="#21252B"),
                cliponaxis=False,
                customdata=part[
                    ["issuer", "ticker", "isin", "maturity_text", "nominal_text", "currency", "days_left", "face_text", "portfolio"]
                ].to_numpy(),
                hovertemplate=hover,
                showlegend=False,
            )
        )
        if several:
            figure.add_trace(
                go.Scatter(
                    x=[],
                    y=[],
                    mode="markers",
                    name=label,
                    legendgroup="portfolio",
                    legendgrouptitle=dict(text="Портфель", font=dict(color="#21252B", size=13)),
                    marker=dict(
                        size=12,
                        color="#d7dde5",
                        symbol=_PORTFOLIO_SYMBOLS[index % len(_PORTFOLIO_SYMBOLS)],
                        line=dict(width=1, color="#21252B"),
                    ),
                    hoverinfo="skip",
                    showlegend=True,
                )
            )
    highest = float(plot["nominal_scaled"].max())
    grouped = plot.groupby("key", sort=False)["nominal_scaled"].max()
    tickers = plot.drop_duplicates("key").set_index("key")["short_ticker"]
    for key in keys:
        figure.add_annotation(
            x=key,
            y=float(grouped.loc[key]),
            text=tickers.loc[key],
            showarrow=False,
            yanchor="bottom",
            yshift=4 + (font + 3) * (2 if several else 1),
            font=dict(size=font, color="#21252B"),
            align="center",
        )
    figure.update_xaxes(
        type="category",
        categoryorder="array",
        categoryarray=keys,
        tickmode="array",
        tickvals=keys,
        ticktext=[date_by_key.loc[key] for key in keys],
        tickangle=0,
        title=None,
        tickfont=dict(size=font, color="#21252B"),
        fixedrange=False,
        automargin=True,
    )
    figure.update_yaxes(
        title=plot["axis_title"].iloc[0],
        title_font=dict(size=max(12, font), color="#21252B"),
        tickformat=".2f",
        tickfont=dict(size=max(12, font), color="#21252B"),
        rangemode="tozero",
        range=[0, highest * (1.22 + 0.1 * label_lines)],
        fixedrange=True,
        automargin=True,
    )
    figure.update_layout(
        title=None,
        width=width,
        height=680 if several else 640,
        barmode="group",
        bargap=0.28,
        bargroupgap=0.08,
        autosize=False,
        dragmode="zoom",
        margin=dict(l=72, r=24, t=max(72, 28 + label_lines * (font + 3)), b=16),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Inter, Arial, sans-serif", color="#21252B", size=font),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            x=0,
            font=dict(size=13, color="#21252B"),
            itemsizing="constant",
            tracegroupgap=18,
            itemclick=False,
            itemdoubleclick=False,
        ),
        hoverlabel=dict(font_size=13),
        meta={"longest": _longest_label(plot, several), "count": len(keys), "bars": bars},
    )
    figure.update_xaxes(gridcolor="#E1E4E8", zerolinecolor="#E1E4E8")
    figure.update_yaxes(gridcolor="#E1E4E8", zerolinecolor="#E1E4E8")
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
    plot = bonds.copy()
    if "portfolio" not in plot.columns:
        plot["portfolio"] = ""
    if "label" not in plot.columns:
        plot["label"] = ""
    if "report_date" not in plot.columns:
        plot["report_date"] = report_date
    plot["maturity_text"] = plot["maturity"].map(_format_date)
    plot["nominal_text"] = plot["nominal_kzt"].map(_format_kzt)
    plot["face_text"] = plot["nominal_amount"].map(_format_kzt)
    scale, axis_title = _nominal_axis(plot["nominal_kzt"])
    plot["nominal_scaled"] = plot["nominal_kzt"] / scale
    plot["axis_title"] = axis_title
    plot["days_left"] = [
        (pd.Timestamp(day) - pd.Timestamp(reported)).days
        for day, reported in zip(plot["maturity"], plot["report_date"])
    ]
    plot["currency"] = plot["currency"].fillna("—").astype(str).str.strip().replace({"": "—"})
    plot["ticker"] = plot["ticker"].fillna("—").astype(str)
    plot["short_ticker"] = plot["ticker"].map(lambda ticker: _wrap_ticker(_short_ticker(ticker)))
    plot["portfolio"] = plot["portfolio"].fillna("").astype(str)
    plot["label"] = plot["label"].fillna("").astype(str).replace({"": "Портфель"})
    issues = plot[["maturity", "isin"]].drop_duplicates().sort_values(["maturity", "isin"])
    key_by_issue = {
        (maturity, isin): f"{index:02d}-{isin}"
        for index, (maturity, isin) in enumerate(zip(issues["maturity"], issues["isin"]))
    }
    plot["key"] = [key_by_issue[(maturity, isin)] for maturity, isin in zip(plot["maturity"], plot["isin"])]
    plot["bar_text"] = plot["nominal_kzt"].map(_compact_amount)
    return plot.sort_values(["maturity", "isin", "label"]).reset_index(drop=True)


def _longest_label(plot: pd.DataFrame, several: bool = False) -> int:
    lines = [
        _longest_line(plot["maturity_text"]),
        _longest_line(plot["short_ticker"]),
        _longest_line(plot["bar_text"]),
    ]
    if several:
        lines.append(_longest_line(plot["label"]))
    return max(*lines, 1)


def _longest_line(values) -> int:
    longest = 1
    for value in values.astype(str):
        for line in value.split("<br>"):
            longest = max(longest, len(line))
    return longest


def _label_lines(plot: pd.DataFrame, several: bool = False) -> int:
    """Ticker lines, the amount, and the portfolio label when several funds are shown."""
    ticker_lines = max(str(value).count("<br>") + 1 for value in plot["short_ticker"])
    return ticker_lines + (2 if several else 1)


def _chart_width(plot: pd.DataFrame, zoom: float, bars: int = 1) -> tuple[int, int, int]:
    """Font size, column width, and figure width for this zoom.

    Zoom 1 fills the screen when the labels still fit. A larger zoom widens
    each column. The font never drops below the readable floor; the figure
    gets wider and scrolls instead of letting labels overlap. ``bars`` is the
    number of side-by-side portfolio bars reserved in every column.
    """
    longest = _longest_label(plot, bars > 1)
    count = max(plot["key"].nunique(), 1)
    room = _VIEWPORT_PX - _AXIS_MARGIN_PX
    min_slot = math.ceil(longest * _MIN_FONT * _CHAR_EM + _LABEL_GAP_PX) * bars
    desired = max(min_slot, (room / count) * zoom)
    font = int((desired / bars - _LABEL_GAP_PX) / (longest * _CHAR_EM))
    font = max(_MIN_FONT, min(_MAX_FONT, font))
    text_slot = math.ceil(longest * font * _CHAR_EM + _LABEL_GAP_PX) * bars
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


def _cashflow(bonds: pd.DataFrame, portfolio_labels: list[str]) -> None:
    with st.container(border=True, key="cashflow-block"):
        _block_title("Денежный поток", "#DCA262")
        _cashflow_body(bonds, portfolio_labels)


def _cashflow_body(bonds: pd.DataFrame, portfolio_labels: list[str]) -> None:
    st.caption(
        "Купоны и погашение номинала, в тенге по курсу из шапки файла. "
        "Цвет — валюта выпуска: купон светлее, погашение насыщеннее. "
        "Купон в дату отчёта включён. Полный список выплат открывается у столбца и прокручивается."
    )
    horizon_label = st.radio(
        "Горизонт денежного потока",
        list(HORIZONS),
        index=0,
        horizontal=True,
        key="cash_horizon",
    )
    if bonds.empty:
        st.info("Нет облигаций для денежного потока.")
        return
    if "payments_per_year" not in bonds.columns:
        st.info("Загрузите файлы ещё раз: для купонов нужна колонка «Кол-во купонных выплат в год».")
        return
    events = cash_events(bonds, HORIZONS[horizon_label])
    if events.empty:
        st.info("В выбранном горизонте нет купонов и погашений.")
        return
    events, labels = _chart_filters(events, portfolio_labels, "cashflow")
    if not labels:
        st.info("На графике ничего не выбрано. Отметьте хотя бы один портфель и одну валюту.")
        return
    if events.empty:
        st.info("В выбранном горизонте нет купонов и погашений для отмеченных портфелей и валют.")
        return
    coupons = events.loc[events["kind"] == "купон", "amount_kzt"].sum()
    redemptions = events.loc[events["kind"] == "погашение", "amount_kzt"].sum()
    left, right = st.columns(2)
    left.metric("Купоны, KZT", _compact_amount(coupons))
    right.metric("Погашения, KZT", _compact_amount(redemptions))
    figure = _cashflow_figure(events)
    _show_scrollable_chart(figure, 1.0, int(figure.layout.meta["bars"]))
    _cashflow_table(events)


def _cashflow_table(events: pd.DataFrame) -> None:
    ordered = events.sort_values(
        ["date", "kind", "label", "ticker"],
        ascending=[True, True, True, True],
    )
    view = pd.DataFrame(
        {
            "Дата": pd.to_datetime(ordered["date"]),
            "Портфель": ordered["label"].where(ordered["label"].astype(str).str.len() > 0, ordered["portfolio"]),
            "Тикер": ordered["ticker"],
            "Выплата": ordered["kind"],
            "Валюта": ordered["currency"],
            "Сумма в валюте": ordered["amount_ccy"],
            "Сумма, KZT": ordered["amount_kzt"],
        }
    )
    st.dataframe(
        view,
        hide_index=True,
        width="stretch",
        column_config={
            "Дата": st.column_config.DateColumn(format="DD.MM.YYYY"),
            "Сумма в валюте": st.column_config.NumberColumn(format="%,.0f"),
            "Сумма, KZT": st.column_config.NumberColumn(format="%,.0f"),
        },
    )


def _cashflow_figure(events: pd.DataFrame):
    events = events.copy()
    events["label"] = events["label"].fillna("").astype(str)
    months = _cash_months(events)
    month_label = {month: f"{month.month:02d}.{month.year}" for month in months}
    grouped = events.groupby(["month", "currency", "kind"], as_index=False)["amount_kzt"].sum()
    scale, axis_title = _nominal_axis(events.groupby("month")["amount_kzt"].sum())
    currencies = _currency_order(grouped["currency"])
    colors = _currency_color_map(currencies)
    details = {
        key: _cash_hover(piece)
        for key, piece in events.groupby(["month", "currency", "kind"])
    }
    plot = pd.DataFrame(
        {
            "key": [month_label[month] for month in months],
            "maturity_text": [month_label[month] for month in months],
            "short_ticker": [month_label[month] for month in months],
            "bar_text": [month_label[month] for month in months],
            "label": "",
        }
    )
    font, _slot, width = _chart_width(plot, 1.0, 1)
    figure = go.Figure()
    # Coupons sit under redemptions. Each segment is one currency: lighter for
    # coupons, solid for redemptions. Portfolios are summed into that segment.
    for kind in ("купон", "погашение"):
        for currency in currencies:
            part = grouped.loc[(grouped["currency"] == currency) & (grouped["kind"] == kind)]
            if part.empty:
                continue
            by_month = part.set_index("month")["amount_kzt"]
            is_coupon = kind == "купон"
            figure.add_trace(
                go.Bar(
                    x=[month_label[month] for month in months],
                    y=[None if month not in by_month.index else float(by_month.loc[month]) / scale for month in months],
                    name=f"{currency} · {kind}",
                    legendgroup=f"{currency}|{kind}",
                    marker=dict(
                        color=_tint(colors[currency], 0.42) if is_coupon else colors[currency],
                        line=dict(color="rgba(255,255,255,0.9)", width=1),
                    ),
                    customdata=[
                        f"{currency} · {kind}<br>{_compact_amount(by_month.loc[month])}<br>{details.get((month, currency, kind), '')}"
                        if month in by_month.index
                        else ""
                        for month in months
                    ],
                    hovertemplate="%{x}<br>%{customdata}<extra></extra>",
                    showlegend=True,
                )
            )
    figure.update_layout(
        barmode="stack",
        width=width,
        height=560,
        bargap=0.35,
        autosize=False,
        dragmode="zoom",
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Inter, Arial, sans-serif", color="#21252B", size=font),
        margin=dict(l=72, r=24, t=48, b=16),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            x=0,
            font=dict(size=13, color="#21252B"),
            title=dict(text="Валюта и тип", font=dict(size=13, color="#21252B")),
        ),
        hoverlabel=dict(
            font_size=1,
            bgcolor="rgba(0,0,0,0)",
            bordercolor="rgba(0,0,0,0)",
            font_color="rgba(0,0,0,0)",
        ),
        meta={
            "longest": _longest_label(plot, False),
            "count": len(months),
            "bars": 1,
            "scroll_detail": True,
        },
    )
    figure.update_xaxes(
        type="category",
        categoryorder="array",
        categoryarray=plot["key"].tolist(),
        tickfont=dict(size=font, color="#21252B"),
        gridcolor="#E1E4E8",
        zerolinecolor="#E1E4E8",
        fixedrange=False,
    )
    figure.update_yaxes(
        title=axis_title.replace("Номинал", "Поток"),
        title_font=dict(size=max(12, font), color="#21252B"),
        tickformat=".2f",
        tickfont=dict(size=max(11, font), color="#21252B"),
        gridcolor="#E1E4E8",
        zerolinecolor="#E1E4E8",
        rangemode="tozero",
        fixedrange=True,
    )
    return figure


def _cash_months(events: pd.DataFrame) -> list[date]:
    first = events["month"].min()
    last = events["month"].max()
    span = (last.year - first.year) * 12 + last.month - first.month
    if span > 36:
        return sorted(events["month"].unique())
    months = []
    cursor = first
    while cursor <= last:
        months.append(cursor)
        year = cursor.year + (cursor.month // 12)
        month = cursor.month % 12 + 1
        months_cursor = date(year, month, 1)
        cursor = months_cursor
    return months


def _tint(color: str, toward_white: float) -> str:
    raw = color.lstrip("#")
    red, green, blue = int(raw[0:2], 16), int(raw[2:4], 16), int(raw[4:6], 16)
    red = round(red + (255 - red) * toward_white)
    green = round(green + (255 - green) * toward_white)
    blue = round(blue + (255 - blue) * toward_white)
    return f"#{red:02x}{green:02x}{blue:02x}"


def _cash_hover(piece: pd.DataFrame) -> str:
    lines = []
    ordered = piece.sort_values(["date", "amount_kzt"], ascending=[True, False])
    for row in ordered.itertuples(index=False):
        who = row.label or row.portfolio or ""
        prefix = f"{html.escape(str(who))} · " if who else ""
        paid = pd.Timestamp(row.date).strftime("%d.%m.%Y")
        ticker = html.escape(str(row.ticker))
        lines.append(f"{paid} · {prefix}{ticker}: {_compact_amount(row.amount_kzt)}")
    return "<br>".join(lines)


def _table(bonds: pd.DataFrame, report_date: date) -> None:
    st.subheader("Выпуски в выбранном горизонте")
    if bonds.empty:
        st.caption("Таблица пуста.")
        return
    view = pd.DataFrame(
        {
            "Портфель": bonds["portfolio"] if "portfolio" in bonds.columns else "",
            "Эмитент": bonds["issuer"],
            "Тикер": bonds["ticker"],
            "ISIN": bonds["isin"],
            "Тип": bonds["security_type"],
            "Валюта": bonds["currency"],
            "Дата погашения": pd.to_datetime(bonds["maturity"]),
            "Дней до погашения": [
                (pd.Timestamp(day) - pd.Timestamp(reported)).days
                for day, reported in zip(
                    bonds["maturity"],
                    bonds["report_date"] if "report_date" in bonds.columns else [report_date] * len(bonds),
                )
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
