import base64
import html
import re
from io import BytesIO
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Сегментация ТВ-стоек", layout="wide")

APP_DIR = Path(__file__).resolve().parent
DATA_FILE = APP_DIR / "sample_test_onkron.xlsx"

SEGMENTS = [
    {
        "name": "BASIC",
        "load_label": "35 kg",
        "diagonal": '17"-60"',
        "margin": "12%",
        "vesa": "400x400",
    },
    {
        "name": "LIGHT",
        "load_label": "60 kg",
        "diagonal": '32"-65"',
        "margin": "20%",
        "vesa": "600x400",
    },
    {
        "name": "STANDART",
        "load_label": "70 kg",
        "diagonal": '40"-75"',
        "margin": "25%",
        "vesa": "800x400",
    },
    {
        "name": "MEDIUM",
        "load_label": "100 kg",
        "diagonal": '50"-90"',
        "margin": "28%",
        "vesa": "900x600",
    },
    {
        "name": "HEAVY",
        "load_label": "120 kg",
        "diagonal": '60"-100"',
        "margin": "30%",
        "vesa": "1100x600",
    },
    {
        "name": "HEAVY XL",
        "load_label": "150 kg",
        "diagonal": '75"-120"',
        "margin": "40%",
        "vesa": "1500x600",
    },
]

SEGMENT_ORDER = {
    "BASIC": 1,
    "LIGHT": 2,
    "STANDART": 3,
    "MEDIUM": 4,
    "HEAVY": 5,
    "HEAVY XL": 6,
}

SEGMENT_BY_DIAGONAL = {
    '17"-60"': "BASIC",
    '32"-65"': "LIGHT",
    '40"-75"': "STANDART",
    '50"-90"': "MEDIUM",
    '60"-100"': "HEAVY",
    '75"-120"': "HEAVY XL",
}

SEGMENT_BY_LOAD = {
    35: "BASIC",
    60: "LIGHT",
    70: "STANDART",
    100: "MEDIUM",
    120: "HEAVY",
    150: "HEAVY XL",
}

REQUIRED_COLUMNS = [
    "image_url",
    "sku",
    "image",
    "Type",
    "максимальная диагональ",
    "Diagonal category",
    "максимальная нагрузка кг",
    "Load capacity category kg",
    "максимальная VESA",
    "VESA category",
    "максимальная суммарная нагрузка (с полками) кг",
    "описание",
]

SERIES_COLUMN = "serie"
TYPE_ORDER = [
    "tv stands",
    "mobile stands",
    "motorised",
    "design | interior",
    "touch panel",
    "universal alluminum",
]


def normalize_type(value) -> str:
    if pd.isna(value):
        return "без типа"

    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)

    if text in ["professional | touch panel", "professional touch panel"]:
        return "touch panel"

    if text in ["pro", "prof", "professional"]:
        return "pro"

    if text in ["touchpanel", "touch-panel", "touch panel"]:
        return "touch panel"

    if text in ["mobile tv stands", "mobile stand", "mobile stands"]:
        return "mobile stands"

    return text


def normalize_series(value) -> str:
    if pd.isna(value) or not str(value).strip():
        return "Все товары"

    return str(value).strip()


def normalize_diagonal_category(value) -> Optional[str]:
    if pd.isna(value):
        return None

    text = str(value).strip()
    text = text.replace("“", '"').replace("”", '"').replace("″", '"')
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", "", text)

    mapping = {
        '17"-60"': '17"-60"',
        "17-60": '17"-60"',
        '17"60"': '17"-60"',
        '32"-65"': '32"-65"',
        "32-65": '32"-65"',
        '32"65"': '32"-65"',
        '40"-75"': '40"-75"',
        "40-75": '40"-75"',
        '40"75"': '40"-75"',
        '43"-75"': '40"-75"',
        "43-75": '40"-75"',
        '50"-90"': '50"-90"',
        "50-90": '50"-90"',
        '50"90"': '50"-90"',
        '60"-100"': '60"-100"',
        "60-100": '60"-100"',
        '60"100"': '60"-100"',
        '75"-120"': '75"-120"',
        "75-120": '75"-120"',
        '75"120"': '75"-120"',
    }

    return mapping.get(text)


def extract_number(value) -> Optional[float]:
    if pd.isna(value):
        return None

    text = str(value).replace(",", ".")
    match = re.search(r"\d+(?:\.\d+)?", text)

    return float(match.group()) if match else None


def normalize_load_category(value) -> Optional[int]:
    number = extract_number(value)

    if number is None:
        return None

    for max_load in [35, 60, 70, 100, 120, 150]:
        if number <= max_load:
            return max_load

    return 150


def detect_diagonal_segment(row: pd.Series) -> str:
    diagonal_category = normalize_diagonal_category(row.get("Diagonal category"))

    if diagonal_category in SEGMENT_BY_DIAGONAL:
        return SEGMENT_BY_DIAGONAL[diagonal_category]

    return "НЕ ОПРЕДЕЛЕНО"


def detect_load_segment(row: pd.Series) -> str:
    # Размещение определяется только категорией нагрузки из Excel.
    load_category = normalize_load_category(row.get("Load capacity category kg"))

    if load_category in SEGMENT_BY_LOAD:
        return SEGMENT_BY_LOAD[load_category]

    return "НЕ ОПРЕДЕЛЕНО"


def detect_load_status(row: pd.Series) -> str:
    diagonal_segment = row.get("diagonal_segment")
    load_segment = row.get("load_segment")

    if diagonal_segment == "НЕ ОПРЕДЕЛЕНО" or load_segment == "НЕ ОПРЕДЕЛЕНО":
        return "unknown"

    diagonal_rank = SEGMENT_ORDER.get(diagonal_segment)
    load_rank = SEGMENT_ORDER.get(load_segment)

    if diagonal_rank is None or load_rank is None:
        return "unknown"

    if load_rank < diagonal_rank:
        return "low"

    if load_rank > diagonal_rank:
        return "high"

    return "ok"


def build_final_segment(row: pd.Series) -> str:
    diagonal_segment = row.get("diagonal_segment")
    status = row.get("load_status")

    if diagonal_segment == "НЕ ОПРЕДЕЛЕНО":
        return "НЕ ОПРЕДЕЛЕНО"

    if status == "ok":
        return diagonal_segment

    if status == "low":
        return f"{diagonal_segment} LOW LOAD"

    if status == "high":
        return f"{diagonal_segment} HIGH LOAD"

    return f"{diagonal_segment} / нагрузка не определена"


@st.cache_data(show_spinner=False)
def prepare_df(file_path: str, file_mtime: float) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        st.warning("В файле нет части ожидаемых столбцов: " + ", ".join(missing))

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = None

    if SERIES_COLUMN not in df.columns:
        df[SERIES_COLUMN] = "Все товары"

    df[SERIES_COLUMN] = df[SERIES_COLUMN].apply(normalize_series)
    df["Type"] = df["Type"].apply(normalize_type)
    df["sku"] = df["sku"].fillna("").astype(str).str.strip()

    df["diagonal_segment"] = df.apply(detect_diagonal_segment, axis=1)
    df["load_segment"] = df.apply(detect_load_segment, axis=1)
    df["load_status"] = df.apply(detect_load_status, axis=1)
    df["final_segment"] = df.apply(build_final_segment, axis=1)
    # Основное размещение в матрице определяется категорией нагрузки.
    # Диагональ используется только для статуса соответствия нагрузки.
    df["segment"] = df["load_segment"]

    return df


def clean_url(value) -> str:
    if pd.isna(value):
        return ""

    url = str(value).strip()

    if not url or url.lower() == "nan":
        return ""

    if url.startswith("//"):
        url = "https:" + url
    elif url.startswith("www"):
        url = "https://" + url

    return url


def safe_text(value) -> str:
    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text.lower() == "nan":
        return ""

    return text


def guess_mime(url: str, content_type: str = "") -> str:
    content_type = (content_type or "").split(";")[0].strip().lower()

    if content_type.startswith("image/"):
        return content_type

    path = urlparse(url).path.lower()

    if path.endswith(".png"):
        return "image/png"
    if path.endswith(".webp"):
        return "image/webp"
    if path.endswith(".gif"):
        return "image/gif"

    return "image/jpeg"


def find_local_image(source: str) -> Optional[Path]:
    """Return an image in the project folder, accepting a name without extension."""
    try:
        candidate = Path(source)
        if candidate.is_absolute():
            return None

        image_path = (APP_DIR / candidate).resolve()
        if APP_DIR.resolve() not in image_path.parents:
            return None

        paths = [image_path]
        if not image_path.suffix:
            paths.extend(image_path.with_suffix(ext) for ext in (".png", ".jpg", ".jpeg", ".webp", ".gif"))

        return next((path for path in paths if path.is_file()), None)
    except (OSError, ValueError):
        return None


@st.cache_data(show_spinner=False)
def local_image_to_data_uri(path_text: str, modified_at: int) -> str:
    """Embed a local image so it can be shown in HTML generated by Streamlit."""
    try:
        path = Path(path_text)
        encoded = base64.b64encode(path.read_bytes()).decode("ascii")
        return f"data:{guess_mime(path.name)};base64,{encoded}"
    except OSError:
        return ""


def product_tile_html(row: pd.Series) -> str:
    sku = html.escape(str(row.get("sku", "") or ""))

    diagonal = html.escape(safe_text(row.get("максимальная диагональ")))
    load = html.escape(safe_text(row.get("максимальная нагрузка кг")))
    load_category = html.escape(safe_text(row.get("Load capacity category kg")))
    vesa = html.escape(safe_text(row.get("максимальная VESA")))

    diagonal_segment = html.escape(safe_text(row.get("diagonal_segment")))
    load_segment = html.escape(safe_text(row.get("load_segment")))
    final_segment = html.escape(safe_text(row.get("final_segment")))

    load_status = safe_text(row.get("load_status"))
    risk_class = f"risk-{load_status}"

    tooltip = html.escape(
        f"Диагональный сегмент: {diagonal_segment}\n"
        f"Нагрузочный сегмент: {load_segment}\n"
        f"Итог: {final_segment}\n"
        f"Максимальная диагональ: {diagonal}\n"
        f"Максимальная нагрузка кг: {load}\n"
        f"Категория нагрузки: {load_category}\n"
        f"Максимальная VESA: {vesa}",
        quote=True,
    )

    img_url = clean_url(row.get("image")) or clean_url(row.get("image_url"))
    product_link = clean_url(row.get("image_url")) or img_url
    if img_url.startswith("http"):
        image_html = (
            f'<img class="product-img {risk_class}" '
            f'src="{html.escape(img_url, quote=True)}" '
            f'loading="lazy" referrerpolicy="no-referrer" />'
        )
    else:
        local_image = find_local_image(img_url)
        if local_image:
            data_uri = local_image_to_data_uri(
                str(local_image), local_image.stat().st_mtime_ns
            )
            image_html = (
                f'<img class="product-img {risk_class}" src="{data_uri}" loading="lazy" />'
                if data_uri
                else f'<div class="product-img product-img-empty {risk_class}">нет фото</div>'
            )
        else:
            image_html = f'<div class="product-img product-img-empty {risk_class}">нет фото</div>'

    if product_link.startswith("http"):
        sku_html = (
            f'<a class="sku-label {risk_class}" '
            f'href="{html.escape(product_link, quote=True)}" '
            f'target="_blank" '
            f'title="{tooltip}">{sku}</a>'
        )
    else:
        sku_html = f'<span class="sku-label {risk_class}" title="{tooltip}">{sku}</span>'

    return f'<div class="product-tile">{image_html}<div>{sku_html}</div></div>'


def cell_status_class(cell_df: pd.DataFrame) -> str:
    if cell_df.empty:
        return "cell-empty"

    statuses = set(cell_df["load_status"].dropna().astype(str))

    if "high" in statuses:
        return "cell-high"

    if "low" in statuses:
        return "cell-low"

    if "unknown" in statuses:
        return "cell-unknown"

    return "cell-ok"


def active_segments(df: pd.DataFrame) -> list[dict]:
    """Show every configured load category, including an empty one."""
    return SEGMENTS


def render_matrix(df: pd.DataFrame, segments: list[dict]) -> None:
    available_types = list(df["Type"].dropna().unique())
    type_values = [type_name for type_name in TYPE_ORDER if type_name in available_types]
    type_values.extend(type_name for type_name in available_types if type_name not in type_values)

    html_parts = ["<div class='matrix-wrap'><table class='matrix'>"]

    html_parts.append("<tr><th class='black-head'>СЕГМЕНТАЦИЯ</th>")

    for s in segments:
        html_parts.append(f"<th class='segment-head'>{s['name']}</th>")

    html_parts.append("</tr>")

    html_parts.append("<tr><td class='left-title'>МАКС.<br>НАГРУЗКА</td>")

    for s in segments:
        html_parts.append(f"<td class='top-cell'><b>{s['load_label']}</b></td>")

    html_parts.append("</tr>")

    html_parts.append("<tr><td class='left-title'>VESA</td>")

    for s in segments:
        html_parts.append(f"<td class='vesa-cell'>{html.escape(s['vesa'])}</td>")

    html_parts.append("</tr>")

    html_parts.append("<tr><td class='left-title'>РАЗМЕР<br>ЭКРАНОВ</td>")

    for s in segments:
        html_parts.append(f"<td class='top-cell'><b>{html.escape(s['diagonal'])}</b></td>")

    html_parts.append("</tr>")

    for type_name in type_values:
        html_parts.append(f"<tr><td class='type-cell'>{html.escape(type_name)}</td>")

        for s in segments:
            cell_df = df[
                (df["Type"] == type_name)
                & (df["segment"] == s["name"])
            ]

            count = len(cell_df)
            cls = cell_status_class(cell_df)

            products = "".join(
                product_tile_html(row)
                for _, row in cell_df.iterrows()
            )

            content = (
                f"<div class='count'>{count}</div>"
                f"<div class='products-grid'>{products}</div>"
            )

            html_parts.append(f"<td class='data-cell {cls}'>{content}</td>")

        html_parts.append("</tr>")

    html_parts.append("<tr><td class='margin-title'>МАРЖИНАЛЬНОСТЬ</td>")

    for s in segments:
        html_parts.append(f"<td class='margin-cell'>{s['margin']}</td>")

    html_parts.append("</tr></table></div>")

    st.markdown("".join(html_parts), unsafe_allow_html=True)


st.markdown(
    """
    <style>
    .title-block h1 {font-size: 34px; line-height: 0.95; margin-bottom: 0; color: #10243a;}
    .title-block h3 {font-size: 18px; margin-top: 8px; color: #10243a;}

    [data-testid="stMain"] div.stButton > button {
        min-width: 230px;
        min-height: 82px;
        padding: 14px 28px;
        border: 3px solid #34c8c6;
        border-radius: 10px;
        background: #fff;
        color: #34c8c6;
        font-size: 27px;
        font-weight: 800;
        letter-spacing: 0.02em;
    }

    [data-testid="stMain"] div.stButton > button:hover {
        border-color: #159b99;
        background: #e9fbfa;
        color: #159b99;
    }

    [data-testid="stMain"] div.stButton > button[kind="primary"] {
        background: #34c8c6;
        color: #fff;
    }

    .legend {
        display: flex;
        gap: 14px;
        align-items: center;
        margin: 10px 0 16px 0;
        font-size: 13px;
        font-weight: 700;
    }

    .legend-item {
        display: flex;
        align-items: center;
        gap: 6px;
    }

    .legend-box {
        width: 16px;
        height: 16px;
        border-radius: 4px;
        display: inline-block;
    }

    .legend-low {background: #fff3cd; border: 2px solid #f0ad00;}
    .legend-ok {background: #e7f8ee; border: 2px solid #2ca25f;}
    .legend-high {background: #fde2e2; border: 2px solid #d93025;}
    .legend-unknown {background: #eeeeee; border: 2px solid #999;}

    .matrix-wrap {overflow-x: auto; padding-bottom: 12px;}

    table.matrix {
        border-collapse: collapse;
        width: 100%;
        min-width: 1180px;
        font-family: Arial, sans-serif;
        table-layout: fixed;
    }

    .matrix th,
    .matrix td {
        border: 1px solid #333;
        text-align: center;
        vertical-align: middle;
        padding: 10px;
    }

    .black-head {
        background: #1f1f1f;
        color: #fff;
        width: 175px;
        font-size: 14px;
        white-space: nowrap;
    }

    .segment-head {
        background: #9c9c9c;
        color: #fff;
        font-size: 14px;
        height: 38px;
    }

    .segment-head:nth-child(4) {background: #838383;}
    .segment-head:nth-child(5) {background: #707070;}
    .segment-head:nth-child(6) {background: #5f5f5f;}

    .left-title {
        font-weight: 800;
        background: #f4f4f4;
        width: 175px;
    }

    .top-cell {
        background: #fafafa;
        height: 58px;
        font-size: 13px;
    }

    .vesa-cell {
        background: #f7f7f7;
        font-size: 10px;
        line-height: 1.25;
        height: 105px;
    }

    .type-cell {
        font-weight: 700;
        text-align: right !important;
        background: #fff;
        font-size: 13px;
    }

    .data-cell {
        height: auto;
        min-height: 135px;
        border-style: dashed !important;
        font-size: 12px;
        transition: 0.15s;
        vertical-align: top !important;
    }

    .cell-empty {background: #ffffff !important;}
    .cell-ok {background: #e7f8ee !important;}
    .cell-low {background: #fff3cd !important;}
    .cell-high {background: #fde2e2 !important;}
    .cell-unknown {background: #eeeeee !important;}

    .count {
        font-size: 20px;
        color: #111;
        text-decoration: none !important;
        font-weight: 800;
        margin-bottom: 8px;
    }

    .products-grid {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        justify-content: center;
        align-items: flex-start;
    }

    .product-tile {
        width: 90px;
        text-align: center;
        line-height: 1.15;
        overflow-wrap: anywhere;
    }

    .product-img {
        display: flex;
        align-items: center;
        justify-content: center;
        width: 76px;
        height: 62px;
        object-fit: contain;
        margin: 0 auto 5px auto;
        background: rgba(255,255,255,0.78);
        border-radius: 5px;
        font-size: 9px;
        color: #777;
        box-sizing: border-box;
    }

    img.product-img {
        display: block;
    }

    .product-img-empty {
        font-size: 9px;
    }

    .sku-label,
    .sku-label:visited,
    .sku-label:hover,
    .sku-label:active {
        font-size: 13px;
        line-height: 1.2;
        color: #111;
        text-decoration: none !important;
        font-weight: 800;
        word-break: break-word;
        overflow-wrap: anywhere;
        white-space: normal;
        cursor: pointer;
    }

    .risk-ok {
        border: 2px solid #2ca25f !important;
        border-radius: 6px;
    }

    .risk-low {
        border: 3px solid #f0ad00 !important;
        border-radius: 6px;
    }

    .risk-high {
        border: 3px solid #d93025 !important;
        border-radius: 6px;
    }

    .risk-unknown {
        border: 2px solid #999 !important;
        border-radius: 6px;
    }

    .sku-label.risk-ok,
    .sku-label.risk-low,
    .sku-label.risk-high,
    .sku-label.risk-unknown {
        border: none !important;
    }

    .sku-label.risk-low {
        color: #7a5200 !important;
    }

    .sku-label.risk-high {
        color: #b00020 !important;
    }

    .margin-title,
    .margin-cell {
        background: #34c8c6;
        color: #fff;
        font-weight: 800;
        font-size: 12px;
        white-space: nowrap;
    }
    </style>

    <div class="title-block">
      <h1>СЕГМЕНТАЦИЯ<br>ТВ-СТОЕК</h1>
      <h3>ПО НАГРУЗКЕ, VESA, ДИАГОНАЛИ</h3>
    </div>

    <div class="legend">
        <div class="legend-item"><span class="legend-box legend-low"></span> нагрузка ниже диагонали</div>
        <div class="legend-item"><span class="legend-box legend-ok"></span> нагрузка соответствует диагонали</div>
        <div class="legend-item"><span class="legend-box legend-high"></span> нагрузка выше диагонали</div>
        <div class="legend-item"><span class="legend-box legend-unknown"></span> нагрузка не определена</div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Настройки")
    st.caption(f"Файл данных: {DATA_FILE.name}")

    if st.button("Обновить данные"):
        st.cache_data.clear()
        st.rerun()

    show_table = st.checkbox("Показать исходную таблицу", value=False)

if not DATA_FILE.exists():
    st.error(
        f"Не найден файл данных: {DATA_FILE}. "
        f"Положите Excel в папку проекта и назовите его sample_test_onkron.xlsx"
    )
    st.stop()

with st.spinner("Загружаю файл и картинки..."):
    df = prepare_df(str(DATA_FILE), DATA_FILE.stat().st_mtime)

segments_for_export = active_segments(df)
summary = (
    df.pivot_table(
        index=[SERIES_COLUMN, "Type"],
        columns="segment",
        values="sku",
        aggfunc="count",
        fill_value=0,
    )
    .reindex(columns=[s["name"] for s in segments_for_export], fill_value=0)
    .reset_index()
)

series_values = list(df[SERIES_COLUMN].dropna().unique())
if "selected_series" not in st.session_state or st.session_state.selected_series not in series_values:
    st.session_state.selected_series = series_values[0]

with st.container(horizontal=True, horizontal_alignment="left"):
    for series_name in series_values:
        if st.button(
            series_name,
            key=f"series_button_{series_name}",
            type="primary" if series_name == st.session_state.selected_series else "secondary",
        ):
            st.session_state.selected_series = series_name

series_df = df[df[SERIES_COLUMN] == st.session_state.selected_series]
series_segments = active_segments(series_df)

col1, col2, col3 = st.columns(3)
col4, col5, col6 = st.columns(3)

col1.metric("Всего SKU", len(series_df))
col2.metric("Типов", series_df["Type"].nunique())
col3.metric("Не определено", int((series_df["segment"] == "НЕ ОПРЕДЕЛЕНО").sum()))
col4.metric("Нагрузка ниже диагонали", int((series_df["load_status"] == "low").sum()))
col5.metric("Нагрузка соответствует", int((series_df["load_status"] == "ok").sum()))
col6.metric("Нагрузка выше диагонали", int((series_df["load_status"] == "high").sum()))

render_matrix(series_df, series_segments)

if show_table:
    st.subheader("Исходные данные")
    st.dataframe(df, width="stretch")

buffer = BytesIO()
summary.to_excel(buffer, index=False)

st.download_button(
    "Скачать сводную таблицу Excel",
    data=buffer.getvalue(),
    file_name="segmentation_summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
