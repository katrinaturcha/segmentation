import base64
import html
import re
from io import BytesIO
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

import pandas as pd
import requests
import streamlit as st
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

st.set_page_config(page_title="Сегментация ТВ-стоек", layout="wide")

APP_DIR = Path(__file__).resolve().parent
DATA_FILE = APP_DIR / "sample_test_onkron.xlsx"

SEGMENTS = [
    {
        "name": "BASIC",
        "load_label": "30 kg",
        "diagonal": '17"-55"',
        "margin": "12%",
        "vesa": "75x75, 100x100, 200x100, 200x200, 300x200, 300x300, 400x200, 400x300, 400x400",
    },
    {
        "name": "LIGHT",
        "load_label": "60 kg",
        "diagonal": '32"-65"',
        "margin": "20%",
        "vesa": "200x100, 200x200, 300x200, 300x300, 400x200, 400x300, 400x400, 400x500, 600x300, 600x400",
    },
    {
        "name": "STANDART",
        "load_label": "70 kg",
        "diagonal": '40"-75"',
        "margin": "25%",
        "vesa": "300x300, 400x200, 400x300, 400x400, 500x400, 500x500, 600x300, 600x400, Prof 600x500, 600x600, 700x400, 700x500, 700x700, 800x400",
    },
    {
        "name": "HEAVY",
        "load_label": "120 kg",
        "diagonal": '60"-100"',
        "margin": "30%",
        "vesa": "400x400, 500x400, 500x500, 600x300, 600x400, 600x500, 700x700, 800x400, 800x600, Prof 900x600, 1000x600, 1000x800, 1100x600",
    },
    {
        "name": "HEAVY XL",
        "load_label": "150 kg",
        "diagonal": '75"-120"',
        "margin": "40%",
        "vesa": "400x400, 600x500, 600x600, 700x400, 700x500, 700x700, 800x600, 900x600, Prof 900x600, 1000x600, 1000x800, 1500x600",
    },
]

SEGMENT_BY_DIAGONAL = {
    '17"-55"': "BASIC",
    '32"-65"': "LIGHT",
    '40"-75"': "STANDART",
    '60"-100"': "HEAVY",
    '75"-120"': "HEAVY XL",
}

DEFAULT_TYPE_ORDER = [
    "tv stands",
    "design | interior",
    "mobile tv stands",
    "universal aluminum",
    "motorised",
    "professional | touch panel",
    "pro",
]

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


def normalize_diagonal_category(value) -> Optional[str]:
    if pd.isna(value):
        return None

    text = str(value).strip()
    text = text.replace("“", '"').replace("”", '"').replace("″", '"')
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", "", text)

    mapping = {
        '17"-55"': '17"-55"',
        '17-55': '17"-55"',
        '17"55"': '17"-55"',

        '32"-65"': '32"-65"',
        '32-65': '32"-65"',
        '32"65"': '32"-65"',

        '40"-75"': '40"-75"',
        '40-75': '40"-75"',
        '40"75"': '40"-75"',
        '43"-75"': '40"-75"',
        '43-75': '40"-75"',

        '60"-100"': '60"-100"',
        '60-100': '60"-100"',
        '60"100"': '60"-100"',

        '75"-120"': '75"-120"',
        '75-120': '75"-120"',
        '75"120"': '75"-120"',
    }

    return mapping.get(text)


def detect_segment(row: pd.Series) -> str:
    diagonal_category = normalize_diagonal_category(row.get("Diagonal category"))

    if diagonal_category in SEGMENT_BY_DIAGONAL:
        return SEGMENT_BY_DIAGONAL[diagonal_category]

    return "НЕ ОПРЕДЕЛЕНО"


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

    df["Type"] = df["Type"].fillna("без типа").astype(str).str.strip().str.lower()
    df["sku"] = df["sku"].fillna("").astype(str).str.strip()
    df["segment"] = df.apply(detect_segment, axis=1)

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


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24)
def image_to_data_uri(url: str) -> str:
    url = clean_url(url)

    if not url.startswith("http"):
        return ""

    try:
        response = requests.get(
            url,
            timeout=12,
            verify=False,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
                "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
            },
        )

        response.raise_for_status()

        if not response.content:
            return ""

        mime = guess_mime(url, response.headers.get("Content-Type", ""))
        encoded = base64.b64encode(response.content).decode("ascii")

        return f"data:{mime};base64,{encoded}"

    except Exception:
        return ""


def product_tile_html(row: pd.Series) -> str:
    sku = html.escape(str(row.get("sku", "") or "без sku"))

    diagonal = html.escape(safe_text(row.get("максимальная диагональ")))
    load = html.escape(safe_text(row.get("максимальная нагрузка кг")))
    vesa = html.escape(safe_text(row.get("максимальная VESA")))

    tooltip = html.escape(
        f"максимальная диагональ: {diagonal}\n"
        f"максимальная нагрузка кг: {load}\n"
        f"максимальная VESA: {vesa}",
        quote=True,
    )

    img_url = clean_url(row.get("image")) or clean_url(row.get("image_url"))
    product_link = clean_url(row.get("image_url")) or img_url
    data_uri = image_to_data_uri(img_url) if img_url else ""

    if data_uri:
        image_html = f'<img class="product-img" src="{data_uri}" loading="lazy" />'
    else:
        image_html = '<div class="product-img product-img-empty">нет фото</div>'

    if product_link.startswith("http"):
        sku_html = (
            f'<a class="sku-label" '
            f'href="{html.escape(product_link, quote=True)}" '
            f'target="_blank" '
            f'title="{tooltip}">{sku}</a>'
        )
    else:
        sku_html = f'<span class="sku-label" title="{tooltip}">{sku}</span>'

    return f'<div class="product-tile">{image_html}<div>{sku_html}</div></div>'


def heat_class(count: int, max_count: int) -> str:
    if count <= 0 or max_count <= 0:
        return "heat-0"

    ratio = count / max_count

    if ratio <= 0.2:
        return "heat-1"
    if ratio <= 0.4:
        return "heat-2"
    if ratio <= 0.6:
        return "heat-3"
    if ratio <= 0.8:
        return "heat-4"

    return "heat-5"


def render_matrix(df: pd.DataFrame) -> None:
    type_values = [t for t in DEFAULT_TYPE_ORDER if t in set(df["Type"])]
    type_values += sorted([t for t in df["Type"].dropna().unique() if t not in type_values])

    counts = [
        len(df[(df["Type"] == type_name) & (df["segment"] == s["name"])])
        for type_name in type_values
        for s in SEGMENTS
    ]

    max_count = max(counts) if counts else 0

    html_parts = ["<div class='matrix-wrap'><table class='matrix'>"]

    html_parts.append("<tr><th class='black-head'>СЕГМЕНТАЦИЯ</th>")

    for s in SEGMENTS:
        html_parts.append(f"<th class='segment-head'>{s['name']}</th>")

    html_parts.append("</tr>")

    html_parts.append("<tr><td class='left-title'>МАКС. НАГРУЗКА</td>")

    for s in SEGMENTS:
        html_parts.append(f"<td class='top-cell'><b>{s['load_label']}</b></td>")

    html_parts.append("</tr>")

    html_parts.append("<tr><td class='left-title'>VESA</td>")

    for s in SEGMENTS:
        html_parts.append(f"<td class='vesa-cell'>{html.escape(s['vesa'])}</td>")

    html_parts.append("</tr>")

    html_parts.append("<tr><td class='left-title'>РАЗМЕР ЭКРАНОВ</td>")

    for s in SEGMENTS:
        html_parts.append(f"<td class='top-cell'><b>{html.escape(s['diagonal'])}</b></td>")

    html_parts.append("</tr>")

    for type_name in type_values:
        html_parts.append(f"<tr><td class='type-cell'>{html.escape(type_name)}</td>")

        for s in SEGMENTS:
            cell_df = df[
                (df["Type"] == type_name)
                & (df["segment"] == s["name"])
            ]

            count = len(cell_df)
            cls = heat_class(count, max_count)

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

    for s in SEGMENTS:
        html_parts.append(f"<td class='margin-cell'>{s['margin']}</td>")

    html_parts.append("</tr></table></div>")

    st.markdown("".join(html_parts), unsafe_allow_html=True)


st.markdown(
    """
    <style>
    .title-block h1 {font-size: 34px; line-height: 0.95; margin-bottom: 0; color: #10243a;}
    .title-block h3 {font-size: 18px; margin-top: 8px; color: #10243a;}

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

    .heat-0 {background: #ffffff;}
    .heat-1 {background: #e9fbfb;}
    .heat-2 {background: #caf4f3;}
    .heat-3 {background: #95e8e7;}
    .heat-4 {background: #5ed9d8;}
    .heat-5 {background: #22c5c3;}

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
    }

    img.product-img {
        display: block;
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

summary = (
    df.pivot_table(index="Type", columns="segment", values="sku", aggfunc="count", fill_value=0)
    .reindex(columns=[s["name"] for s in SEGMENTS], fill_value=0)
    .reset_index()
)

col1, col2, col3 = st.columns(3)

col1.metric("Всего SKU", len(df))
col2.metric("Типов", df["Type"].nunique())
col3.metric("Не определено", int((df["segment"] == "НЕ ОПРЕДЕЛЕНО").sum()))

render_matrix(df)

if show_table:
    st.subheader("Исходные данные")
    st.dataframe(df, use_container_width=True)

buffer = BytesIO()
summary.to_excel(buffer, index=False)

st.download_button(
    "Скачать сводную таблицу Excel",
    data=buffer.getvalue(),
    file_name="segmentation_summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)