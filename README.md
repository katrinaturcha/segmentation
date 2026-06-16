# Streamlit-проект: сегментация ТВ-стоек

Приложение загружает Excel с товарами и строит матрицу сегментации по типам товаров и сегментам BASIC / LIGHT / STANDART / HEAVY / HEAVY XL.

## Столбцы во входном файле
Обязательные/ожидаемые столбцы:

- `image_url`
- `sku`
- `image`
- `Type`
- `максимальная диагональ`
- `Diagonal category`
- `максимальная нагрузка кг`
- `Load capacity category kg`
- `максимальная VESA`
- `VESA category`
- `максимальная суммарная нагрузка (с полками) кг`
- `описание`

Дополнительные столбцы не мешают работе.

## Запуск

```bash
cd tv_stands_segmentation_app
pip install -r requirements.txt
streamlit run app.py
```

## Логика

Сегмент определяется по `Load capacity category kg`:

- `30 kg` → BASIC
- `60 kg` → LIGHT
- `70 kg` → STANDART
- `120 kg` → HEAVY
- `150 kg` → HEAVY XL

Если категория нагрузки пустая, приложение пытается определить сегмент по числу из `максимальная нагрузка кг`.
