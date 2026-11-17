"""Задания 4–7. Проверка данных, вывод таблицы, извлечение столбцов варианта, явные типы."""
import re
import warnings

import numpy as np
import pandas as pd

EXCEL_FIRST_DATA_ROW = 3  # строка 1 — «Вариант N», строка 2 — заголовки столбцов


def check_data_quality(df: pd.DataFrame, numeric_cols=None, row_offset: int = EXCEL_FIRST_DATA_ROW,
                       warn: bool = True) -> dict:
    """
    Задание 4. Ищет пустые ячейки, некорректные (нечисловые, бесконечные) значения и
    несовпадение типов в числовых по смыслу столбцах. При нахождении выдаёт предупреждение.
    Номера строк приводятся в нумерации листа Excel (row_offset — номер строки первого значения).
    """
    if numeric_cols is None:
        numeric_cols = list(df.columns)
    report = {"shape": list(df.shape), "dtypes": df.dtypes.astype(str).to_dict(),
              "missing": {}, "non_numeric": {}, "inf": {}, "type_mismatch": []}
    for col in numeric_cols:
        s = df[col]
        empty = s.isna()
        conv = pd.to_numeric(s, errors="coerce")
        bad = conv.isna() & ~empty
        inf = np.isinf(conv.fillna(0).to_numpy(dtype=float))
        rows = lambda m: (np.flatnonzero(np.asarray(m)) + row_offset).tolist()
        report["missing"][col] = {"count": int(empty.sum()), "rows": rows(empty)}
        if bad.any():
            report["non_numeric"][col] = {"rows": rows(bad), "values": s[bad].astype(str).tolist()[:10]}
        if inf.any():
            report["inf"][col] = rows(inf)
        if not pd.api.types.is_numeric_dtype(s):
            report["type_mismatch"].append(col)
        if warn:
            if empty.any():
                warnings.warn(f"Столбец {col}: пустые ячейки ({int(empty.sum())}) в строках {rows(empty)}")
            if bad.any():
                warnings.warn(f"Столбец {col}: некорректные значения {s[bad].tolist()} в строках {rows(bad)}")
            if inf.any():
                warnings.warn(f"Столбец {col}: бесконечные значения в строках {rows(inf)}")
            if not pd.api.types.is_numeric_dtype(s):
                warnings.warn(f"Столбец {col}: несовпадение типов — тип {s.dtype} вместо числового")
    return report


def scan_sheet(raw: pd.DataFrame) -> list:
    """Задание 4 для всего листа: проблемы в каждом из 24 вариантов (без предупреждений)."""
    out = []
    for v_name, cols in variant_columns(raw).items():
        sub = _variant_block(raw, cols)
        rep = check_data_quality(sub, warn=False)
        out.append({"variant": v_name, "rows": len(sub),
                    "missing": sum(m["count"] for m in rep["missing"].values()),
                    "non_numeric": sum(len(v["rows"]) for v in rep["non_numeric"].values()),
                    "values": sorted({x for v in rep["non_numeric"].values() for x in v["values"]})})
    return out


def show_table(df: pd.DataFrame, max_rows=None) -> None:
    """Задание 5. Выводит содержимое таблицы (по умолчанию — полностью), размер и типы."""
    with pd.option_context("display.max_rows", max_rows, "display.max_columns", None,
                           "display.width", 250):
        print(df)
    print(f"\nРазмер: {df.shape}")


def variant_columns(raw: pd.DataFrame) -> dict:
    """Номера столбцов каждого варианта по строке «Вариант N» (первая строка листа)."""
    res = {}
    for j, v in enumerate(raw.iloc[0].tolist()):
        m = re.fullmatch(r"\s*Вариант\s*(\d+)\s*", str(v))
        if m:
            res[int(m.group(1))] = [j, j + 1]
    return res


def _variant_block(raw: pd.DataFrame, cols) -> pd.DataFrame:
    block = raw.iloc[1:, cols]
    block.columns = [str(c) for c in block.iloc[0]]   # заголовки столбцов (строка 2 листа)
    block = block.iloc[1:]                             # строка заголовка таблицы отброшена
    last = block.dropna(how="all").index.max()         # конец данных варианта — автоматически
    return block.loc[:last].reset_index(drop=True)


def extract_variant(raw: pd.DataFrame, variant: int) -> pd.DataFrame:
    """
    Задание 6. Извлекает столбцы варианта: отбрасывает строку «Вариант N» (заголовок таблицы)
    и данные остальных вариантов, сохраняет заголовки столбцов. Возвращает копию —
    исходный DataFrame и файл не изменяются.
    """
    cols = variant_columns(raw)[variant]
    return _variant_block(raw, cols).copy()


def cast_types(df: pd.DataFrame, type_map: dict) -> pd.DataFrame:
    """Задание 7. Явно задаёт типы столбцов; type_map: {'col': 'float64', ...}."""
    df = df.copy()
    for col, dtype in type_map.items():
        df[col] = df[col].astype(dtype)
    return df
