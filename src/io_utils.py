"""Задания 1–3. Чтение и запись табличных данных, сохранение изображений."""
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.io import loadmat, savemat

from src.config import save_figure, timestamped_filename  # noqa: F401 (save_figure — задание 3)


def _has_header(first_row: pd.Series) -> bool:
    """Первая строка считается заголовком, если в ней есть хотя бы одно нечисловое значение."""
    values = first_row.dropna()
    return len(values) > 0 and pd.to_numeric(values, errors="coerce").isna().any()


def read_table(path: str, header="auto") -> pd.DataFrame:
    """
    Задание 1. Универсальное чтение файлов .xlsx, .csv, .txt, .mat.
    header="auto" — наличие заголовков определяется по первой строке; если заголовков
    нет, столбцы автоматически получают имена col_1, col_2, ...
    """
    p = Path(path)
    ext = p.suffix.lower()
    if ext == ".mat":
        mat = loadmat(p)
        arrays = {k: v for k, v in mat.items() if not k.startswith("__")}
        if "data" in arrays:
            df = pd.DataFrame(arrays["data"])
            if "columns" in arrays:
                names = [str(np.ravel(c)[0]).strip() for c in np.ravel(arrays["columns"])]
                if len(names) == df.shape[1]:
                    df.columns = names
        else:  # набор одномерных переменных -> столбцы
            df = pd.DataFrame({k: np.ravel(v) for k, v in arrays.items()})
        has_names = not all(isinstance(c, (int, np.integer)) for c in df.columns)
    else:
        if ext == ".xlsx":
            raw = pd.read_excel(p, header=None)
        elif ext == ".csv":
            raw = pd.read_csv(p, header=None)
        elif ext == ".txt":  # разделитель — табуляция или пробелы
            raw = pd.read_csv(p, header=None, sep=r"\s+", engine="python")
        else:
            raise ValueError(f"Неподдерживаемый формат: {ext}")
        use_header = _has_header(raw.iloc[0]) if header == "auto" else bool(header)
        if use_header:
            df = raw.iloc[1:].reset_index(drop=True)
            df.columns = raw.iloc[0].astype(str).tolist()
            for c in df.columns:  # строки-числа после отделения заголовка -> числа
                conv = pd.to_numeric(df[c], errors="coerce")
                if conv.notna().sum() == df[c].notna().sum():
                    df[c] = conv
        else:
            df = raw
        has_names = use_header

    # Автоматическое именование столбцов, если заголовков в файле не было
    if not has_names:
        df.columns = [f"col_{i + 1}" for i in range(df.shape[1])]
    return df


def write_table(df: pd.DataFrame, base_name: str) -> dict:
    """Задание 2. Сохраняет DataFrame в .xlsx, .csv, .txt и .mat с заголовками и меткой времени."""
    paths = {}
    paths["xlsx"] = timestamped_filename(base_name, "xlsx")
    df.to_excel(paths["xlsx"], index=False, engine="openpyxl")

    paths["csv"] = timestamped_filename(base_name, "csv")
    df.to_csv(paths["csv"], index=False)

    paths["txt"] = timestamped_filename(base_name, "txt")
    df.to_csv(paths["txt"], sep="\t", index=False)

    paths["mat"] = timestamped_filename(base_name, "mat")
    savemat(paths["mat"], {"data": df.to_numpy(dtype=np.float64),
                           "columns": np.array(df.columns, dtype=object)})
    return paths


def write_without_header(df: pd.DataFrame, base_name: str) -> dict:
    """Вспомогательная запись без заголовков — для проверки автоименования (задание 1)."""
    paths = {"csv": timestamped_filename(base_name, "csv"), "txt": timestamped_filename(base_name, "txt"),
             "xlsx": timestamped_filename(base_name, "xlsx"), "mat": timestamped_filename(base_name, "mat")}
    df.to_csv(paths["csv"], index=False, header=False)
    df.to_csv(paths["txt"], sep="\t", index=False, header=False)
    df.to_excel(paths["xlsx"], index=False, header=False, engine="openpyxl")
    savemat(paths["mat"], {"data": df.to_numpy(dtype=np.float64)})
    return paths
