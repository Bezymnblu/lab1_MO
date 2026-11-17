"""Задания 8–13. Обработка проблем в данных, перевод в NumPy, масштабирование, разбиение."""
import json

import numpy as np
import pandas as pd

from src.config import timestamped_filename
from src.data_loader import check_data_quality

# действия меню задания 8, сгруппированные по типу проблемы
MENU = {
    "bad": [("drop_bad", "выбросить строки, содержащие некорректные данные"),
            ("cast", "выполнить приведение типов (некорректные значения станут пустыми)")],
    "missing": [("drop", "выбросить строки с пустыми ячейками"),
                ("mean", "заполнить пустые ячейки средним значением по столбцу"),
                ("ffill", "заполнить пустые ячейки предыдущим значением"),
                ("bfill", "заполнить пустые ячейки последующим значением"),
                ("zero", "заполнить пустые ячейки нулями"),
                ("value", "заполнить ячейку указанным значением")],
}


def handle_missing(df: pd.DataFrame, col: str, action: str, value=None, row=None) -> pd.DataFrame:
    """
    Задание 8. Обработка проблемного столбца средствами pandas.
      'drop_bad' — удалить строки с нечисловыми значениями в col
      'cast'     — привести к числу (нечисловые значения станут NaN)
      'drop'     — удалить строки с NaN в col
      'mean'     — заполнить средним по столбцу
      'ffill'    — заполнить предыдущим значением
      'bfill'    — заполнить последующим значением
      'zero'     — заполнить нулём
      'value'    — заполнить указанным значением (все пропуски или одну ячейку row)
    """
    df = df.copy()
    if action == "drop_bad":
        bad = pd.to_numeric(df[col], errors="coerce").isna() & df[col].notna()
        df = df[~bad]
    elif action == "cast":
        df[col] = pd.to_numeric(df[col], errors="coerce")
    elif action == "drop":
        df = df.dropna(subset=[col])
    elif action == "mean":
        df[col] = df[col].fillna(df[col].mean())
    elif action == "ffill":
        df[col] = df[col].ffill()
    elif action == "bfill":
        df[col] = df[col].bfill()
    elif action == "zero":
        df[col] = df[col].fillna(0)
    elif action == "value":
        if row is None:
            df[col] = df[col].fillna(float(value))
        else:
            df.loc[row, col] = float(value)
    else:
        raise ValueError(f"Неизвестное действие: {action}")
    return df


def interactive_menu(df: pd.DataFrame, numeric_cols=None) -> pd.DataFrame:
    """
    Задание 8. Интерактивное меню (python main.py --interactive): для каждого проблемного
    столбца предлагаются действия, соответствующие типу найденной проблемы.
    """
    numeric_cols = numeric_cols or list(df.columns)
    while True:
        rep = check_data_quality(df, numeric_cols, warn=False)
        problems = []
        for c in numeric_cols:
            if c in rep["non_numeric"] or c in rep["type_mismatch"]:
                problems.append((c, "bad"))
            if rep["missing"][c]["count"]:
                problems.append((c, "missing"))
        if not problems:
            print("Проблем в данных не осталось.")
            return df
        print("\n=== Меню обработки данных (задание 8) ===")
        for i, (c, kind) in enumerate(problems, 1):
            what = ("некорректные данные / несовпадение типов" if kind == "bad"
                    else f"пустые ячейки в строках {rep['missing'][c]['rows']}")
            print(f"{i}. Столбец {c}: {what}")
        print("0. Выход")
        ch = input("Номер проблемы: ").strip()
        if ch == "0":
            return df
        if not ch.isdigit() or not 1 <= int(ch) <= len(problems):
            print("Нет такого пункта.")
            continue
        col, kind = problems[int(ch) - 1]
        for i, (_, text) in enumerate(MENU[kind], 1):
            print(f"  {i}. {text}")
        a = input("  Действие: ").strip()
        if not a.isdigit() or not 1 <= int(a) <= len(MENU[kind]):
            print("Нет такого действия.")
            continue
        action = MENU[kind][int(a) - 1][0]
        value = row = None
        if action == "value":
            value = input("  Значение: ").strip()
            r = input("  Номер строки таблицы (пусто — все пропуски): ").strip()
            row = int(r) if r else None
        df = handle_missing(df, col, action, value, row)
        print(f"  Готово. Строк: {len(df)}, пропусков в {col}: {int(df[col].isna().sum())}")


def to_numpy(df: pd.DataFrame) -> np.ndarray:
    """Задание 9. DataFrame -> массив NumPy float64."""
    return df.to_numpy(dtype=np.float64)


class MinMaxScalerCustom:
    """
    Задания 10–11. Мин-макс-нормализация в [0, 1] (a = 0, b = 1) и в произвольных
    границах [a, b]; обратное преобразование по сохранённым min и max.
    """

    def __init__(self, a: float = 0.0, b: float = 1.0):
        self.a, self.b = a, b
        self.min_, self.max_ = None, None

    def fit(self, x: np.ndarray):
        self.min_ = np.min(x, axis=0)
        self.max_ = np.max(x, axis=0)
        return self

    def _denom(self):
        rng = self.max_ - self.min_
        return np.where(rng == 0, 1, rng)  # защита от деления на ноль

    def transform(self, x: np.ndarray) -> np.ndarray:
        return self.a + (x - self.min_) * (self.b - self.a) / self._denom()

    def fit_transform(self, x):
        return self.fit(x).transform(x)

    def inverse_transform(self, x_new: np.ndarray) -> np.ndarray:
        return (x_new - self.a) * self._denom() / (self.b - self.a) + self.min_

    def params(self) -> dict:
        return {"method": "minmax", "a": self.a, "b": self.b,
                "min": np.atleast_1d(self.min_).tolist(), "max": np.atleast_1d(self.max_).tolist()}


class StandardScalerCustom:
    """Задания 10–11. Стандартизация z = (x − μ) / σ (σ с делителем n − ddof) и обратное."""

    def __init__(self, ddof: int = 0):
        self.ddof = ddof
        self.mean_, self.std_ = None, None

    def fit(self, x: np.ndarray):
        self.mean_ = np.mean(x, axis=0)
        self.std_ = np.std(x, axis=0, ddof=self.ddof)
        return self

    def _std(self):
        return np.where(self.std_ == 0, 1, self.std_)

    def transform(self, x):
        return (x - self.mean_) / self._std()

    def fit_transform(self, x):
        return self.fit(x).transform(x)

    def inverse_transform(self, x_new):
        return x_new * self._std() + self.mean_

    def params(self) -> dict:
        return {"method": "z-score", "ddof": self.ddof, "mean": np.atleast_1d(self.mean_).tolist(),
                "std": np.atleast_1d(self.std_).tolist()}


def save_scaler_params(scalers: dict, base: str = "task10_scaling_params") -> str:
    """Задание 10. Значения, используемые при масштабировании, сохраняются в JSON."""
    path = timestamped_filename(base, "json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({k: s.params() for k, s in scalers.items()}, f, ensure_ascii=False, indent=1)
    return path


def load_scaler(params: dict):
    """Восстанавливает масштабировщик из сохранённых параметров (для задания 11)."""
    if params["method"] == "minmax":
        s = MinMaxScalerCustom(params["a"], params["b"])
        s.min_, s.max_ = np.array(params["min"]), np.array(params["max"])
    else:
        s = StandardScalerCustom(params["ddof"])
        s.mean_, s.std_ = np.array(params["mean"]), np.array(params["std"])
    return s


def split_train_val_test(x: np.ndarray, ratios, shuffle: bool = False, name: str = "data") -> dict:
    """
    Задание 12. Разбиение на 3 массива в заданном соотношении: в процентах (75, 10, 15),
    в долях (0,75; 0,1; 0,15) или в частях (1, 1, 1) — соотношение нормируется на сумму,
    последняя часть получает остаток. Массивы получают метки _train, _val, _test.
    """
    r = np.array(ratios, dtype=float)
    r = r / r.sum()
    n = len(x)
    n_train = int(round(n * r[0]))
    n_val = int(round(n * r[1]))
    idx = np.random.permutation(n) if shuffle else np.arange(n)
    x = x[idx]
    return {f"{name}_train": x[:n_train], f"{name}_val": x[n_train:n_train + n_val],
            f"{name}_test": x[n_train + n_val:]}
