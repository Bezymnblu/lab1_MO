"""Задания 36–37. Скользящее окно и скользящее среднее."""
import matplotlib.pyplot as plt
import numpy as np


def sliding_window(x: np.ndarray, width: int) -> np.ndarray:
    """Задание 36. Матрица окон формы (n - width + 1, width)."""
    if width <= 0 or width > len(x):
        raise ValueError("Некорректная ширина окна")
    n = len(x)
    idx = np.arange(width)[None, :] + np.arange(n - width + 1)[:, None]
    return x[idx]


def moving_average(x: np.ndarray, width: int) -> np.ndarray:
    """Задание 37. Скользящее среднее через матрицу окон."""
    return sliding_window(x, width).mean(axis=1)


def plot_moving_average(x, widths, name) -> plt.Figure:
    fig, ax = plt.subplots(figsize=(12, 4.5))
    ax.plot(np.arange(1, len(x) + 1), x, lw=0.7, color="grey", label=f"{name} (исходный)")
    for i, w in enumerate(widths):
        ma = moving_average(x, w)
        # значение окна относим к его центру
        ax.plot(np.arange(len(ma)) + 1 + (w - 1) / 2, ma, lw=1.8, color=f"C{i + 1}",
                label=f"скользящее среднее, окно {w}")
    ax.set_xlabel("Номер измерения n")
    ax.set_ylabel(name)
    ax.legend()
    fig.tight_layout()
    return fig
