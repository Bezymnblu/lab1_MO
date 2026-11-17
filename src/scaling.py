"""Задания 34–35. Сравнение собственного масштабирования со scikit-learn, инверсия."""
import matplotlib.pyplot as plt
import numpy as np
from sklearn.preprocessing import MinMaxScaler, StandardScaler

from src.preprocess import MinMaxScalerCustom, StandardScalerCustom


def compare_scalers(x: np.ndarray) -> dict:
    """Масштабирует x собственными классами и классами sklearn, сравнивает результат и инверсию."""
    pairs = {
        "minmax": (MinMaxScalerCustom().fit(x), MinMaxScaler().fit(x)),
        "standard": (StandardScalerCustom().fit(x), StandardScaler().fit(x)),
    }
    res = {}
    for name, (custom, sk) in pairs.items():
        xc, xs = custom.transform(x), sk.transform(x)
        bc, bs = custom.inverse_transform(xc), sk.inverse_transform(xs)
        assert np.allclose(xc, xs, atol=1e-10)
        assert np.allclose(x, bc, atol=1e-10) and np.allclose(x, bs, atol=1e-10)
        res[name] = {
            "custom": xc, "sklearn": xs, "inv_custom": bc, "inv_sklearn": bs,
            "max_diff": float(np.max(np.abs(xc - xs))),
            "inv_err_custom": float(np.max(np.abs(x - bc))),
            "inv_err_sklearn": float(np.max(np.abs(x - bs))),
            "inv_diff": float(np.max(np.abs(bc - bs))),
        }
    return res


def plot_scaling(X: np.ndarray, results: dict, names) -> plt.Figure:
    """Задание 34. Наложение результатов собственного и sklearn-масштабирования для каждого столбца."""
    fig, axes = plt.subplots(len(names), 3, figsize=(15, 4 * len(names)))
    n = np.arange(1, len(X) + 1)
    for i, name in enumerate(names):
        res = results[name]
        axes[i, 0].plot(n, X[:, i], lw=1, color=f"C{i}")
        axes[i, 0].set_title(f"{name}: исходный ряд")
        for ax, key, title in [(axes[i, 1], "minmax", "MinMax в [0, 1]"),
                               (axes[i, 2], "standard", "Стандартизация")]:
            ax.plot(n, res[key]["custom"], lw=2.6, alpha=0.55, color=f"C{i}", label="собственный класс")
            ax.plot(n, res[key]["sklearn"], "--", lw=1, color="black", label="scikit-learn")
            ax.set_title(f"{name}: {title}; max |разность| = {res[key]['max_diff']:.1e}")
            ax.legend(loc="lower right", fontsize=8)
        for ax in axes[i]:
            ax.set_xlabel("Номер измерения n")
    fig.tight_layout()
    return fig
