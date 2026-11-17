"""Задания 14–33. Графики, статистики, корреляции, спектральный анализ, интерполяция, маски."""
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import stats
from scipy.interpolate import UnivariateSpline, interp1d
from scipy.signal import convolve, correlate, correlation_lags, periodogram, spectrogram

XLAB = "Номер измерения n"


def plot_series(data: np.ndarray, names, gaps=None) -> plt.Figure:
    """Задание 14. Графики исходных рядов и диаграмма рассеяния y(x)."""
    n = np.arange(1, len(data) + 1)
    fig, axes = plt.subplots(1, 2, figsize=(13, 4.6), gridspec_kw={"width_ratios": [1.6, 1]})
    for i, name in enumerate(names):
        axes[0].plot(n, data[:, i], lw=1.1, color=f"C{i}", label=name)
    if gaps is not None and len(gaps):
        axes[0].scatter(np.asarray(gaps) + 1, data[gaps, 0], s=60, facecolors="none",
                        edgecolors="red", zorder=5, label="заполненные пропуски x")
    axes[0].set_xlabel(XLAB)
    axes[0].set_ylabel("Значение")
    axes[0].set_title("Ряды x и y")
    axes[0].legend()
    axes[1].scatter(data[:, 0], data[:, 1], s=10, color="C2")
    axes[1].plot([data[:, 0].min(), data[:, 0].max()], [data[:, 0].min(), data[:, 0].max()], "k--", lw=1,
                 label="y = x")
    axes[1].set_xlabel("x")
    axes[1].set_ylabel("y")
    axes[1].set_title("Диаграмма рассеяния y(x)")
    axes[1].legend()
    fig.tight_layout()
    return fig


def plot_histograms(data: np.ndarray, names, bins: int = 15) -> plt.Figure:
    """Задание 15. Гистограмма частот и нормализованная гистограмма (оценка плотности)."""
    fig, axes = plt.subplots(2, len(names), figsize=(12, 7.5))
    for i, name in enumerate(names):
        x = data[:, i]
        edges = np.linspace(x.min(), x.max(), bins + 1)  # явные границы — столбцы на своих интервалах
        axes[0, i].hist(x, bins=edges, edgecolor="black", color=f"C{i}", alpha=0.7)
        axes[0, i].set_title(f"{name}: гистограмма частот")
        axes[0, i].set_ylabel("Число значений")
        axes[1, i].hist(x, bins=edges, density=True, edgecolor="black", color=f"C{i}", alpha=0.5,
                        label="нормализованная гистограмма")
        grid = np.linspace(x.min(), x.max(), 300)
        axes[1, i].plot(grid, stats.gaussian_kde(x)(grid), color="black", lw=1.5, label="ядерная оценка")
        axes[1, i].set_title(f"{name}: оценка плотности распределения")
        axes[1, i].set_ylabel("Плотность")
        axes[1, i].legend(fontsize=8, loc="upper left")
        for ax in axes[:, i]:
            ax.set_xlabel("Значение")
    fig.tight_layout()
    return fig


def sorted_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Задание 16. Каждый столбец, отсортированный по возрастанию (выводится на экран)."""
    out = pd.DataFrame({c: np.sort(df[c].to_numpy()) for c in df.columns})
    with pd.option_context("display.max_rows", None):
        print(out)
    return out


def plot_ecdf(data: np.ndarray, names) -> plt.Figure:
    """Задание 17. Эмпирические функции распределения (scipy.stats.ecdf)."""
    fig, axes = plt.subplots(1, len(names), figsize=(12, 4.2))
    for i, ax in enumerate(np.atleast_1d(axes)):
        res = stats.ecdf(data[:, i])
        res.cdf.plot(ax, color=f"C{i}")
        ax.set_title(f"Эмпирическая функция распределения {names[i]}")
        ax.set_xlabel(names[i])
        ax.set_ylabel("F(x)")
    fig.tight_layout()
    return fig


def column_stats(x: np.ndarray, bins: int = 15) -> dict:
    """Задание 18. Среднее, выборочная дисперсия, мода, медиана (NumPy) и мода интервального ряда."""
    mode = stats.mode(x, keepdims=False)
    hist, edges = np.histogram(x, bins=bins)
    k = int(np.argmax(hist))
    return {
        "mean": float(np.mean(x)),
        "var": float(np.var(x, ddof=1)),
        "mode": float(mode.mode),
        "mode_count": int(mode.count),
        "mode_interval": [float(edges[k]), float(edges[k + 1])],
        "mode_interval_mid": float((edges[k] + edges[k + 1]) / 2),
        "median": float(np.median(x)),
    }


def pandas_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Задание 18 средствами pandas: mean, var (ddof = 1), mode, median."""
    return pd.DataFrame({"mean": df.mean(), "var": df.var(), "mode": df.mode().iloc[0],
                         "median": df.median()})


def plot_stats(data: np.ndarray, names, st: list, bins: int = 15) -> plt.Figure:
    """Задание 18. Среднее, медиана, мода и ±СКО на гистограммах столбцов."""
    fig, axes = plt.subplots(1, len(names), figsize=(13, 4.4))
    for i, ax in enumerate(np.atleast_1d(axes)):
        x, s = data[:, i], st[i]
        edges = np.linspace(x.min(), x.max(), bins + 1)
        ax.hist(x, bins=edges, color=f"C{i}", alpha=0.35, edgecolor="black")
        sd = s["var"] ** 0.5
        ax.axvspan(s["mean"] - sd, s["mean"] + sd, color="grey", alpha=0.15, label="среднее ± СКО")
        ax.axvline(s["mean"], color="C3", lw=2, label=f"среднее {s['mean']:.3f}")
        ax.axvline(s["median"], color="C2", lw=2, ls="--", label=f"медиана {s['median']:.3f}")
        ax.axvline(s["mode_interval_mid"], color="C4", lw=2, ls=":",
                   label=f"мода (модальный интервал) {s['mode_interval_mid']:.3f}")
        ax.set_title(f"{names[i]}: дисперсия {s['var']:.3f}")
        ax.set_xlabel("Значение")
        ax.set_ylabel("Число значений")
        ax.legend(fontsize=8, loc="upper left")
    fig.tight_layout()
    return fig


def ci_mean(x: np.ndarray, alpha: float = 0.05):
    """Задание 19. Доверительный интервал для среднего (t-распределение)."""
    n = len(x)
    lo, hi = stats.t.interval(1 - alpha, df=n - 1, loc=np.mean(x), scale=stats.sem(x))
    return float(lo), float(hi)


def ci_var(x: np.ndarray, alpha: float = 0.05):
    """Задание 19. Доверительный интервал для дисперсии (распределение хи-квадрат)."""
    n = len(x)
    s2 = np.var(x, ddof=1)
    return (float((n - 1) * s2 / stats.chi2.ppf(1 - alpha / 2, n - 1)),
            float((n - 1) * s2 / stats.chi2.ppf(alpha / 2, n - 1)))


def plot_corr(df: pd.DataFrame, cov: pd.DataFrame, corr: pd.DataFrame) -> plt.Figure:
    """Задание 20. Матрицы ковариации и корреляции (тепловые карты)."""
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.4))
    for ax, m, title, cmap in [(axes[0], cov, "Ковариационная матрица", "viridis"),
                               (axes[1], corr, "Корреляционная матрица", "coolwarm")]:
        kw = {"vmin": -1, "vmax": 1} if m is corr else {}
        im = ax.imshow(m.to_numpy(), cmap=cmap, **kw)
        ax.set_xticks(range(len(m)), m.columns)
        ax.set_yticks(range(len(m)), m.index)
        ax.grid(False)
        for i in range(len(m)):
            for j in range(len(m)):
                ax.text(j, i, f"{m.iat[i, j]:.4f}", ha="center", va="center", fontsize=11,
                        color="white" if m is cov else "black")
        fig.colorbar(im, ax=ax)
        ax.set_title(title)
    fig.tight_layout()
    return fig


def corr_significance(df: pd.DataFrame, alpha: float = 0.05) -> pd.DataFrame:
    """Задание 21. Значимость коэффициентов корреляции Пирсона и Спирмена (scipy.stats)."""
    rows = []
    cols = list(df.columns)
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            res = stats.pearsonr(df[cols[i]], df[cols[j]])
            ci = res.confidence_interval(0.95)
            rho, p_s = stats.spearmanr(df[cols[i]], df[cols[j]])
            n = len(df)
            t = res.statistic * np.sqrt((n - 2) / (1 - res.statistic ** 2))
            rows.append({"Пара": f"{cols[i]}–{cols[j]}", "r": res.statistic, "p": res.pvalue,
                         "t": float(t), "t_crit": float(stats.t.ppf(1 - alpha / 2, n - 2)),
                         "ci_low": ci.low, "ci_high": ci.high, "spearman": rho, "p_spearman": p_s,
                         "Значима": "да" if res.pvalue < alpha else "нет"})
    return pd.DataFrame(rows)


def cross_correlation(x, y, max_lags: int = 40):
    """Задание 22. Нормированная взаимная корреляционная функция (scipy.signal.correlate)."""
    xc, yc = x - x.mean(), y - y.mean()
    c = correlate(xc, yc, mode="full") / (np.std(x) * np.std(y) * len(x))
    lags = correlation_lags(len(x), len(y), mode="full")
    m = np.abs(lags) <= max_lags
    return lags[m], c[m]


def plot_xcorr(results: list) -> plt.Figure:
    """Задание 22. Графики взаимной корреляции для нескольких пар."""
    fig, axes = plt.subplots(1, len(results), figsize=(6.2 * len(results), 4))
    for ax, (name, lags, c) in zip(np.atleast_1d(axes), results):
        ax.stem(lags, c, basefmt=" ")
        k = int(np.argmax(np.abs(c)))
        ax.set_title(f"{name}: максимум {c[k]:.3f} при сдвиге {lags[k]}")
        ax.set_xlabel("Сдвиг, число элементов")
        ax.set_ylabel("R(сдвиг)")
    fig.tight_layout()
    return fig


def plot_gradient(data, grads, names, theory) -> plt.Figure:
    """Задание 23. Численная производная (np.diff) и градиент (np.gradient) каждого столбца."""
    n = np.arange(1, len(data) + 1)
    fig, axes = plt.subplots(1, len(names), figsize=(13, 4.4))
    for i, ax in enumerate(np.atleast_1d(axes)):
        ax.plot(n[1:], np.diff(data[:, i]), ".", ms=4, color="grey", label="np.diff (разность вперёд)")
        ax.plot(n, grads[:, i], lw=1.2, color=f"C{i}", label="np.gradient")
        ax.plot(n, theory, "k--", lw=1.5, label="1 / (2√n)")
        ax.set_title(f"Производная {names[i]} по номеру измерения")
        ax.set_xlabel(XLAB)
        ax.set_ylabel(f"d{names[i]}/dn")
        ax.legend(fontsize=8)
    axes[0].set_ylim(0, 0.55)
    fig.tight_layout()
    return fig


def convolution(x, kernel):
    """Задание 24. Свёртка средствами NumPy и SciPy."""
    return np.convolve(x, kernel, mode="full"), convolve(x, kernel, mode="full")


def plot_convolution(x, y, cxy, ysm, width) -> plt.Figure:
    """Задание 24. Свёртка столбцов x∗y и сглаживание y прямоугольным ядром."""
    fig, axes = plt.subplots(1, 2, figsize=(13, 4.3))
    axes[0].plot(np.arange(len(cxy)), cxy, color="C4")
    axes[0].set_title(f"Свёртка столбцов x ∗ y (длина {len(cxy)})")
    axes[0].set_xlabel("Индекс k")
    axes[0].set_ylabel("(x ∗ y)[k]")
    n = np.arange(1, len(y) + 1)
    axes[1].plot(n, y, lw=0.8, alpha=0.6, color="C1", label="y (исходный)")
    h = (width - 1) // 2
    axes[1].plot(n[h:len(y) - h], ysm, lw=1.8, color="C3", label=f"y ∗ ядро ширины {width}")
    axes[1].plot(n, np.sqrt(n), "k--", lw=1, label="√n")
    axes[1].set_xlabel(XLAB)
    axes[1].set_ylabel("y")
    axes[1].set_title("Сглаживание y свёрткой с прямоугольным ядром")
    axes[1].legend()
    fig.tight_layout()
    return fig


def distribution_tests(df: pd.DataFrame) -> pd.DataFrame:
    """
    Задание 27. Гипотеза о равномерном распределении — для первого столбца, о нормальном —
    для второго (дополнительно проверены обе гипотезы для обоих столбцов).
    """
    rows = []
    for col in df.columns:
        x = df[col].to_numpy()
        _, p_unif = stats.kstest(x, stats.uniform(loc=x.min(), scale=x.max() - x.min()).cdf)
        _, p_norm = stats.kstest(x, stats.norm(loc=x.mean(), scale=x.std(ddof=1)).cdf)
        _, p_sh = stats.shapiro(x)
        _, p_dag = stats.normaltest(x)
        rows.append({"Столбец": col, "p (КС, равномерное)": p_unif, "p (КС, нормальное)": p_norm,
                     "p (Шапиро–Уилк)": p_sh, "p (Д'Агостино)": p_dag})
    return pd.DataFrame(rows)


def plot_spectral(series: dict, nperseg: int = 32, noverlap: int = 24) -> plt.Figure:
    """Задания 28–30. Спектрограмма, периодограмма и АЧХ (БПФ) для каждого ряда."""
    k = len(series)
    fig, axes = plt.subplots(3, k, figsize=(5.2 * k, 11))
    for j, (name, x) in enumerate(series.items()):
        xc = x - x.mean()
        f, t, sxx = spectrogram(xc, fs=1.0, nperseg=nperseg, noverlap=noverlap)
        pcm = axes[0, j].pcolormesh(t, f, 10 * np.log10(sxx + 1e-12), shading="gouraud", cmap="viridis")
        fig.colorbar(pcm, ax=axes[0, j], label="Мощность, дБ")
        axes[0, j].set_xlabel(XLAB)
        axes[0, j].set_ylabel("Частота, 1/отсчёт")
        axes[0, j].set_title(f"Спектрограмма {name}")

        f, pxx = periodogram(xc, fs=1.0)
        axes[1, j].semilogy(f[1:], pxx[1:], lw=0.9, color=f"C{j}")
        axes[1, j].axhline(pxx[1:].mean(), color="C3", ls="--", label="средний уровень")
        axes[1, j].legend(loc="upper right", fontsize=8)
        axes[1, j].set_xlabel("Частота, 1/отсчёт")
        axes[1, j].set_ylabel("Спектральная плотность")
        axes[1, j].set_title(f"Периодограмма {name}")

        amp = np.abs(np.fft.rfft(xc)) / len(xc)
        freq = np.fft.rfftfreq(len(xc), d=1.0)
        axes[2, j].plot(freq[1:], amp[1:], lw=0.9, color=f"C{j}")
        axes[2, j].set_xlabel("Частота, 1/отсчёт")
        axes[2, j].set_ylabel("|X(f)| / N")
        axes[2, j].set_title(f"АЧХ {name} (БПФ)")
    fig.tight_layout()
    return fig


def interpolate(x, factor: int = 10, s=0):
    """Задания 31–32. Кубическая интерполяция (interp1d) и интерполяция сплайном (UnivariateSpline)."""
    n = np.arange(1, len(x) + 1, dtype=float)
    n_new = np.linspace(1, len(x), factor * (len(x) - 1) + 1)
    y_cubic = interp1d(n, x, kind="cubic")(n_new)
    y_spl = UnivariateSpline(n, x, k=3, s=s)(n_new)
    return n, n_new, y_cubic, y_spl


def plot_interpolation(parts) -> plt.Figure:
    """parts: список (имя, n, x, n_new, y_cubic, y_spl, y_smooth)."""
    fig, axes = plt.subplots(1, len(parts), figsize=(13, 4.5))
    for ax, (name, n, x, n_new, y_cub, y_spl, y_sm) in zip(np.atleast_1d(axes), parts):
        ax.plot(n_new, y_cub, lw=2.4, alpha=0.6, label="кубическая (interp1d)")
        ax.plot(n_new, y_spl, "--", lw=1.2, color="C3", label="сплайн k = 3, s = 0")
        if y_sm is not None:
            ax.plot(n_new, y_sm, lw=1.6, color="C2", label="сглаживающий сплайн")
        ax.plot(n, x, "ko", ms=4, label="исходные точки")
        ax.set_xlabel(XLAB)
        ax.set_ylabel(name)
        ax.set_title(f"Интерполяция фрагмента {name}")
        ax.legend(fontsize=8, loc="lower right")
    fig.tight_layout()
    return fig


def binary_masks(x):
    """Задание 33. Бинарные маски."""
    return {"x > 0": x > 0, "x < 0": x < 0, "x = 0": x == 0, "-1 <= x <= 1": (x >= -1) & (x <= 1)}


def plot_masks(z: np.ndarray, masks: dict, names) -> plt.Figure:
    """Маски стандартизованных столбцов: ряд и четыре маски в виде полос 0/1."""
    fig, axes = plt.subplots(2, len(names), figsize=(13, 6.5), gridspec_kw={"height_ratios": [1.6, 1]})
    n = np.arange(1, z.shape[0] + 1)
    for i, name in enumerate(names):
        ax = axes[0, i]
        m = masks[name]
        ax.plot(n, z[:, i], color="grey", lw=0.7)
        ax.scatter(n[m["x > 0"]], z[m["x > 0"], i], s=8, color="C2", label="z > 0")
        ax.scatter(n[m["x < 0"]], z[m["x < 0"], i], s=8, color="C3", label="z < 0")
        mm = m["-1 <= x <= 1"]
        ax.scatter(n[mm], z[mm, i], s=26, facecolors="none", edgecolors="C0", label="−1 ≤ z ≤ 1")
        ax.axhspan(-1, 1, color="C0", alpha=0.1)
        ax.set_title(f"Стандартизованный {name} (z) и маски")
        ax.set_xlabel(XLAB)
        ax.set_ylabel("z")
        ax.legend(fontsize=8, loc="upper left")
        img = np.vstack([m[k] for k in ("x > 0", "x < 0", "x = 0", "-1 <= x <= 1")]).astype(int)
        axes[1, i].imshow(img, aspect="auto", cmap="Greys", interpolation="nearest",
                          extent=(0.5, len(n) + 0.5, 3.5, -0.5))
        axes[1, i].set_yticks(range(4), ["z > 0", "z < 0", "z = 0", "−1 ≤ z ≤ 1"])
        axes[1, i].set_xlabel(XLAB)
        axes[1, i].set_title("Бинарные маски (чёрный — 1)")
        axes[1, i].grid(False)
    fig.tight_layout()
    return fig
