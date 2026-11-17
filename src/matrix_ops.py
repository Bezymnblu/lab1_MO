"""Задания 46–49. Разреженная матрица, обратная матрица, массив F, пайплайн с PCA."""
import matplotlib.pyplot as plt
import numpy as np
from scipy import sparse
from sklearn.decomposition import PCA
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import FunctionTransformer, MinMaxScaler, StandardScaler

from src.config import SEED
from src.preprocess import MinMaxScalerCustom, StandardScalerCustom

LAYERS = ["исходные", "MinMax [0, 1]", "MinMax [−1, 1]", "z-оценка",
          "маска z > 0", "маска z < 0", "маска z = 0", "маска −1 ≤ z ≤ 1"]
# наборы признаков — непрерывные срезы по оси слоёв, поэтому каждый из них — представление F
FEATURE_SETS = {"исходные": slice(0, 1), "масштабированные": slice(1, 4), "маски": slice(4, 8),
                "исходные + масштабированные": slice(0, 4), "все слои": slice(0, 8)}


def sparse_demo() -> dict:
    """Задание 46. Случайная разреженная матрица 100 000 × 100 000 (scipy.sparse.random)."""
    # Generator выбирает позиции ненулевых элементов без перебора всех 10^10 индексов
    S = sparse.random(100_000, 100_000, density=1e-5, format="csr", random_state=np.random.default_rng(SEED))
    mem = S.data.nbytes + S.indices.nbytes + S.indptr.nbytes
    dense = 100_000 * 100_000 * 8
    print(f"Задание 46. Форма {S.shape}, ненулевых элементов: {S.nnz}, "
          f"память CSR: {mem / 2**20:.2f} МиБ (плотная заняла бы {dense / 2**30:.1f} ГиБ)")
    return {"shape": list(S.shape), "nnz": int(S.nnz), "mem_mib": mem / 2**20, "dense_gib": dense / 2**30,
            "max_row_nnz": int(np.diff(S.indptr).max()), "density": S.nnz / 1e10}


def inverse_demo() -> dict:
    """Задание 47. Обратная матрица 100 × 100 (если существует) и проверка M·M⁻¹ = E."""
    M = np.random.rand(100, 100)
    rank = int(np.linalg.matrix_rank(M))
    res = {"rank": rank, "det": float(np.linalg.det(M)), "cond": float(np.linalg.cond(M))}
    if rank == 100:
        M_inv = np.linalg.inv(M)
        res["max_err"] = float(np.max(np.abs(M @ M_inv - np.eye(100))))
        res["max_err_left"] = float(np.max(np.abs(M_inv @ M - np.eye(100))))
        print(f"Задание 47. Ранг {rank}, det(M) = {res['det']:.4e}, число обусловленности {res['cond']:.1f}, "
              f"max |M·M⁻¹ − E| = {res['max_err']:.2e}")
    else:
        print("Задание 47. Матрица вырождена, обратной не существует")
    return res


def build_F(X: np.ndarray, mask_source: np.ndarray) -> tuple:
    """
    Задание 48. Многомерный массив F формы (n, m, 8): исходные данные, три варианта
    масштабирования и четыре бинарные маски (задание 33). Массив создаётся один раз,
    слои и наборы признаков извлекаются срезами — представлениями без копирования.
    """
    n, m = X.shape
    F = np.empty((n, m, len(LAYERS)))
    F[..., 0] = X
    F[..., 1] = MinMaxScalerCustom(0, 1).fit_transform(X)
    F[..., 2] = MinMaxScalerCustom(-1, 1).fit_transform(X)
    F[..., 3] = StandardScalerCustom().fit_transform(X)
    z = mask_source
    for k, m_ in enumerate([z > 0, z < 0, z == 0, (z >= -1) & (z <= 1)]):
        F[..., 4 + k] = m_
    views = {name: F[:, :, sl] for name, sl in FEATURE_SETS.items()}
    info = {"shape": list(F.shape), "nbytes": int(F.nbytes),
            "views": {k: {"shape": list(v.shape), "shares_memory": bool(np.shares_memory(F, v)),
                          "base_is_F": v.base is F} for k, v in views.items()},
            "flat_is_view": bool(np.shares_memory(F, F.reshape(n, -1)))}
    print("Задание 48. Форма F:", F.shape, "| все наборы — представления:",
          all(v["shares_memory"] for v in info["views"].values()))
    return F, views, info


def _flat(view: np.ndarray) -> np.ndarray:
    flat = view.reshape(len(view), -1)
    return flat[:, flat.std(axis=0) > 0]  # постоянные признаки (например, маска z = 0) исключаются


def pca_pipeline(F: np.ndarray, views: dict) -> dict:
    """
    Задание 49. Пайплайн «масштабирование → PCA» для каждого набора признаков F и каждого
    варианта масштабирования. Наилучшая проекция — наибольшая доля дисперсии, сохраняемая
    проекцией на 2 главные компоненты (для наборов из 3 и более признаков); для каждой
    комбинации вычисляется логарифм правдоподобия модели PCA от числа компонент.
    """
    scalers = {"без масштабирования": FunctionTransformer, "MinMaxScaler": MinMaxScaler,
               "StandardScaler": StandardScaler}
    rows, curves = [], {}
    for set_name, view in views.items():
        flat = _flat(view)
        rank_c = int(np.linalg.matrix_rank(flat - flat.mean(axis=0)))
        for sc_name, Sc in scalers.items():
            pipe = Pipeline([("scaler", Sc()), ("pca", PCA(random_state=SEED))]).fit(flat)
            evr = pipe.named_steps["pca"].explained_variance_ratio_
            z = pipe.named_steps["scaler"].transform(flat)
            ll = []
            for k in range(1, rank_c):  # при k = rank дисперсия шума нулевая — правдоподобие не определено
                p = Pipeline([("scaler", Sc()), ("pca", PCA(n_components=k, random_state=SEED))]).fit(flat)
                ll.append(float(p.score(flat)))
            curves[(set_name, sc_name)] = ll
            rows.append({"set": set_name, "scaler": sc_name, "n_features": int(flat.shape[1]),
                         "rank_centered": rank_c, "evr1": float(evr[0]),
                         "evr2": float(evr[:2].sum()) if len(evr) > 1 else 1.0,
                         "ll1": ll[0] if ll else None, "total_var": float(z.var(axis=0).sum())})
    # сравниваются пайплайны с масштабированием; наборы с рангом 2 проецируются на плоскость
    # без потерь (100 %) тривиально, поэтому в выборе участвуют наборы с рангом не менее 3
    cands = [r for r in rows if r["rank_centered"] >= 3 and r["scaler"] != "без масштабирования"]
    best = max(cands, key=lambda r: (round(r["evr2"], 12), r["evr1"]))
    print("Задание 49. Наилучшая проекция:", best["set"], "+", best["scaler"],
          f"(2 компоненты сохраняют {best['evr2']:.4%} дисперсии)")
    # проекция лучшей комбинации на 2 главные компоненты
    Sc = scalers[best["scaler"]]
    pipe = Pipeline([("scaler", Sc()), ("pca", PCA(n_components=2, random_state=SEED))])
    proj = pipe.fit_transform(_flat(views[best["set"]]))
    return {"rows": rows, "curves": {f"{a} | {b}": v for (a, b), v in curves.items()},
            "best": best, "proj": proj.tolist()}


def plot_pca(results: dict) -> plt.Figure:
    """Задание 49. Логарифм правдоподобия PCA от числа компонент и проекция лучшей комбинации."""
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    styles = {"без масштабирования": ":", "MinMaxScaler": "-", "StandardScaler": "--"}
    sets = list(FEATURE_SETS)
    for key, ll in results["curves"].items():
        if len(ll) < 2:
            continue
        s, sc = key.split(" | ")
        axes[0].plot(range(1, len(ll) + 1), ll, styles[sc], marker="o", ms=3,
                     color=f"C{sets.index(s)}", label=f"{s}, {sc}")
    axes[0].set_yscale("symlog", linthresh=1)
    axes[0].set_xlabel("Число главных компонент")
    axes[0].set_ylabel("Средний логарифм правдоподобия")
    axes[0].set_title("Логарифм правдоподобия модели PCA")
    axes[0].xaxis.get_major_locator().set_params(integer=True)
    axes[0].legend(fontsize=7, ncol=2)
    b = results["best"]
    proj = np.array(results["proj"])
    sc_ = axes[1].scatter(proj[:, 0], proj[:, 1], c=np.arange(1, len(proj) + 1), cmap="viridis", s=12)
    fig.colorbar(sc_, ax=axes[1], label="Номер измерения n")
    axes[1].set_xlabel("ГК 1")
    axes[1].set_ylabel("ГК 2")
    axes[1].set_title(f"Проекция: {b['set']}, {b['scaler']} ({b['evr2']:.2%} дисперсии)")
    fig.tight_layout()
    return fig
