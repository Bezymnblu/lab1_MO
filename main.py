"""
Лабораторная работа №1. Библиотеки Python для машинного обучения. Работа с данными.
Вариант 7 (Lab1_Dataset.xlsx, столбцы x и y). Точка входа: выполняет задания 1–50.

Запуск:
    python main.py                 — неинтерактивно (сценарий обработки из задания 8)
    python main.py --interactive   — с интерактивным меню задания 8
"""
import argparse
import glob
import hashlib
import json
import os
import sys
import warnings

os.environ.setdefault("TF_CPP_MIN_LOG_LEVEL", "2")
os.environ.setdefault("TF_ENABLE_ONEDNN_OPTS", "0")

import numpy as np
import pandas as pd

from src import config
from src.config import OUTPUT_DIR, VARIANT, VARIANT_COLUMNS, save_figure, set_seeds, timestamped_filename


class Tee:
    """Дублирует вывод в консоль и в файл журнала."""

    def __init__(self, path):
        self.file = open(path, "w", encoding="utf-8")
        self.stdout = sys.stdout

    def write(self, s):
        self.stdout.write(s)
        self.file.write(s)

    def flush(self):
        self.stdout.flush()
        self.file.flush()


def header(text):
    print("\n" + "=" * 78 + f"\n{text}\n" + "=" * 78)


def md5(path):
    return hashlib.md5(open(path, "rb").read()).hexdigest()


def main(interactive: bool = False):
    # Перед запуском удаляются результаты прошлых запусков, чтобы в output был один комплект
    for f in glob.glob(os.path.join(OUTPUT_DIR, "*")):
        os.remove(f)
    sys.stdout = Tee(timestamped_filename("run_log", "txt"))
    R = {"run_ts": config.RUN_TS}  # числовые результаты для отчёта

    from src import (data_loader as dl, image_proc, io_utils, matrix_ops, preprocess as pp,
                     scaling, stats_analysis as sa, tensors, windows)

    # ------------------------------------------------------------------ Задание 13
    header("Задание 13. Фиксация seed для всех генераторов случайных чисел")
    set_seeds(config.SEED)
    a = np.random.rand(3)
    set_seeds(config.SEED)
    b = np.random.rand(3)
    print(f"SEED = {config.SEED}; np.random.rand(3) после set_seeds дважды: {a} и {b}; "
          f"совпадают: {np.array_equal(a, b)}")
    R["seed_repeat"] = bool(np.array_equal(a, b))
    set_seeds(config.SEED)

    # ------------------------------------------------------------------ Задание 1
    header("Задание 1. Чтение файла датасета")
    hash_before = md5(config.DATASET_PATH)
    raw = io_utils.read_table(config.DATASET_PATH, header=False)   # лист целиком, без заголовков
    print(f"Прочитан {os.path.basename(config.DATASET_PATH)}: {raw.shape[0]} строк, {raw.shape[1]} столбцов")
    print("Автоматические имена столбцов:", list(raw.columns[:6]), "...", list(raw.columns[-2:]))
    var_cols = dl.variant_columns(raw)
    print(f"Найдено вариантов: {len(var_cols)}; вариант {VARIANT} — столбцы листа "
          f"{[c for c in raw.columns[var_cols[VARIANT]]]}")
    R["raw_shape"] = list(raw.shape)
    R["n_variants"] = len(var_cols)
    R["variant_cols"] = [int(c) + 1 for c in var_cols[VARIANT]]

    # ------------------------------------------------------------------ Задание 5
    header("Задание 5. Вывод содержимого таблицы")
    dl.show_table(raw)

    # ------------------------------------------------------------------ Задание 6
    header("Задание 6. Извлечение столбцов варианта")
    raw_copy = raw.copy()
    sub = dl.extract_variant(raw, VARIANT)
    print(f"Столбцы варианта {VARIANT}: {list(sub.columns)}, форма {sub.shape}")
    dl.show_table(sub)
    R["sub_shape"] = list(sub.shape)
    R["raw_unchanged"] = bool(raw.equals(raw_copy))
    R["file_unchanged"] = md5(config.DATASET_PATH) == hash_before
    print("DataFrame листа не изменён:", R["raw_unchanged"], "| файл не изменён (MD5):", R["file_unchanged"])

    # ------------------------------------------------------------------ Задание 4
    header("Задание 4. Проверка на пустые ячейки, некорректные данные, несовпадение типов")
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        quality = dl.check_data_quality(sub, VARIANT_COLUMNS)
    R["warnings"] = [str(w.message) for w in caught]
    for w in R["warnings"]:
        print("ПРЕДУПРЕЖДЕНИЕ:", w)
    print(json.dumps(quality, ensure_ascii=False, indent=1))
    with open(timestamped_filename("task4_quality", "json"), "w", encoding="utf-8") as f:
        json.dump(quality, f, ensure_ascii=False, indent=1)
    R["quality"] = quality
    sheet = dl.scan_sheet(raw)
    sheet_df = pd.DataFrame(sheet)
    print("Проверка всех вариантов листа:\n", sheet_df.to_string(index=False))
    sheet_df.to_excel(timestamped_filename("task4_sheet_scan", "xlsx"), index=False)
    R["sheet"] = sheet

    # ------------------------------------------------------------------ Задание 7
    header("Задание 7. Явные типы данных pandas")
    print("Типы до приведения:\n", sub.dtypes)
    typed = dl.cast_types(sub, {"x": "float64", "y": "float64"})
    print("Типы после приведения:\n", typed.dtypes)
    R["dtypes_before"] = sub.dtypes.astype(str).to_dict()
    R["dtypes_after"] = typed.dtypes.astype(str).to_dict()

    # ------------------------------------------------------------------ Задание 8
    header("Задание 8. Обработка проблем в данных")
    gaps = np.flatnonzero(typed["x"].isna().to_numpy())
    n_true = np.sqrt(gaps + 1.0)  # x = √n, n — номер измерения
    demo_rows = []
    for action, val in [("drop", None), ("mean", None), ("ffill", None), ("bfill", None),
                        ("zero", None), ("value", -1)]:
        d = pp.handle_missing(typed, "x", action, val)
        filled = d["x"].reindex(gaps).dropna()
        demo_rows.append({"action": action, "rows": len(d), "missing": int(d["x"].isna().sum()),
                          "values": filled.round(4).tolist(),
                          "err": float(np.max(np.abs(filled.to_numpy() - n_true))) if len(filled) else None})
    print(pd.DataFrame(demo_rows).to_string(index=False))
    R["handle_demo"] = demo_rows
    # действия для некорректных значений — на варианте, где они есть
    bad_var = next(s for s in sheet if s["non_numeric"])
    vb = dl.extract_variant(raw, bad_var["variant"])
    qb = dl.check_data_quality(vb, warn=False)
    colb = next(iter(qb["non_numeric"]))
    R["bad_demo"] = {"variant": bad_var["variant"], "col": colb, "values": qb["non_numeric"][colb]["values"],
                     "rows_excel": qb["non_numeric"][colb]["rows"], "n": len(vb),
                     "drop_bad_rows": len(pp.handle_missing(vb, colb, "drop_bad")),
                     "cast_dtype": str(pp.handle_missing(vb, colb, "cast")[colb].dtype),
                     "cast_missing": int(pp.handle_missing(vb, colb, "cast")[colb].isna().sum())}
    print("Действия для некорректных данных (вариант", bad_var["variant"], "):", R["bad_demo"])

    if interactive:
        clean = pp.interactive_menu(sub.copy(), VARIANT_COLUMNS)
        clean = dl.cast_types(clean, {"x": "float64", "y": "float64"}).reset_index(drop=True)
    else:
        # несовпадение типов (object) -> приведение типов; x — монотонный гладкий ряд √n,
        # поэтому пропуски заполняются предыдущим значением (ffill), а не средним
        clean = pp.handle_missing(sub, "x", "cast")
        clean = pp.handle_missing(clean, "y", "cast")
        clean = dl.cast_types(clean, {"x": "float64", "y": "float64"})
        R["x_autocorr"] = float(clean["x"].autocorr(1))
        R["x_mean_before"] = float(clean["x"].mean())
        R["gap_neighbors"] = [float(clean["x"].iloc[gaps[0] - 1]), float(clean["x"].iloc[gaps[-1] + 1])]
        clean = pp.handle_missing(clean, "x", "ffill")
        R["gaps_excel"] = (gaps + dl.EXCEL_FIRST_DATA_ROW).tolist()
        R["gaps_n"] = (gaps + 1).tolist()
        R["gap_filled"] = clean["x"].iloc[gaps].tolist()
        R["gap_true"] = n_true.tolist()
        known = ~typed["x"].isna().to_numpy()
        R["x_is_sqrt"] = bool(np.allclose(typed["x"].to_numpy()[known], np.sqrt(np.flatnonzero(known) + 1.0)))
        print(f"x: cast + ffill; заполнено {R['gap_filled']} (закон √n даёт {np.round(n_true, 4).tolist()})")
    left = dl.check_data_quality(clean, VARIANT_COLUMNS, warn=False)
    R["left_missing"] = sum(m["count"] for m in left["missing"].values())
    print("Пропусков после обработки:", R["left_missing"], "| типы:", clean.dtypes.astype(str).to_dict())
    data_df = clean[VARIANT_COLUMNS].copy()

    # ------------------------------------------------------------------ Задание 2
    header("Задание 2. Запись файлов с заголовками (+ обратное чтение, задание 1)")
    paths = io_utils.write_table(data_df, f"task2_variant{VARIANT}_clean")
    rt = {}
    for ext, p in paths.items():
        back = io_utils.read_table(p)
        ok = back.shape == data_df.shape and np.allclose(back.to_numpy(float), data_df.to_numpy(float))
        rt[ext] = {"file": os.path.basename(p), "shape": list(back.shape), "columns": list(back.columns),
                   "equal": bool(ok)}
        print(f"{ext}: {os.path.basename(p)} -> прочитано {back.shape}, столбцы {list(back.columns)}, "
              f"совпадает: {ok}")
    R["roundtrip"] = rt
    nh = {}
    for ext, p in io_utils.write_without_header(data_df, "task1_no_header").items():
        back = io_utils.read_table(p)
        nh[ext] = {"file": os.path.basename(p), "columns": list(back.columns), "shape": list(back.shape),
                   "equal": bool(np.allclose(back.to_numpy(float), data_df.to_numpy(float)))}
        print(f"Без заголовков, {ext}: столбцы {list(back.columns)}, форма {back.shape}")
    R["noheader"] = nh

    # ------------------------------------------------------------------ Задание 9
    header("Задание 9. DataFrame -> NumPy")
    X = pp.to_numpy(data_df)
    names = VARIANT_COLUMNS
    N = len(X)
    nn = np.arange(1, N + 1)
    print("X:", type(X).__name__, X.shape, X.dtype)
    R["X_shape"] = list(X.shape)

    # ------------------------------------------------------------------ Задания 10–11
    header("Задания 10–11. Собственное масштабирование и восстановление")
    scalers = {"minmax_0_1": pp.MinMaxScalerCustom(0, 1).fit(X),
               "minmax_m1_1": pp.MinMaxScalerCustom(-1, 1).fit(X),
               "zscore": pp.StandardScalerCustom().fit(X)}
    ppath = pp.save_scaler_params(scalers)
    saved = json.load(open(ppath, encoding="utf-8"))
    sc_res = {}
    for k, s in scalers.items():
        Z = s.transform(X)
        restored = pp.load_scaler(saved[k]).inverse_transform(Z)  # по параметрам из файла
        sc_res[k] = {"min": Z.min(0).tolist(), "max": Z.max(0).tolist(), "mean": Z.mean(0).tolist(),
                     "std": Z.std(0).tolist(), "inv_err": float(np.max(np.abs(restored - X)))}
        print(f"{k}: min {np.round(Z.min(0), 6)}, max {np.round(Z.max(0), 6)}, "
              f"mean {np.round(Z.mean(0), 6)}, std {np.round(Z.std(0), 6)}; "
              f"ошибка восстановления {sc_res[k]['inv_err']:.2e}")
    R["scale"] = sc_res
    R["scale_params"] = saved
    R["scale_params_file"] = os.path.basename(ppath)

    # ------------------------------------------------------------------ Задание 12
    header("Задание 12. Разбиение на 3 массива _train, _val, _test")
    s_pct = pp.split_train_val_test(X, (70, 15, 15), name="X")
    s_frac = pp.split_train_val_test(X, (0.7, 0.15, 0.15), name="X")
    s_150 = pp.split_train_val_test(np.arange(150), (1, 1, 1), name="a150")
    s_100 = pp.split_train_val_test(np.arange(100), (75, 10, 15), name="a100")
    s_sh = pp.split_train_val_test(X, (60, 20, 20), shuffle=True, name="Xs")
    R["split"] = {"pct": {k: len(v) for k, v in s_pct.items()},
                  "same": all(np.array_equal(s_pct[k], s_frac[k]) for k in s_pct),
                  "a150": {k: len(v) for k, v in s_150.items()},
                  "a100": {k: len(v) for k, v in s_100.items()},
                  "shuffle": {k: len(v) for k, v in s_sh.items()}}
    for k in ("pct", "a150", "a100", "shuffle"):
        print(k, R["split"][k])
    print("Проценты и доли дают одинаковое разбиение:", R["split"]["same"])

    figs = {}
    # ------------------------------------------------------------------ Задание 14
    header("Задание 14. График исходных данных")
    figs["series"] = save_figure(sa.plot_series(X, names, gaps), "task14_series")
    e = X[:, 1] - np.sqrt(nn)  # отклонение y от закона √n
    R["resid"] = {"mean": float(e.mean()), "std": float(e.std(ddof=1)), "min": float(e.min()),
                  "max": float(e.max()), "ac1": float(np.corrcoef(e[:-1], e[1:])[0, 1]),
                  "y_gt_x": int((X[:, 1] > X[:, 0]).sum())}
    print("Отклонение y − √n:", R["resid"])

    # ------------------------------------------------------------------ Задание 15
    header("Задание 15. Гистограммы и нормализованные гистограммы")
    figs["hist"] = save_figure(sa.plot_histograms(X, names), "task15_histograms")
    hb = {}
    for i, c in enumerate(names):
        h, ed = np.histogram(X[:, i], bins=15, density=True)
        hb[c] = {"first": float(h[0]), "last": float(h[-1]), "width": float(ed[1] - ed[0]),
                 "integral": float((h * np.diff(ed)).sum())}
    R["hist"] = hb
    print(hb)

    # ------------------------------------------------------------------ Задание 16
    header("Задание 16. Отсортированные по возрастанию столбцы")
    sorted_df = sa.sorted_columns(data_df)
    sorted_df.to_excel(timestamped_filename("task16_sorted_columns", "xlsx"), index=False)
    R["col_range"] = {c: [float(sorted_df[c].iloc[0]), float(sorted_df[c].iloc[-1])] for c in names}
    R["y_sorted_changed"] = int((sorted_df["y"].to_numpy() != data_df["y"].to_numpy()).sum())
    R["x_sorted_same"] = bool(np.array_equal(sorted_df["x"].to_numpy(), data_df["x"].to_numpy()))
    print("x уже упорядочен:", R["x_sorted_same"], "| в y изменили позицию:", R["y_sorted_changed"])

    # ------------------------------------------------------------------ Задание 17
    header("Задание 17. Эмпирическая функция распределения")
    figs["ecdf"] = save_figure(sa.plot_ecdf(X, names), "task17_ecdf")
    R["ecdf"] = {c: {"F_mid": float(np.mean(X[:, i] <= (X[:, i].min() + X[:, i].max()) / 2)),
                     "mid": float((X[:, i].min() + X[:, i].max()) / 2)} for i, c in enumerate(names)}
    print(R["ecdf"])

    # ------------------------------------------------------------------ Задания 18–19
    header("Задания 18–19. Статистики (NumPy и pandas) и доверительные интервалы (alpha = 0,05)")
    rows = []
    for i, c in enumerate(names):
        s = sa.column_stats(X[:, i])
        cm, cv = sa.ci_mean(X[:, i]), sa.ci_var(X[:, i])
        rows.append({"col": c, **s, "ci_mean": cm, "ci_var": cv})
    pst = sa.pandas_stats(data_df)
    R["np_pd_same"] = bool(np.allclose(pst["mean"], [r["mean"] for r in rows]) and
                           np.allclose(pst["var"], [r["var"] for r in rows]) and
                           np.allclose(pst["median"], [r["median"] for r in rows]))
    print(pd.DataFrame(rows).to_string(index=False))
    print("pandas:\n", pst, "\nNumPy и pandas совпадают:", R["np_pd_same"])
    pd.DataFrame(rows).to_excel(timestamped_filename("task18_19_stats", "xlsx"), index=False)
    R["stats"] = rows
    R["pd_mode_count"] = {c: int(data_df[c].mode().size) for c in names}
    figs["stats"] = save_figure(sa.plot_stats(X, names, rows), "task18_stats_plot")

    # ------------------------------------------------------------------ Задания 20–21
    header("Задания 20–21. Ковариация, корреляция, значимость")
    cov = pd.DataFrame(np.cov(X, rowvar=False), index=names, columns=names)
    corr = pd.DataFrame(np.corrcoef(X, rowvar=False), index=names, columns=names)
    assert np.allclose(corr, data_df.corr()) and np.allclose(cov, data_df.cov())
    print("Ковариация:\n", cov.round(5), "\nКорреляция:\n", corr.round(5))
    figs["corr"] = save_figure(sa.plot_corr(data_df, cov, corr), "task20_cov_corr")
    sig = sa.corr_significance(data_df)
    print(sig.to_string(index=False))
    with pd.ExcelWriter(timestamped_filename("task20_21_cov_corr", "xlsx")) as w:
        cov.to_excel(w, sheet_name="covariance")
        corr.to_excel(w, sheet_name="correlation")
        sig.to_excel(w, sheet_name="significance", index=False)
    R["cov"] = cov.to_dict()
    R["corr"] = corr.to_dict()
    R["sig"] = sig.to_dict("records")
    from scipy import stats as st_
    R["corr_e_x"] = [float(v) for v in st_.pearsonr(X[:, 0], e)]

    # ------------------------------------------------------------------ Задание 22
    header("Задание 22. Взаимная корреляция в зависимости от сдвига")
    xc = []
    R["xcorr"] = {}
    for label, u, v in [("x–y", X[:, 0], X[:, 1]), ("Δx–Δy", np.diff(X[:, 0]), np.diff(X[:, 1]))]:
        lags, c = sa.cross_correlation(u, v, 40)
        k = int(np.argmax(np.abs(c)))
        R["xcorr"][label] = {"lag": int(lags[k]), "value": float(c[k]), "at0": float(c[lags == 0][0]),
                             "at10": float(c[lags == 10][0]), "at40": float(c[lags == 40][0]),
                             "atm40": float(c[lags == -40][0]),
                             "max_abs_nonzero": float(np.max(np.abs(c[lags != 0])))}
        print(label, R["xcorr"][label])
        xc.append((label, lags, c))
    figs["xcorr"] = save_figure(sa.plot_xcorr(xc), "task22_xcorr")

    # ------------------------------------------------------------------ Задание 23
    header("Задание 23. Численная производная и градиент")
    dX = np.diff(X, axis=0)
    gX = np.gradient(X, axis=0)
    theory = 1 / (2 * np.sqrt(nn))
    inner = np.ones(N, bool)
    inner[[0, N - 1]] = False
    for g in gaps:  # окрестность заполненных пропусков
        inner[max(g - 1, 0):g + 2] = False
    R["grad"] = {"x_err_inner": float(np.max(np.abs(gX[inner, 0] - theory[inner]))),
                 "x_first": [float(v) for v in gX[:3, 0]], "x_last": float(gX[-1, 0]),
                 "x_at_gaps": [float(v) for v in gX[gaps[0] - 1:gaps[-1] + 2, 0]],
                 "y_mean": float(gX[:, 1].mean()), "y_std": float(gX[:, 1].std()),
                 "y_min": float(gX[:, 1].min()), "y_max": float(gX[:, 1].max()),
                 "y_neg": int((dX[:, 1] < 0).sum()), "diff_len": int(len(dX)),
                 "x_mean": float(gX[:, 0].mean())}
    print(R["grad"])
    figs["grad"] = save_figure(sa.plot_gradient(X, gX, names, theory), "task23_gradient")

    # ------------------------------------------------------------------ Задание 24
    header("Задание 24. Свёртка")
    cxy_np, cxy_sp = sa.convolution(X[:, 0], X[:, 1])
    w_ = 11
    ysm = np.convolve(X[:, 1], np.ones(w_) / w_, mode="valid")
    h_ = (w_ - 1) // 2
    trend = np.sqrt(nn) + e.mean()
    R["conv"] = {"len": len(cxy_np), "diff": float(np.max(np.abs(cxy_np - cxy_sp))),
                 "k_max": int(np.argmax(cxy_np)), "max": float(cxy_np.max()),
                 "c0": float(cxy_np[0]), "c_last": float(cxy_np[-1]), "width": w_,
                 "len_valid": len(ysm), "noise_before": float(np.std(X[:, 1] - trend)),
                 "noise_after": float(np.std(ysm - trend[h_:N - h_]))}
    print(R["conv"])
    figs["conv"] = save_figure(sa.plot_convolution(X[:, 0], X[:, 1], cxy_np, ysm, w_), "task24_convolution")

    # ------------------------------------------------------------------ Задания 25–26
    header("Задания 25–26. Скалярное и векторное произведения, нормы L1 и L2")
    dot = float(np.dot(X[:, 0], X[:, 1]))
    cos = dot / np.linalg.norm(X[:, 0]) / np.linalg.norm(X[:, 1])
    a3, b3 = X[:3, 0], X[:3, 1]
    cross = np.cross(a3, b3)
    norms = {c: {"L1": float(np.linalg.norm(X[:, i], 1)), "L2": float(np.linalg.norm(X[:, i], 2))}
             for i, c in enumerate(names)}
    R["products"] = {"dot": dot, "cos": float(cos), "angle_deg": float(np.degrees(np.arccos(cos))),
                     "a": a3.tolist(), "b": b3.tolist(), "cross": cross.tolist(),
                     "cross_orth": [float(np.dot(cross, a3)), float(np.dot(cross, b3))],
                     "dot_sum_check": float(np.sum(X[:, 0] * X[:, 1]))}
    R["norms"] = norms
    print(R["products"], norms)

    # ------------------------------------------------------------------ Задание 27
    header("Задание 27. Проверка гипотез о распределении")
    tests = sa.distribution_tests(data_df)
    print(tests.to_string(index=False))
    tests.to_excel(timestamped_filename("task27_distribution_tests", "xlsx"), index=False)
    R["tests"] = tests.to_dict("records")
    _, p_eu = st_.kstest(e, st_.uniform(loc=e.min(), scale=e.max() - e.min()).cdf)
    R["resid_tests"] = {"p_unif": float(p_eu), "p_shapiro": float(st_.shapiro(e).pvalue)}
    print("Отклонение y − √n:", R["resid_tests"])

    # ------------------------------------------------------------------ Задания 28–30
    header("Задания 28–30. Спектрограмма, периодограмма, АЧХ")
    figs["spectral"] = save_figure(sa.plot_spectral({"x": X[:, 0], "y": X[:, 1], "y − √n": e}),
                                   "task28_30_spectral")
    from scipy.signal import periodogram
    spec = {}
    for label, s in (("x", X[:, 0]), ("y", X[:, 1]), ("e", e)):
        f_, p_ = periodogram(s - s.mean())
        pp_ = p_[1:]
        k = int(np.argmax(pp_)) + 1
        amp = np.abs(np.fft.rfft(s - s.mean())) / N
        spec[label] = {"f_max": float(f_[k]), "share_f1": float(p_[1] / p_.sum()),
                       "share_f3": float(p_[1:4].sum() / p_.sum()), "max_to_mean": float(p_[k] / pp_.mean()),
                       "low_half": float(pp_[:len(pp_) // 2].sum() / pp_.sum()), "n_freq": int(len(pp_)),
                       "amp_f1": float(amp[1]), "amp_mean_high": float(amp[len(amp) // 2:].mean())}
    R["spectral"] = spec
    R["f1"] = float(1 / N)
    print(spec)

    # ------------------------------------------------------------------ Задания 31–32
    header("Задания 31–32. Кубическая интерполяция и интерполяция сплайнами")
    parts, R["interp"] = [], {}
    for i, c in enumerate(names):
        seg = X[:40, i]
        n_, n_new, y_cub, y_spl = sa.interpolate(seg, 10)
        y_sm = None
        if c == "y":
            from scipy.interpolate import UnivariateSpline
            y_sm = UnivariateSpline(n_, seg, k=3, s=len(seg) * e.var())(n_new)
        R["interp"][c] = {"n": len(seg), "n_new": len(n_new), "diff": float(np.max(np.abs(y_cub - y_spl))),
                          "err_sqrt": float(np.max(np.abs(y_cub - np.sqrt(n_new)))) if c == "x" else None,
                          "sm_vs_trend": float(np.std(y_sm - np.sqrt(n_new) - e.mean())) if y_sm is not None
                          else None}
        parts.append((c, n_, seg, n_new, y_cub, y_spl, y_sm))
    figs["interp"] = save_figure(sa.plot_interpolation(parts), "task31_32_interpolation")
    # восстановление пропусков x по известным точкам всего ряда
    from scipy.interpolate import CubicSpline, interp1d
    xr = typed["x"].to_numpy()
    kn = ~np.isnan(xr)
    rec_c = interp1d(nn[kn], xr[kn], kind="cubic")(gaps + 1.0)
    rec_s = CubicSpline(nn[kn], xr[kn])(gaps + 1.0)
    R["gap_interp"] = {"cubic": rec_c.tolist(), "spline": rec_s.tolist(),
                       "err_cubic": float(np.max(np.abs(rec_c - n_true))),
                       "err_spline": float(np.max(np.abs(rec_s - n_true))),
                       "err_ffill": float(np.max(np.abs(np.array(R.get("gap_filled", rec_c)) - n_true)))}
    print("Интерполяция", R["interp"], "\nВосстановление пропусков x:", R["gap_interp"])

    # ------------------------------------------------------------------ Задание 33
    header("Задание 33. Бинарные маски")
    masks_raw = {c: sa.binary_masks(X[:, i]) for i, c in enumerate(names)}
    Zs = pp.StandardScalerCustom().fit_transform(X)
    masks_z = {c: sa.binary_masks(Zs[:, i]) for i, c in enumerate(names)}
    R["masks"] = {c: {k: int(v.sum()) for k, v in masks_raw[c].items()} for c in names}
    R["masks_z"] = {c: {k: int(v.sum()) for k, v in masks_z[c].items()} for c in names}
    R["masks_z_bounds"] = {c: [float(X[masks_z[c]["-1 <= x <= 1"], i].min()),
                               float(X[masks_z[c]["-1 <= x <= 1"], i].max())] for i, c in enumerate(names)}
    R["z_first_pos"] = {c: int(np.argmax(masks_z[c]["x > 0"]) + 1) for c in names}
    print("Маски исходных столбцов:", R["masks"], "\nМаски z-оценок:", R["masks_z"])
    figs["masks"] = save_figure(sa.plot_masks(Zs, masks_z, names), "task33_masks")

    # ------------------------------------------------------------------ Задания 34–35
    header("Задания 34–35. Сравнение масштабирования со scikit-learn и инверсия")
    sc, full = {}, {}
    for i, c in enumerate(names):
        r = scaling.compare_scalers(X[:, [i]])
        full[c] = r
        sc[c] = {k_: {kk: v for kk, v in r[k_].items() if not isinstance(v, np.ndarray)} for k_ in r}
    figs["scaling"] = save_figure(scaling.plot_scaling(X, full, names), "task34_scaling")
    out = pd.DataFrame({"n": nn})
    for c in names:
        out[c] = data_df[c].to_numpy()
        for k_ in ("minmax", "standard"):
            out[f"{c}_{k_}_custom"] = full[c][k_]["custom"].ravel()
            out[f"{c}_{k_}_sklearn"] = full[c][k_]["sklearn"].ravel()
            out[f"{c}_{k_}_inv_custom"] = full[c][k_]["inv_custom"].ravel()
            out[f"{c}_{k_}_inv_sklearn"] = full[c][k_]["inv_sklearn"].ravel()
    io_utils.write_table(out, "task34_35_scaling")
    for c, v in sc.items():
        print(c, v)
    R["scaling_cmp"] = sc

    # ------------------------------------------------------------------ Задания 36–37
    header("Задания 36–37. Скользящее окно и скользящее среднее")
    y = X[:, 1]
    win = windows.sliding_window(y, 10)
    from numpy.lib.stride_tricks import sliding_window_view
    same = np.array_equal(win, sliding_window_view(y, 10))
    ma_widths = (5, 11, 25)
    noise_left, ma_pd = {}, {}
    for w in ma_widths:
        ma = windows.moving_average(y, w)
        h = (w - 1) // 2
        noise_left[w] = float(np.std(ma - trend[h:N - h]))
        ma_pd[w] = bool(np.allclose(ma, pd.Series(y).rolling(w).mean().dropna().to_numpy()))
    R["windows"] = {"shape": list(win.shape), "same": bool(same), "widths": list(ma_widths),
                    "first": win[0].tolist(), "noise_left": noise_left, "pandas_same": ma_pd,
                    "noise_theory": {w: float(np.std(y - trend) / np.sqrt(w)) for w in ma_widths},
                    "noise_before": float(np.std(y - trend))}
    print(R["windows"])
    figs["ma"] = save_figure(windows.plot_moving_average(y, ma_widths, "y"), "task37_moving_average")

    # ------------------------------------------------------------------ Задания 38–45
    header("Задания 38–45. Тензорные вычисления")
    R["tensors"] = tensors.run_tensors(X)

    # ------------------------------------------------------------------ Задания 46–49
    header("Задания 46–49. Матричные операции и пайплайн")
    R["sparse"] = matrix_ops.sparse_demo()
    R["inverse"] = matrix_ops.inverse_demo()
    F, views, R["F"] = matrix_ops.build_F(X, Zs)
    np.save(timestamped_filename("task48_F", "npy"), F)
    R["pca"] = matrix_ops.pca_pipeline(F, views)
    pd.DataFrame(R["pca"]["rows"]).to_excel(timestamped_filename("task49_pca_combinations", "xlsx"), index=False)
    print(pd.DataFrame(R["pca"]["rows"]).to_string(index=False))
    figs["pca"] = save_figure(matrix_ops.plot_pca(R["pca"]), "task49_pca")
    R["pca"].pop("proj")

    # ------------------------------------------------------------------ Задание 50
    header("Задание 50. Обработка изображения средствами SciPy")
    if not os.path.exists(config.IMAGE_PATH):
        image_proc.generate_image(config.IMAGE_PATH)
    R["image"] = image_proc.run_task50(config.IMAGE_PATH)

    R["figs"] = {k_: os.path.basename(v) for k_, v in figs.items()}
    with open(timestamped_filename("results", "json"), "w", encoding="utf-8") as f:
        json.dump(R, f, ensure_ascii=False, indent=1,
                  default=lambda o: o.tolist() if isinstance(o, np.ndarray) else
                  (o.item() if isinstance(o, np.generic) else str(o)))
    header("Все задания выполнены. Результаты — в каталоге output/")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=f"ЛР1, вариант {VARIANT}")
    parser.add_argument("--interactive", action="store_true", help="интерактивное меню задания 8")
    main(parser.parse_args().interactive)
