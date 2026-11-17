"""Задание 50. Обработка изображения средствами SciPy."""
import matplotlib.pyplot as plt
import numpy as np
from scipy.ndimage import gaussian_filter, rotate
from sklearn.decomposition import PCA

from src.config import save_figure

# фрагмент (y0, y1, x0, x1) для увеличенного показа размытия и резкости: большой шар
CROP = (60, 220, 90, 230)


def generate_image(path: str, h: int = 384, w: int = 512) -> np.ndarray:
    """
    Строит цветное тестовое изображение «воздушные шары над полем»: голубое небо с
    градиентом, белые облака, пять разноцветных полосатых шаров с бликами и корзинами,
    холмы и поле с полосами посевов, дорога и деревья. Размер кратен 8 (для PCA по блокам 8×8).
    """
    rng = np.random.default_rng(11)
    Y, X = np.mgrid[0:h, 0:w].astype(np.float32)
    horizon = int(0.62 * h)

    # небо: от насыщенного синего вверху к светло-голубому у горизонта
    t = (Y / horizon).clip(0, 1)[..., None]
    img = (1 - t) * np.array([40, 110, 205]) + t * np.array([175, 215, 240])

    # облака — суммы гауссовых пятен, белые с серой нижней частью
    for cx, cy, sc in [(80, 70, 1.0), (300, 45, 1.3), (450, 120, 0.8), (210, 150, 0.7)]:
        dens = np.zeros((h, w), np.float32)
        for _ in range(9):
            ox, oy = rng.normal(0, 28 * sc), rng.normal(0, 8 * sc)
            r = rng.uniform(14, 26) * sc
            dens += np.exp(-(((X - cx - ox) ** 2 + (Y - cy - oy) ** 2) / (2 * r * r)))
        a = np.clip(dens * 1.4, 0, 1)[..., None]
        shade = np.clip((Y - cy) / (40 * sc), 0, 1)[..., None]
        cloud = (1 - shade) * np.array([252, 252, 252]) + shade * np.array([200, 205, 215])
        img = img * (1 - a) + cloud * a

    # холмы: дальний (сине-зелёный) и ближний (зелёный)
    far = horizon - 18 - 14 * np.sin(X / 70) - 8 * np.sin(X / 23 + 1)
    img[Y > far] = [96, 140, 120]
    near = horizon + 6 * np.sin(X / 55 + 2)
    field = Y > near
    # поле: полосы посевов, расходящиеся к зрителю, светлее и темнее
    stripe = np.sin((X - w / 2) / (6 + 0.25 * (Y - horizon).clip(1, None)) * 3)
    base = np.where(stripe[..., None] > 0, np.array([88, 150, 52]), np.array([60, 115, 40]))
    depth = ((Y - horizon) / (h - horizon)).clip(0, 1)[..., None]
    fcol = base * (0.85 + 0.3 * depth)
    img[field] = fcol[field]
    # жёлтое поле подсолнечника справа и дорога
    sun = field & (X > 0.68 * w + 0.5 * (Y - horizon))
    img[sun] = (np.array([235, 195, 45]) * (0.85 + 0.25 * depth))[sun]
    road = field & (np.abs(X - (0.3 * w + 0.9 * (Y - horizon))) < 3 + 0.12 * (Y - horizon))
    img[road] = [180, 160, 120]
    # деревья — тёмные круглые кроны с тенью
    for tx in rng.integers(20, w - 20, 9):
        ty = int(near[0, tx]) + int(rng.integers(4, 30))
        r = int(rng.integers(7, 13))
        crown = np.hypot(X - tx, Y - ty) < r
        img[crown] = [30, 70, 30]
        img[(Y > ty) & (Y < ty + r + 6) & (np.abs(X - tx) < 2)] = [70, 45, 25]

    # воздушные шары: полосатая оболочка, блик, стропы, корзина
    palettes = [([215, 40, 40], [250, 200, 40]), ([120, 50, 170], [245, 245, 245]),
                ([30, 140, 220], [250, 120, 30]), ([230, 90, 160], [255, 230, 90]),
                ([20, 160, 110], [250, 250, 250])]
    for (bx, by, br), (c1, c2) in zip([(150, 120, 44), (340, 170, 36), (430, 60, 26),
                                       (60, 200, 22), (260, 85, 18)], palettes):
        dx, dy = (X - bx) / br, (Y - by) / (br * 1.15)
        env = dx ** 2 + dy ** 2 < 1
        cone = (Y >= by) & (Y < by + 1.7 * br) & (np.abs(X - bx) < br * np.clip(1 - (Y - by) / (1.7 * br), 0, 1) ** 0.8)
        balloon = env | cone
        ang = np.arctan2(X - bx, np.sqrt(np.clip(br * br - (X - bx) ** 2, 1, None)))
        stripes = (np.floor(ang * 8 / np.pi) % 2 == 0)[..., None]
        col = np.where(stripes, np.array(c1), np.array(c2)).astype(np.float32)
        light = 1.15 - 0.5 * np.clip(dx + 0.2, -1, 1)[..., None] ** 2
        hl = np.exp(-(((X - bx + 0.35 * br) ** 2 + (Y - by + 0.4 * br) ** 2) / (0.06 * br * br)))[..., None]
        col = np.clip(col * light + 120 * hl, 0, 255)
        img[balloon] = col[balloon]
        by2 = int(by + 1.7 * br)
        for sx in (-0.28, 0.28):
            line = (Y > by2) & (Y < by2 + 0.35 * br) & (np.abs(X - (bx + sx * br * (1 - (Y - by2) / br))) < 0.8)
            img[line] = [60, 50, 40]
        bw = max(3, int(0.2 * br))
        img[by2 + int(0.35 * br):by2 + int(0.35 * br) + bw + 2, int(bx) - bw:int(bx) + bw] = [110, 70, 35]

    img = np.clip(img + rng.normal(0, 2.5, img.shape), 0, 255).astype(np.uint8)
    plt.imsave(path, img)
    return img


def _pair(img_a, img_b, title_a, title_b, base, cmap_b=None):
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.4))
    axes[0].imshow(img_a)
    axes[0].set_title(title_a)
    axes[1].imshow(img_b, cmap=cmap_b)
    axes[1].set_title(title_b)
    for ax in axes:
        ax.axis("off")
    fig.tight_layout()
    return save_figure(fig, base)


def split_channels(img):
    """1. Разделение цветовых каналов в отдельные массивы."""
    r, g, b = (img[..., i].copy() for i in range(3))
    fig, axes = plt.subplots(1, 4, figsize=(17, 3.6))
    axes[0].imshow(img)
    axes[0].set_title("Исходное")
    for ax, ch, cmap, name in zip(axes[1:], (r, g, b), ("Reds", "Greens", "Blues"),
                                  ("Красный", "Зелёный", "Синий")):
        ax.imshow(ch, cmap=cmap, vmin=0, vmax=255)  # чем темнее цвет, тем больше значение канала
        ax.set_title(f"{name} канал (среднее {ch.mean():.0f})")
    for ax in axes:
        ax.axis("off")
    fig.tight_layout()
    path = save_figure(fig, "task50_1_channels")
    return {"r": r, "g": g, "b": b}, path


def rotate_image(img, angle: float = 40.0):
    """2. Поворот на 40° против часовой стрелки (scipy.ndimage.rotate, холст расширяется)."""
    rotated = rotate(img.astype(np.float32), angle=angle, axes=(1, 0), reshape=True,
                     order=3, mode="constant", cval=0)
    rotated = np.clip(rotated, 0, 255).astype(np.uint8)
    path = _pair(img, rotated, f"Исходное {img.shape[1]}×{img.shape[0]}",
                 f"Поворот на {angle:.0f}° ({rotated.shape[1]}×{rotated.shape[0]})", "task50_2_rotated")
    return rotated, path


def to_grayscale(img):
    """3. Перевод в оттенки серого: Y = 0,299R + 0,587G + 0,114B."""
    gray = np.tensordot(img.astype(np.float32), np.array([0.299, 0.587, 0.114]), axes=([2], [0]))
    gray = np.clip(np.round(gray), 0, 255).astype(np.uint8)
    path = _pair(img, gray, "Исходное", "Чёрно-белое (оттенки серого)", "task50_3_gray", cmap_b="gray")
    return gray, path


def gray_histogram(gray):
    """4. Гистограмма яркостей: 256 интервалов шириной 1, левая граница = уровень яркости."""
    hist, edges = np.histogram(gray.ravel(), bins=256, range=(0, 256))
    fig, ax = plt.subplots(figsize=(11, 4))
    ax.bar(edges[:-1], hist, width=1.0, align="edge", color="black")
    ax.set_xlim(0, 256)
    ax.set_xticks(range(0, 257, 32))
    ax.set_xlabel("Яркость")
    ax.set_ylabel("Число пикселей")
    ax.set_title("Гистограмма чёрно-белого изображения")
    fig.tight_layout()
    path = save_figure(fig, "task50_4_gray_hist")
    return {"hist": hist, "edges": edges}, path


def split_by_histogram(gray, hist_info):
    """5. Три изображения по порогам равных долей пикселей (накопленная гистограмма)."""
    cum = np.cumsum(hist_info["hist"])
    total = cum[-1]
    t1 = int(np.searchsorted(cum, total / 3))
    t2 = int(np.searchsorted(cum, 2 * total / 3))
    parts = [np.where(gray <= t1, gray, 0), np.where((gray > t1) & (gray <= t2), gray, 0),
             np.where(gray > t2, gray, 0)]
    shares = [float(np.mean(gray <= t1)), float(np.mean((gray > t1) & (gray <= t2))),
              float(np.mean(gray > t2))]
    fig, axes = plt.subplots(1, 3, figsize=(16, 4.2))
    titles = [f"Тёмные: 0…{t1}", f"Средние: {t1 + 1}…{t2}", f"Светлые: {t2 + 1}…255"]
    for ax, part, title, s in zip(axes, parts, titles, shares):
        ax.imshow(part, cmap="gray", vmin=0, vmax=255)
        ax.set_title(f"{title} ({s:.1%} пикселей)")
        ax.axis("off")
    fig.tight_layout()
    path = save_figure(fig, "task50_5_split3")
    return {"t1": t1, "t2": t2, "shares": shares}, path


def add_circular_frame(img, width: int = 10):
    """6. Чёрная круговая рамка: всё за пределами круга радиуса R - width зануляется."""
    h, w = img.shape[:2]
    Y, X = np.ogrid[:h, :w]
    R = min(h, w) / 2.0
    outside = (X - w / 2) ** 2 + (Y - h / 2) ** 2 > (R - width) ** 2
    framed = img.copy()
    framed[outside] = 0
    path = _pair(img, framed, "Исходное", f"В чёрной круговой рамке (R = {R - width:.0f} пикс.)",
                 "task50_6_frame")
    return framed, path


def add_noise(img, sigma: float = 25.0):
    """7. Гауссовский шум N(0, sigma)."""
    noisy = np.clip(img.astype(np.float32) + np.random.normal(0, sigma, img.shape), 0, 255).astype(np.uint8)
    path = _pair(img, noisy, "Исходное", f"С гауссовским шумом (σ = {sigma:.0f})", "task50_7_noisy")
    return noisy, path


def gaussian_blur(img, sigma: float = 2.0):
    """8. Размытие фильтром Гаусса по каждому каналу (по оси каналов sigma = 0)."""
    blurred = gaussian_filter(img.astype(np.float32), sigma=(sigma, sigma, 0))
    blurred = np.clip(blurred, 0, 255).astype(np.uint8)
    path = _crop_compare(img, blurred, "Исходное", f"Размытие по Гауссу (σ = {sigma:.0f})", "task50_8_blurred")
    return blurred, path


def sharpen_image(img, sigma: float = 2.0, amount: float = 1.5):
    """9. Повышение резкости (нерезкое маскирование): img + amount·(img − blur(img))."""
    f = img.astype(np.float32)
    blurred = gaussian_filter(f, sigma=(sigma, sigma, 0))
    sharpened = np.clip(f + amount * (f - blurred), 0, 255).astype(np.uint8)
    path = _crop_compare(img, sharpened, "Исходное", f"Повышение резкости (k = {amount})",
                         "task50_9_sharpened")
    return sharpened, path


def _crop_compare(a, b, ta, tb, base):
    """Полные изображения и увеличенный фрагмент (воздушный шар), где эффект заметнее."""
    fig, axes = plt.subplots(2, 2, figsize=(11, 8.2))
    y0, y1, x0, x1 = CROP
    for col, (im, t) in enumerate([(a, ta), (b, tb)]):
        axes[0, col].imshow(im)
        axes[0, col].add_patch(plt.Rectangle((x0, y0), x1 - x0, y1 - y0, fill=False, ec="yellow", lw=1.5))
        axes[0, col].set_title(t)
        axes[1, col].imshow(im[y0:y1, x0:x1], interpolation="nearest")
        axes[1, col].set_title(f"{t}: увеличенный фрагмент")
    for ax in axes.ravel():
        ax.axis("off")
    fig.tight_layout()
    return save_figure(fig, base)


def _to_blocks(img, b):
    """(H, W, C) -> (число блоков, b·b): каждый блок b×b каждого канала — один объект."""
    h, w, c = img.shape
    x = img.reshape(h // b, b, w // b, b, c).transpose(4, 0, 2, 1, 3)
    return x.reshape(-1, b * b)


def _from_blocks(blocks, shape, b):
    h, w, c = shape
    x = blocks.reshape(c, h // b, w // b, b, b).transpose(1, 3, 2, 4, 0)
    return x.reshape(h, w, c)


def apply_pca_to_image(img, n_components: int = 8, block: int = 8):
    """
    10. PCA к изображению. Объект — блок 8×8 одного канала (64 признака).
    Восстановление по n_components главным компонентам, кривая ошибки, вид компонент.
    """
    flat = _to_blocks(img.astype(np.float32), block)
    pca = PCA(n_components=n_components).fit(flat)
    restored = _from_blocks(pca.inverse_transform(pca.transform(flat)), img.shape, block)
    restored = np.clip(restored, 0, 255).astype(np.uint8)
    mse = float(np.mean((img.astype(np.float32) - restored) ** 2))
    psnr = float(10 * np.log10(255 ** 2 / mse))
    paths = [_pair(img, restored, "Исходное",
                   f"PCA: {n_components} из {block * block} компонент (ср. квадрат ошибки {mse:.1f})",
                   "task50_10_pca_restored")]

    full = PCA().fit(flat)
    ks = np.arange(1, block * block + 1)
    mse_curve = []
    for k in ks:
        z = full.transform(flat)[:, :k]
        rec = z @ full.components_[:k] + full.mean_
        mse_curve.append(float(np.mean((flat - rec) ** 2)))
    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.semilogy(ks, np.clip(mse_curve, 1e-6, None), marker="o", ms=3)
    ax.axvline(n_components, color="C3", ls="--", label=f"k = {n_components}")
    ax.set_xlabel("Число главных компонент k")
    ax.set_ylabel("Среднеквадратичная ошибка (лог. шкала)")
    ax.set_title("Ошибка восстановления изображения от числа компонент")
    ax.legend()
    fig.tight_layout()
    paths.append(save_figure(fig, "task50_10_pca_mse"))

    fig, axes = plt.subplots(4, 4, figsize=(7.5, 8))
    for i, ax in enumerate(axes.ravel()):
        ax.imshow(full.components_[i].reshape(block, block), cmap="gray", interpolation="nearest")
        ax.set_title(f"ГК {i + 1}: {full.explained_variance_ratio_[i]:.1%}", fontsize=9)
        ax.axis("off")
    fig.suptitle("Первые 16 главных компонент (блоки 8×8)")
    fig.tight_layout()
    paths.append(save_figure(fig, "task50_10_pca_components"))

    evr = np.cumsum(full.explained_variance_ratio_)
    return {"mse": mse, "psnr": psnr, "evr_k": float(evr[n_components - 1]),
            "evr1": float(full.explained_variance_ratio_[0]),
            "comp1_range": [float(np.abs(full.components_[0]).min()), float(np.abs(full.components_[0]).max())],
            "comp1_same_sign": bool(np.all(full.components_[0] > 0) or np.all(full.components_[0] < 0)),
            "k95": int(np.searchsorted(evr, 0.95) + 1), "k99": int(np.searchsorted(evr, 0.99) + 1),
            "mse_curve": mse_curve, "n_objects": int(flat.shape[0])}, paths


def run_task50(image_path: str) -> dict:
    """Последовательно выполняет все 10 подзаданий."""
    img = plt.imread(image_path)
    if img.dtype != np.uint8:
        img = np.clip(img * 255, 0, 255).astype(np.uint8)
    img = img[..., :3]
    res, figs = {"shape": list(img.shape)}, {}

    ch, figs["channels"] = split_channels(img)
    res["channel_means"] = [float(ch[k].mean()) for k in "rgb"]
    rotated, figs["rotated"] = rotate_image(img, 40.0)
    res["rotated_shape"] = list(rotated.shape)
    gray, figs["gray"] = to_grayscale(img)
    res["gray_mean"] = float(gray.mean())
    hist, figs["hist"] = gray_histogram(gray)
    h = hist["hist"]
    res["hist_peak"] = int(np.argmax(h))
    res["hist_stats"] = {"lt20": float(h[:20].sum() / h.sum()), "lt60": float(h[:60].sum() / h.sum()),
                         "gt150": float(h[151:].sum() / h.sum()), "gt200": float(h[201:].sum() / h.sum()),
                         "min": int(gray.min()), "max": int(gray.max()),
                         "top_bins": [int(i) for i in np.argsort(h)[::-1][:5]]}
    res["split"], figs["split"] = split_by_histogram(gray, hist)
    _, figs["frame"] = add_circular_frame(img)
    noisy, figs["noisy"] = add_noise(img, 25.0)
    res["noise_mse"] = float(np.mean((noisy.astype(float) - img) ** 2))
    blurred, figs["blurred"] = gaussian_blur(img, 2.0)
    sharp, figs["sharpened"] = sharpen_image(img, 2.0, 1.5)
    # средний модуль градиента яркости — мера резкости
    grad = lambda a: float(np.mean(np.abs(np.diff(a.astype(float).mean(axis=2), axis=1))))
    res["sharpness"] = {"orig": grad(img), "blurred": grad(blurred), "sharpened": grad(sharp)}
    res["pca"], figs["pca"] = apply_pca_to_image(img, 8)
    res["figs"] = figs
    print("Задание 50 выполнено. Рисунки сохранены в output/.")
    return res
