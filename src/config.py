"""Общие настройки: seed, пути, вспомогательные функции для сохранения файлов."""
import os
import random
from datetime import datetime

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

SEED = 42

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

DATASET_PATH = os.path.join(DATA_DIR, "Lab1_Dataset.xlsx")
IMAGE_PATH = os.path.join(DATA_DIR, "image.png")

# Вариант 7: столбцы M (x) и N (y) листа Sheet1
VARIANT = 7
VARIANT_COLUMNS = ["x", "y"]

# Метка времени фиксируется один раз за запуск, чтобы все файлы одного прогона
# имели одинаковый суффикс
RUN_TS = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

plt.rcParams.update({
    "figure.dpi": 100,
    "font.size": 10,
    "axes.grid": True,
    "grid.alpha": 0.3,
})


def set_seeds(seed: int = SEED) -> None:
    """Задание 13. Фиксирует seed для всех генераторов случайных чисел."""
    import tensorflow as tf
    import torch

    os.environ["PYTHONHASHSEED"] = str(seed)
    random.seed(seed)
    np.random.seed(seed)
    tf.random.set_seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)


def timestamped_filename(base: str, ext: str) -> str:
    """Формирует имя файла с датой и временем запуска."""
    return os.path.join(OUTPUT_DIR, f"{base}_{RUN_TS}.{ext}")


def save_figure(fig, base_name: str) -> str:
    """Задание 3. Сохраняет рисунок в PNG с меткой времени (до закрытия фигуры)."""
    path = timestamped_filename(base_name, "png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path
