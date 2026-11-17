"""Задания 38–45. Тензоры TensorFlow, PyTorch и Keras."""
import os

os.environ.setdefault("TF_CPP_MIN_LOG_LEVEL", "2")

import keras
import numpy as np
import tensorflow as tf
import torch
from keras import ops


def run_tensors(data: np.ndarray, k: int = 5, p: int = 4) -> dict:
    res = {}

    # Задание 38. Тензор TensorFlow
    X_tf = tf.constant(data, dtype=tf.float32)
    # Задание 39. Тензор PyTorch
    X_pt = torch.tensor(data, dtype=torch.float32)
    print("Задание 38. X_tf:", X_tf.shape, X_tf.dtype)
    print("Задание 39. X_pt:", tuple(X_pt.shape), X_pt.dtype)
    res["X_shape"] = list(X_tf.shape)
    res["tf_pt_equal"] = bool(np.allclose(X_tf.numpy(), X_pt.numpy()))

    # Задание 40. Случайные тензоры: A (k, n) целочисленный, W (m, p) нормальный, B (k, p)
    n, m = X_tf.shape
    A = tf.random.uniform((k, n), minval=0, maxval=10, dtype=tf.int32)
    W = tf.random.normal((m, p))
    B = tf.random.uniform((k, p))
    print(f"Задание 40. A{tuple(A.shape)} int32, W{tuple(W.shape)} N(0,1), B{tuple(B.shape)} U(0,1)")

    # Задание 41. A·X·W + B в TensorFlow
    result_tf = tf.matmul(tf.matmul(tf.cast(A, tf.float32), X_tf), W) + B
    print("Задание 41. AXW+B (TensorFlow), форма", tuple(result_tf.shape))
    print(np.round(result_tf.numpy(), 2))

    # Задание 42. A·X·W + B в PyTorch (свои случайные тензоры)
    A_pt = torch.randint(0, 10, (k, n)).to(torch.float32)
    W_pt = torch.randn(m, p)
    B_pt = torch.rand(k, p)
    result_pt = A_pt @ X_pt @ W_pt + B_pt
    print("Задание 42. AXW+B (PyTorch), форма", tuple(result_pt.shape))
    print(np.round(result_pt.numpy(), 2))

    # Проверка: если подать в PyTorch те же A, W, B, результаты совпадают
    same_pt = (torch.tensor(tf.cast(A, tf.float32).numpy()) @ X_pt
               @ torch.tensor(W.numpy()) + torch.tensor(B.numpy()))
    rel = float(np.max(np.abs(same_pt.numpy() - result_tf.numpy())) / np.max(np.abs(result_tf.numpy())))
    print(f"Одинаковые A, W, B в TF и PyTorch: максимальная относительная разность {rel:.2e}")
    res["axwb_tf"] = result_tf.numpy().tolist()
    res["axwb_pt"] = result_pt.numpy().tolist()
    res["axwb_same_inputs_rel_diff"] = rel

    # Задание 43. Матрицы T, P, Q формы (3, 10)
    # матрицы отличаются распределением элементов; Q >= 0 — под квадратным корнем
    T = np.random.uniform(-np.pi, np.pi, (3, 10))
    P = np.random.normal(0, 0.5, (3, 10))
    Q = np.random.uniform(0, 4, (3, 10))
    print("Задание 43. T ~ U(−π, π), P ~ N(0; 0,5), Q ~ U(0, 4); формы", T.shape, P.shape, Q.shape)

    # Задание 44. V = |sin(T) - exp(P)·sqrt(Q)| в Keras
    V_k = ops.abs(ops.sin(ops.convert_to_tensor(T)) -
                  ops.exp(ops.convert_to_tensor(P)) * ops.sqrt(ops.convert_to_tensor(Q)))
    V_k = keras.ops.convert_to_numpy(V_k)
    # Задание 45. То же в PyTorch
    V_t = torch.abs(torch.sin(torch.tensor(T)) - torch.exp(torch.tensor(P)) * torch.sqrt(torch.tensor(Q)))
    V_t = V_t.numpy()
    diff = float(np.max(np.abs(V_k - V_t)))
    print("Задание 44. V (Keras), первая строка:", np.round(V_k[0], 4))
    print("Задание 45. V (PyTorch), первая строка:", np.round(V_t[0], 4))
    print(f"Максимальная разность Keras и PyTorch: {diff:.2e}")
    res["V_keras"] = V_k.tolist()
    res["V_torch"] = V_t.tolist()
    res["V_diff"] = diff
    res["V_range"] = [float(V_k.min()), float(V_k.max())]
    res["TPQ_range"] = {k: [float(v.min()), float(v.max())] for k, v in (("T", T), ("P", P), ("Q", Q))}
    res["TPQ_distinct"] = bool(not (np.allclose(T, P) or np.allclose(T, Q) or np.allclose(P, Q)))
    return res
