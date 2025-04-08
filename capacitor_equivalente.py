# -*- coding: utf-8 -*-
"""
Created on Wed Apr  2 17:40:16 2025

@author: angel
"""

import numpy as np
from typing import Tuple


def equivalents_from_matrix(matrix: np.ndarray) -> Tuple[complex, np.ndarray]:

    collumn = 1 / np.sum(1 / matrix, axis=1)
    collumn = collumn.reshape(-1, 1)
    equivalent = complex(np.sum(collumn, axis=0))

    return equivalent, collumn

def parallel_impedance(Z1: complex, Z2: complex) -> complex:
    # Compute the equivalent impedance as the inverse of the sum of inverses
    return 1 / (1/Z1 + 1/Z2)
