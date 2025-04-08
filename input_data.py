# -*- coding: utf-8 -*-
"""
Created on Fri Apr  4 09:08:12 2025

@author: angel
"""
import numpy as np

# Q = 300e6
Vfase = 100
a = np.exp(1j * 2 * np.pi / 3)
trinta = np.exp(1j * np.pi / 6)
# CUIDADO QUE ENTENDER ESSE Vabc PARECE, MAS NÄO É TRIVIAL
Vabc = Vfase * np.sqrt(3) * np.array([trinta, trinta*a**2, trinta*a])

X = 14#V**2 / Q

w = 2*np.pi*60
C = 1 / (w*X)
N = 14
Su = 3
P = 3  # parallel units in affected string
S = 4  # series groups line to neutral
Pa = 6 # parallel units per phase in "left" wye
Pt = 11 # parallel units per phase
Pst = Pt - Pa - P

Xfase = X/N * (Su/P) * (S/Pst)
Ifase = Vfase / Xfase
