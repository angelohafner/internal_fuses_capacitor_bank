# -*- coding: utf-8 -*-
"""
Created on Fri Apr  4 09:08:12 2025

@author: angel
"""
import numpy as np

# Q = 300e6
Vfase = 34.5e3 / np.sqrt(3)
a = np.exp(1j * 2 * np.pi / 3)
trinta = np.exp(1j * np.pi / 6)
# CUIDADO QUE ENTENDER ESSE Vabc PARECE, MAS NÄO É TRIVIAL
Vabc = Vfase * np.sqrt(3) * np.array([trinta, trinta*a**2, trinta*a])

w = 2*np.pi*60

N = 14
Su = 3
P = 3  # parallel units in affected string
S = 4  # series groups line to neutral
Pa = 6 # parallel units per phase in "left" wye
Pt = 11 # parallel units per phase
Pst = Pt - Pa - P

pot_reativa_banco = 10e6;
corrente_fase = (pot_reativa_banco/3) / Vfase
reatancia_banco = Vfase / corrente_fase
reatancia_celula = reatancia_banco * Pt / S
reatancia_elemento =  reatancia_celula * N / Su
X = reatancia_elemento






