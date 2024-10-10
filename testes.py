import openpyxl as xl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

df = pd.read_excel('dados\\Matriz Modelo - VERSÃO SISTEMA.xlsx', sheet_name='MATRIZ CONTRATOS')

df2 = df.copy()[6:]
print(df2)

#coluna 35
