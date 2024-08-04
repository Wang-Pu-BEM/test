
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import numpy as np
#import literation

def iterate_convergence(x, tolerance, max_iterations=100):
    def CM_(C0_f, C0_g, Ce_f, Ce_g, Ce_f_M):
        Ce_g_increase = (C0_g / (1000 - C0_f)) * (C0_f - Ce_f_M)
        Ce_g_M = Ce_g - Ce_g_increase
        Ce_f_increase = (C0_f / (1000 - C0_g)) * (C0_g - Ce_g_M)
        Ce_f_M = Ce_f - Ce_f_increase
        return Ce_f_M, Ce_g_M
    
    Ce_f_M, Ce_g_M = x[4], x[5]
    convergence_history = [(Ce_f_M, Ce_g_M)] 
    iterations = 0

    while iterations < max_iterations:
        Ce_f_M_prev, Ce_g_M_prev = Ce_f_M, Ce_g_M
        Ce_f_M, Ce_g_M = CM_(x[2], x[3], x[4], x[5], Ce_f_M)
        iterations += 1

        # 计算当前结果与上一次结果的差值
        diff_ef = abs(Ce_f_M - Ce_f_M_prev)
        diff_eg = abs(Ce_g_M - Ce_g_M_prev)

        # 检查差值是否在容差范围内
        if diff_ef < tolerance and diff_eg < tolerance or iterations >= max_iterations:
            return Ce_f_M, Ce_g_M, iterations, convergence_history
        
        convergence_history.append((Ce_f_M, Ce_g_M))

    raise RuntimeError("Convergence not achieved within the maximum number of iterations.")


def calculate_convergence(row, tolerance):
    result_Ce_f, result_Ce_g, iterations, convergence_history = iterate_convergence(row, tolerance)
    return pd.Series({'Ce_f_M': result_Ce_f, 'Ce_g_M': result_Ce_g, 'iterations':iterations, 'history':convergence_history})
def calculate_Qe(row):
    Qe_f = (row[2]-row[4])*row[0]/row[1]
    Qe_g = (row[3]-row[5])*row[0]/row[1]
    return pd.Series({'Qe_f': Qe_f, 'Qe_g': Qe_g})
def calculate_Qe_M(row):
    Qe_f_M = (row[2]-row[6])*row[0]/row[1]
    Qe_g_M = (row[3]-row[7])*row[0]/row[1]
    return pd.Series({'Qe_f_M': Qe_f_M, 'Qe_g_M': Qe_g_M})
def calculate_Selectivity(row):
    a = (row[11]*row[5])/(row[12]*row[4])
    return pd.Series({'a':a})

def execute_after_select(io):
    if io != '':
        print("选择的文件路径:", io)
        # 在这里执行你的代码块
        df = pd.read_excel(io)
        tolerance = 1e-5
        convergence_results = df.apply(calculate_convergence, axis=1, args=(tolerance,))
        df['Ce_f_M'] = convergence_results['Ce_f_M']
        df['Ce_g_M'] = convergence_results['Ce_g_M']
        df['迭代次数'] = convergence_results['iterations']
        Qe_result = df.apply(calculate_Qe, axis=1)
        Qe_result_M = df.apply(calculate_Qe_M, axis=1)
        df['Qe_f'] = Qe_result['Qe_f']
        df['Qe_g'] = Qe_result['Qe_g']
        df['Qe_f_M'] = Qe_result_M['Qe_f_M']
        df['Qe_g_M'] = Qe_result_M['Qe_g_M']
        a_result = df.apply(calculate_Selectivity, axis=1)
        df['α'] = a_result
        df['history'] = convergence_results['history']
        # 保存结果到新的Excel文件，也可以按需进行其他操作
        df.to_excel('output.xlsx', index=False)
        messagebox.showinfo(message='已执行并保存结果到 output.xlsx')
    else:
        messagebox.showinfo(message='您没有选择任何文件')


def resize(w_box, h_box, pil_image):        
    w, h = pil_image.size #获取图像的原始大小          
    f1 = 1.0*w_box/w        
    f2 = 1.0*h_box/h       
    factor = min([f1, f2])     
    width = int(w*factor)
    height = int(h*factor)       
    return pil_image.resize((width, height), Image.ANTIALIAS)