import numpy as np
import pandas as pd


class DirFactor:
    c_dir = 1


class SeasonFactor:
    c_season = 1


class OrographyFactor:
    c_o = 1


class TurbulenceFactor:
    K_I = 1


class PcParapet:
    h = 7.2  # m
    length = 30.8  # m
    solidity_ratio = 1   # can be between 0 and 1 "
    zone = "without return corners"
    ratio = length / h

    """without return corners"""
    pc_for_3 = [2.3, 1.4, 1.2, 1.2]
    pc_for_5 = [2.9, 1.8, 1.4, 1.2]
    slope_3_5 = []
    for i in range(0, 4):
        slope_3_5 += [(pc_for_5[i]-pc_for_3[i])/(5-3)]

    pc_for_10 = [3.4, 2.1, 1.7, 1.2]
    slope_5_10 = []
    for j in range(0, 4):
        slope_5_10 += [(pc_for_10[j]-pc_for_5[j])/(10-5)]

    """with return corners"""
    p_c_h = [2.1, 1.8, 1.4, 1.2]
    slope_0_h = []
    for k in range(0, 4):
        slope_0_h += [(p_c_h[k] - 0)/h]

    p_c = []
    if solidity_ratio == 0.8:
        p_c += [1.2, 1.2, 1.2, 1.2]
    elif solidity_ratio == 1:
        if zone == "without return corners":
            if ratio < 3 or ratio == 3:
                p_c += pc_for_3
            elif ratio == 5:
                p_c += pc_for_5
            elif 3 < ratio < 5:
                for i in range(0, 4):
                    p_c += [slope_3_5[i] * (ratio - 3) + pc_for_3[i]]
            elif ratio > 10 or ratio == 10:
                p_c += pc_for_10
            elif 5 < ratio < 10:
                for i in range(0, 4):
                    p_c += [slope_5_10[i] * (ratio - 5) + pc_for_5[i]]
        elif zone == "with return corners":
            if length > h or length == h:
                p_c += p_c_h
            else:
                for i in range(0, 4):
                    p_c += [(slope_0_h[i] * (length - 0) + 0)]


def ShapeFactor(h,d):

    ratio = h / d
    Cpe10_Zone = []
    Ratio_5 = [-1.2, -0.8, -0.5, 0.8, -0.7, 1]
    Ratio_1 = [-1.2, -0.8, -0.5, 0.8, -0.5, 0.85]
    Slope_1And5 = []

    for i in range(0, 6):
        Slope_1And5 += [(Ratio_5[i] - Ratio_1[i]) / (5 - 1)]

    Ratio_0_25 = [-1.2, -0.8, -0.5, 0.7, -0.3, 0.85]
    Slope_0_25And1 = []

    for i in range(0, 6):
        Slope_0_25And1 += [(Ratio_1[i] - Ratio_0_25[i]) / (1 - 0.25)]

    for i in range(0, 1):
        if ratio == 5:
            Cpe10_Zone = [[ratio], [-1.2], [-0.8], [-0.5], [0.8], [-0.7], [1], [1.5]]
        elif ratio == 1:
            Cpe10_Zone = [[ratio], [-1.2], [-0.8], [-0.5], [0.8], [-0.5], [0.85], [1.11]]
        elif ratio < 0.25 or ratio == 0.25:
            Cpe10_Zone = [[ratio], [-1.2], [-0.8], [-0.5], [0.7], [-0.3], [0.85], [0.85]]
        elif 1 < ratio < 5:
            Cpe10_Zone = [[ratio]]
            for j in range(0, 6):
                Cpe10_Zone += [[Slope_1And5[j] * (ratio - 1) + Ratio_1[j]]]
            for k in range(0, 1):
                Cpe10_Zone += [[(Cpe10_Zone[4][0] - Cpe10_Zone[5][0]) * Cpe10_Zone[6][0]]]
        elif 0.25 < ratio < 1:
            Cpe10_Zone = [[ratio]]
            for n in range(0, 6):
                Cpe10_Zone += [[Slope_0_25And1[n] * (ratio - 0.25) + Ratio_0_25[n]]]

            for m in range(0, 1):
                Cpe10_Zone += [[(Cpe10_Zone[4][0] - Cpe10_Zone[5][0]) * Cpe10_Zone[6][0]]]
    
    Cpe1 = np.zeros(5)
    
    Cpe1[0] = -1.4 # Zone A
    Cpe1[1] = -1.1 # Zone B
    Cpe1[2] = -0.5 # Zone C
    Cpe1[3] = 1 # Zone D
    Cpe1[4] = Cpe10_Zone[5][0] # Zone E
    
        
  
    
    return Cpe10_Zone, Cpe1


def Cal_Cpe(Cpe1,Cpe10,Area):
    Cpe = np.zeros(np.size(Area))
    array1 = np.array(Cpe10)
    Cpe10 = array1[1:np.size(Area)+1,0]
    
    for i in range(0,np.size(Area)):
        if (Area[i] < 1) or (Area[i] == 1):
            Cpe[i] = Cpe1[i]
        elif (Area[i] > 10) or (Area[i] == 10):
            Cpe[i] = Cpe10[i]
        else:
            
            Cpe[i] = Cpe1[i]-(Cpe1[i]-Cpe10[i])*np.log10(Area[i])
    
    return Cpe

def Cal_Cpe_roof(Cpe1,Cpe10,Area):
    Cpe = np.zeros(np.size(Area)) 
    for i in range(0,np.size(Area)):
        if (Area[i] < 1) or (Area[i] == 1):
            Cpe[i] = Cpe1[i]
        elif (Area[i] > 10) or (Area[i] == 10):
            Cpe[i] = Cpe10[i]
        else:
            
            Cpe[i] = Cpe1[i]-(Cpe1[i]-Cpe10[i])*np.log10(Area[i])
    
    return Cpe


def Duopitch_Cpe10_0(alpha):
    
    """
    Input:
        direction: 0 degree
        alpha: Pitch angle
    """

    data_Cpe10_0_case1 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-0.6, -1.1, -2.5, -2.3, -1.7, -0.9, -0.5, -0, 0.7, 0.8],
        "Zone G": [-0.6, -0.8, -1.3, -1.2, -1.2, -0.8, -0.5, -0, 0.7, 0.8],
        "Zone H": [-0.8, -0.8, -0.9, -0.8, -0.6, -0.3, -0.2, -0, 0.7, 0.8],
        "Zone I": [-0.7, -0.6, -0.5, 0.2, -0.6, -0.4, -0.4, -0.2, -0.2, -0.2],
        "Zone J": [-1, -0.8, -0.7, 0.2, 0.2, -1, -0.5, -0.3, -0.3, -0.3]
        }  
    
    data_Cpe10_0_case2 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-0.6, -1.1, -2.5, -2.3, -1.7, -0.9, -0.5, -0, 0.7, 0.8],
        "Zone G": [-0.6, -0.8, -1.3, -1.2, -1.2, -0.8, -0.5, -0, 0.7, 0.8],
        "Zone H": [-0.8, -0.8, -0.9, -0.8, -0.6, -0.3, -0.2, -0, 0.7, 0.8],
        "Zone I": [-0.7, -0.6, -0.5, -0.6, -0.6, 0, 0, 0, -0.2, -0.2],
        "Zone J": [-1, -0.8, -0.7, -0.6, -0.6, 0, 0, 0, -0.3, -0.3]
        }
    
    data_Cpe10_0_case3 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-0.6, -1.1, -2.5, -2.3, 0, 0.2, 0.7, 0.7, 0.7, 0.8],
        "Zone G": [-0.6, -0.8, -1.3, -1.2, 0, 0.2, 0.7, 0.7, 0.7, 0.8],
        "Zone H": [-0.8, -0.8, -0.9, -0.8, 0, 0.2, 0.4, 0.6, 0.7, 0.8],
        "Zone I": [-0.7, -0.6, -0.5, 0.2, -0.6, -0.4, -0.4, -0.2, -0.2, -0.2],
        "Zone J": [-1, -0.8, -0.7, 0.2, 0.2, -1, -0.5, -0.3, -0.3, -0.3]
        }  
    data_Cpe10_0_case4 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-0.6, -1.1, -2.5, -2.3, 0, 0.2, 0.7, 0.7, 0.7, 0.8],
        "Zone G": [-0.6, -0.8, -1.3, -1.2, 0, 0.2, 0.7, 0.7, 0.7, 0.8],
        "Zone H": [-0.8, -0.8, -0.9, -0.8, 0, 0.2, 0.4, 0.6, 0.7, 0.8],
        "Zone I": [-0.7, -0.6, -0.5, -0.6, -0.6, 0, 0, 0, -0.2, -0.2],
        "Zone J": [-1, -0.8, -0.7, -0.6, -0.6, 0, 0, 0, -0.3, -0.3]
        }
    
    data_Cpe10_0_case1 = pd.DataFrame(data_Cpe10_0_case1)
    data_Cpe10_0_case2 = pd.DataFrame(data_Cpe10_0_case2)
    data_Cpe10_0_case3 = pd.DataFrame(data_Cpe10_0_case3)
    data_Cpe10_0_case4 = pd.DataFrame(data_Cpe10_0_case4)

    Cpe10_0_real_case1 = np.zeros(5) # Cpe10 for zone F, G, H, I, J
    Cpe10_0_real_case2 = np.zeros(5) # Cpe10 for zone F, G, H, I, J
    Cpe10_0_real_case3 = np.zeros(5) # Cpe10 for zone F, G, H, I, J
    Cpe10_0_real_case4 = np.zeros(5) # Cpe10 for zone F, G, H, I, J
    for i in range(0, len(data_Cpe10_0_case1["Pitch angle"])):
        if (alpha == data_Cpe10_0_case1["Pitch angle"][i]):
            Cpe10_0_real_case1[0] = data_Cpe10_0_case1["Zone F"][i]
            Cpe10_0_real_case1[1] = data_Cpe10_0_case1["Zone G"][i]
            Cpe10_0_real_case1[2] = data_Cpe10_0_case1["Zone H"][i]
            Cpe10_0_real_case1[3] = data_Cpe10_0_case1["Zone I"][i]
            Cpe10_0_real_case1[4] = data_Cpe10_0_case1["Zone J"][i]

            Cpe10_0_real_case2[0] = data_Cpe10_0_case2["Zone F"][i]
            Cpe10_0_real_case2[1] = data_Cpe10_0_case2["Zone G"][i]
            Cpe10_0_real_case2[2] = data_Cpe10_0_case2["Zone H"][i]
            Cpe10_0_real_case2[3] = data_Cpe10_0_case2["Zone I"][i]
            Cpe10_0_real_case2[4] = data_Cpe10_0_case2["Zone J"][i]

            Cpe10_0_real_case3[0] = data_Cpe10_0_case3["Zone F"][i]
            Cpe10_0_real_case3[1] = data_Cpe10_0_case3["Zone G"][i]
            Cpe10_0_real_case3[2] = data_Cpe10_0_case3["Zone H"][i]
            Cpe10_0_real_case3[3] = data_Cpe10_0_case3["Zone I"][i]
            Cpe10_0_real_case3[4] = data_Cpe10_0_case3["Zone J"][i]

            Cpe10_0_real_case4[0] = data_Cpe10_0_case4["Zone F"][i]
            Cpe10_0_real_case4[1] = data_Cpe10_0_case4["Zone G"][i]
            Cpe10_0_real_case4[2] = data_Cpe10_0_case4["Zone H"][i]
            Cpe10_0_real_case4[3] = data_Cpe10_0_case4["Zone I"][i]
            Cpe10_0_real_case4[4] = data_Cpe10_0_case4["Zone J"][i]
            break

        if (alpha > data_Cpe10_0_case1["Pitch angle"][i]) and (alpha < data_Cpe10_0_case1["Pitch angle"][i+1]):
            Cpe10_0_real_case1[0] = data_Cpe10_0_case1["Zone F"][i] + (data_Cpe10_0_case1["Zone F"][i+1] - data_Cpe10_0_case1["Zone F"][i]) * (alpha - data_Cpe10_0_case1["Pitch angle"][i]) / (data_Cpe10_0_case1["Pitch angle"][i+1] - data_Cpe10_0_case1["Pitch angle"][i])
            Cpe10_0_real_case1[1] = data_Cpe10_0_case1["Zone G"][i] + (data_Cpe10_0_case1["Zone G"][i+1] - data_Cpe10_0_case1["Zone G"][i]) * (alpha - data_Cpe10_0_case1["Pitch angle"][i]) / (data_Cpe10_0_case1["Pitch angle"][i+1] - data_Cpe10_0_case1["Pitch angle"][i])
            Cpe10_0_real_case1[2] = data_Cpe10_0_case1["Zone H"][i] + (data_Cpe10_0_case1["Zone H"][i+1] - data_Cpe10_0_case1["Zone H"][i]) * (alpha - data_Cpe10_0_case1["Pitch angle"][i]) / (data_Cpe10_0_case1["Pitch angle"][i+1] - data_Cpe10_0_case1["Pitch angle"][i])
            Cpe10_0_real_case1[3] = data_Cpe10_0_case1["Zone I"][i]+ (data_Cpe10_0_case1["Zone I"][i+1] - data_Cpe10_0_case1["Zone I"][i]) * (alpha - data_Cpe10_0_case1["Pitch angle"][i]) / (data_Cpe10_0_case1["Pitch angle"][i+1] - data_Cpe10_0_case1["Pitch angle"][i])
            Cpe10_0_real_case1[4] = data_Cpe10_0_case1["Zone J"][i]+ (data_Cpe10_0_case1["Zone J"][i+1] - data_Cpe10_0_case1["Zone J"][i]) * (alpha - data_Cpe10_0_case1["Pitch angle"][i]) / (data_Cpe10_0_case1["Pitch angle"][i+1] - data_Cpe10_0_case1["Pitch angle"][i])

            Cpe10_0_real_case2[0] = data_Cpe10_0_case2["Zone F"][i] + (data_Cpe10_0_case2["Zone F"][i+1] - data_Cpe10_0_case2["Zone F"][i]) * (alpha - data_Cpe10_0_case2["Pitch angle"][i]) / (data_Cpe10_0_case2["Pitch angle"][i+1] - data_Cpe10_0_case2["Pitch angle"][i])
            Cpe10_0_real_case2[1] = data_Cpe10_0_case2["Zone G"][i] + (data_Cpe10_0_case2["Zone G"][i+1] - data_Cpe10_0_case2["Zone G"][i]) * (alpha - data_Cpe10_0_case2["Pitch angle"][i]) / (data_Cpe10_0_case2["Pitch angle"][i+1] - data_Cpe10_0_case2["Pitch angle"][i])
            Cpe10_0_real_case2[2] = data_Cpe10_0_case2["Zone H"][i] + (data_Cpe10_0_case2["Zone H"][i+1] - data_Cpe10_0_case2["Zone H"][i]) * (alpha - data_Cpe10_0_case2["Pitch angle"][i]) / (data_Cpe10_0_case2["Pitch angle"][i+1] - data_Cpe10_0_case2["Pitch angle"][i])
            Cpe10_0_real_case2[3] = data_Cpe10_0_case2["Zone I"][i]+ (data_Cpe10_0_case2["Zone I"][i+1] - data_Cpe10_0_case2["Zone I"][i]) * (alpha - data_Cpe10_0_case2["Pitch angle"][i]) / (data_Cpe10_0_case2["Pitch angle"][i+1] - data_Cpe10_0_case2["Pitch angle"][i])
            Cpe10_0_real_case2[4] = data_Cpe10_0_case2["Zone J"][i]+ (data_Cpe10_0_case2["Zone J"][i+1] - data_Cpe10_0_case2["Zone J"][i]) * (alpha - data_Cpe10_0_case2["Pitch angle"][i]) / (data_Cpe10_0_case2["Pitch angle"][i+1] - data_Cpe10_0_case2["Pitch angle"][i])

            Cpe10_0_real_case3[0] = data_Cpe10_0_case3["Zone F"][i] + (data_Cpe10_0_case3["Zone F"][i+1] - data_Cpe10_0_case3["Zone F"][i]) * (alpha - data_Cpe10_0_case3["Pitch angle"][i]) / (data_Cpe10_0_case3["Pitch angle"][i+1] - data_Cpe10_0_case3["Pitch angle"][i])
            Cpe10_0_real_case3[1] = data_Cpe10_0_case3["Zone G"][i] + (data_Cpe10_0_case3["Zone G"][i+1] - data_Cpe10_0_case3["Zone G"][i]) * (alpha - data_Cpe10_0_case3["Pitch angle"][i]) / (data_Cpe10_0_case3["Pitch angle"][i+1] - data_Cpe10_0_case3["Pitch angle"][i])
            Cpe10_0_real_case3[2] = data_Cpe10_0_case3["Zone H"][i] + (data_Cpe10_0_case3["Zone H"][i+1] - data_Cpe10_0_case3["Zone H"][i]) * (alpha - data_Cpe10_0_case3["Pitch angle"][i]) / (data_Cpe10_0_case3["Pitch angle"][i+1] - data_Cpe10_0_case3["Pitch angle"][i])
            Cpe10_0_real_case3[3] = data_Cpe10_0_case3["Zone I"][i]+ (data_Cpe10_0_case3["Zone I"][i+1] - data_Cpe10_0_case3["Zone I"][i]) * (alpha - data_Cpe10_0_case3["Pitch angle"][i]) / (data_Cpe10_0_case3["Pitch angle"][i+1] - data_Cpe10_0_case3["Pitch angle"][i])
            Cpe10_0_real_case3[4] = data_Cpe10_0_case3["Zone J"][i]+ (data_Cpe10_0_case3["Zone J"][i+1] - data_Cpe10_0_case3["Zone J"][i]) * (alpha - data_Cpe10_0_case3["Pitch angle"][i]) / (data_Cpe10_0_case3["Pitch angle"][i+1] - data_Cpe10_0_case3["Pitch angle"][i])

            Cpe10_0_real_case4[0] = data_Cpe10_0_case4["Zone F"][i] + (data_Cpe10_0_case4["Zone F"][i+1] - data_Cpe10_0_case4["Zone F"][i]) * (alpha - data_Cpe10_0_case4["Pitch angle"][i]) / (data_Cpe10_0_case4["Pitch angle"][i+1] - data_Cpe10_0_case4["Pitch angle"][i])
            Cpe10_0_real_case4[1] = data_Cpe10_0_case4["Zone G"][i] + (data_Cpe10_0_case4["Zone G"][i+1] - data_Cpe10_0_case4["Zone G"][i]) * (alpha - data_Cpe10_0_case4["Pitch angle"][i]) / (data_Cpe10_0_case4["Pitch angle"][i+1] - data_Cpe10_0_case4["Pitch angle"][i])
            Cpe10_0_real_case4[2] = data_Cpe10_0_case4["Zone H"][i] + (data_Cpe10_0_case4["Zone H"][i+1] - data_Cpe10_0_case4["Zone H"][i]) * (alpha - data_Cpe10_0_case4["Pitch angle"][i]) / (data_Cpe10_0_case4["Pitch angle"][i+1] - data_Cpe10_0_case4["Pitch angle"][i])
            Cpe10_0_real_case4[3] = data_Cpe10_0_case4["Zone I"][i]+ (data_Cpe10_0_case4["Zone I"][i+1] - data_Cpe10_0_case4["Zone I"][i]) * (alpha - data_Cpe10_0_case4["Pitch angle"][i]) / (data_Cpe10_0_case4["Pitch angle"][i+1] - data_Cpe10_0_case4["Pitch angle"][i])
            Cpe10_0_real_case4[4] = data_Cpe10_0_case4["Zone J"][i]+ (data_Cpe10_0_case4["Zone J"][i+1] - data_Cpe10_0_case4["Zone J"][i]) * (alpha - data_Cpe10_0_case4["Pitch angle"][i]) / (data_Cpe10_0_case4["Pitch angle"][i+1] - data_Cpe10_0_case4["Pitch angle"][i])
            break
    
    return Cpe10_0_real_case1, Cpe10_0_real_case2, Cpe10_0_real_case3, Cpe10_0_real_case4

def Duopitch_Cpe1_0(alpha):
    
    """
    Input:
        direction: 0 degree
        alpha: Pitch angle
    """

    data_Cpe1_0_case1 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-0.6, -2, -2.8, -2.5, -2.5, -2, -1.5, -0, 0.7, 0.8],
        "Zone G": [-0.6, -1.5, -2, -2, -2, -1.5, -1.5, -0, 0.7, 0.8],
        "Zone H": [-0.8, -0.8, -1.2, -1.2, -1.2, -0.3, -0.2, -0, 0.7, 0.8],
        "Zone I": [-0.7, -0.6, -0.5, 0.2, -0.6, -0.4, -0.4, -0.2, -0.2, -0.2],
        "Zone J": [-1.5, -1.4, -1.2, 0.2, 0.2, -1.5, -0.5, -0.3, -0.3, -0.3]
        }  
    
    data_Cpe1_0_case2 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-0.6, -2, -2.8, -2.5, -2.5, -2, -1.5, -0, 0.7, 0.8],
        "Zone G": [-0.6, -1.5, -2, -2, -2, -1.5, -1.5, -0, 0.7, 0.8],
        "Zone H": [-0.8, -0.8, -1.2, -1.2, -1.2, -0.3, -0.2, -0, 0.7, 0.8],
        "Zone I": [-0.7, -0.6, -0.5, -0.6, -0.6, 0, 0, 0, -0.2, -0.2],
        "Zone J": [-1.5, -1.4, -1.2, -0.6, -0.6, 0, 0, 0, -0.3, -0.3]
        }
    
    data_Cpe1_0_case3 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-0.6, -2, -2.8, -2.5, 0, 0.2, 0.7, 0.7, 0.7, 0.8],
        "Zone G": [-0.6, -1.5, -2, -2, 0, 0.2, 0.7, 0.7, 0.7, 0.8],
        "Zone H": [-0.8, -0.8, -1.2, -1.2, 0, 0.2, 0.4, 0.6, 0.7, 0.8],
        "Zone I": [-0.7, -0.6, -0.5, 0.2, -0.6, -0.4, -0.4, -0.2, -0.2, -0.2],
        "Zone J": [-1.5, -1.4, -1.2, 0.2, 0.2, -1.5, -0.5, -0.3, -0.3, -0.3]
        }
    data_Cpe1_0_case4 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-0.6, -2, -2.8, -2.5, 0, 0.2, 0.7, 0.7, 0.7, 0.8],
        "Zone G": [-0.6, -1.5, -2, -2, 0, 0.2, 0.7, 0.7, 0.7, 0.8],
        "Zone H": [-0.8, -0.8, -1.2, -1.2, 0, 0.2, 0.4, 0.6, 0.7, 0.8],
        "Zone I": [-0.7, -0.6, -0.5, -0.6, -0.6, 0, 0, 0, -0.2, -0.2],
        "Zone J": [-1.5, -1.4, -1.2, -0.6, -0.6, 0, 0, 0, -0.3, -0.3]
        }
    
    data_Cpe1_0_case1 = pd.DataFrame(data_Cpe1_0_case1)
    data_Cpe1_0_case2 = pd.DataFrame(data_Cpe1_0_case2)
    data_Cpe1_0_case3 = pd.DataFrame(data_Cpe1_0_case3)
    data_Cpe1_0_case4 = pd.DataFrame(data_Cpe1_0_case4)

    Cpe1_0_real_case1 = np.zeros(5) # Cpe1 for zone F, G, H, I, J
    Cpe1_0_real_case2 = np.zeros(5) # Cpe1 for zone F, G, H, I, J
    Cpe1_0_real_case3 = np.zeros(5) # Cpe1 for zone F, G, H, I, J
    Cpe1_0_real_case4 = np.zeros(5) # Cpe1 for zone F, G, H, I, J
    for i in range(0, len(data_Cpe1_0_case1["Pitch angle"])):
        if (alpha == data_Cpe1_0_case1["Pitch angle"][i]):
            Cpe1_0_real_case1[0] = data_Cpe1_0_case1["Zone F"][i]
            Cpe1_0_real_case1[1] = data_Cpe1_0_case1["Zone G"][i]
            Cpe1_0_real_case1[2] = data_Cpe1_0_case1["Zone H"][i]
            Cpe1_0_real_case1[3] = data_Cpe1_0_case1["Zone I"][i]
            Cpe1_0_real_case1[4] = data_Cpe1_0_case1["Zone J"][i]

            Cpe1_0_real_case2[0] = data_Cpe1_0_case2["Zone F"][i]
            Cpe1_0_real_case2[1] = data_Cpe1_0_case2["Zone G"][i]
            Cpe1_0_real_case2[2] = data_Cpe1_0_case2["Zone H"][i]
            Cpe1_0_real_case2[3] = data_Cpe1_0_case2["Zone I"][i]
            Cpe1_0_real_case2[4] = data_Cpe1_0_case2["Zone J"][i]

            Cpe1_0_real_case3[0] = data_Cpe1_0_case3["Zone F"][i]
            Cpe1_0_real_case3[1] = data_Cpe1_0_case3["Zone G"][i]
            Cpe1_0_real_case3[2] = data_Cpe1_0_case3["Zone H"][i]
            Cpe1_0_real_case3[3] = data_Cpe1_0_case3["Zone I"][i]
            Cpe1_0_real_case3[4] = data_Cpe1_0_case3["Zone J"][i]

            Cpe1_0_real_case4[0] = data_Cpe1_0_case4["Zone F"][i]
            Cpe1_0_real_case4[1] = data_Cpe1_0_case4["Zone G"][i]
            Cpe1_0_real_case4[2] = data_Cpe1_0_case4["Zone H"][i]
            Cpe1_0_real_case4[3] = data_Cpe1_0_case4["Zone I"][i]
            Cpe1_0_real_case4[4] = data_Cpe1_0_case4["Zone J"][i]
            break

        if (alpha > data_Cpe1_0_case1["Pitch angle"][i]) and (alpha < data_Cpe1_0_case1["Pitch angle"][i+1]):
            Cpe1_0_real_case1[0] = data_Cpe1_0_case1["Zone F"][i] + (data_Cpe1_0_case1["Zone F"][i+1] - data_Cpe1_0_case1["Zone F"][i]) * (alpha - data_Cpe1_0_case1["Pitch angle"][i]) / (data_Cpe1_0_case1["Pitch angle"][i+1] - data_Cpe1_0_case1["Pitch angle"][i])
            Cpe1_0_real_case1[1] = data_Cpe1_0_case1["Zone G"][i] + (data_Cpe1_0_case1["Zone G"][i+1] - data_Cpe1_0_case1["Zone G"][i]) * (alpha - data_Cpe1_0_case1["Pitch angle"][i]) / (data_Cpe1_0_case1["Pitch angle"][i+1] - data_Cpe1_0_case1["Pitch angle"][i])
            Cpe1_0_real_case1[2] = data_Cpe1_0_case1["Zone H"][i] + (data_Cpe1_0_case1["Zone H"][i+1] - data_Cpe1_0_case1["Zone H"][i]) * (alpha - data_Cpe1_0_case1["Pitch angle"][i]) / (data_Cpe1_0_case1["Pitch angle"][i+1] - data_Cpe1_0_case1["Pitch angle"][i])
            Cpe1_0_real_case1[3] = data_Cpe1_0_case1["Zone I"][i] + (data_Cpe1_0_case1["Zone I"][i+1] - data_Cpe1_0_case1["Zone I"][i]) * (alpha - data_Cpe1_0_case1["Pitch angle"][i]) / (data_Cpe1_0_case1["Pitch angle"][i+1] - data_Cpe1_0_case1["Pitch angle"][i])
            Cpe1_0_real_case1[4] = data_Cpe1_0_case1["Zone J"][i] + (data_Cpe1_0_case1["Zone J"][i+1] - data_Cpe1_0_case1["Zone J"][i]) * (alpha - data_Cpe1_0_case1["Pitch angle"][i]) / (data_Cpe1_0_case1["Pitch angle"][i+1] - data_Cpe1_0_case1["Pitch angle"][i])

            Cpe1_0_real_case2[0] = data_Cpe1_0_case2["Zone F"][i] + (data_Cpe1_0_case2["Zone F"][i+1] - data_Cpe1_0_case2["Zone F"][i]) * (alpha - data_Cpe1_0_case2["Pitch angle"][i]) / (data_Cpe1_0_case2["Pitch angle"][i+1] - data_Cpe1_0_case2["Pitch angle"][i])
            Cpe1_0_real_case2[1] = data_Cpe1_0_case2["Zone G"][i] + (data_Cpe1_0_case2["Zone G"][i+1] - data_Cpe1_0_case2["Zone G"][i]) * (alpha - data_Cpe1_0_case2["Pitch angle"][i]) / (data_Cpe1_0_case2["Pitch angle"][i+1] - data_Cpe1_0_case2["Pitch angle"][i])
            Cpe1_0_real_case2[2] = data_Cpe1_0_case2["Zone H"][i] + (data_Cpe1_0_case2["Zone H"][i+1] - data_Cpe1_0_case2["Zone H"][i]) * (alpha - data_Cpe1_0_case2["Pitch angle"][i]) / (data_Cpe1_0_case2["Pitch angle"][i+1] - data_Cpe1_0_case2["Pitch angle"][i])
            Cpe1_0_real_case2[3] = data_Cpe1_0_case2["Zone I"][i] + (data_Cpe1_0_case2["Zone I"][i+1] - data_Cpe1_0_case2["Zone I"][i]) * (alpha - data_Cpe1_0_case2["Pitch angle"][i]) / (data_Cpe1_0_case2["Pitch angle"][i+1] - data_Cpe1_0_case2["Pitch angle"][i])
            Cpe1_0_real_case2[4] = data_Cpe1_0_case2["Zone J"][i] + (data_Cpe1_0_case2["Zone J"][i+1] - data_Cpe1_0_case2["Zone J"][i]) * (alpha - data_Cpe1_0_case2["Pitch angle"][i]) / (data_Cpe1_0_case2["Pitch angle"][i+1] - data_Cpe1_0_case2["Pitch angle"][i])

            Cpe1_0_real_case3[0] = data_Cpe1_0_case3["Zone F"][i] + (data_Cpe1_0_case3["Zone F"][i+1] - data_Cpe1_0_case3["Zone F"][i]) * (alpha - data_Cpe1_0_case3["Pitch angle"][i]) / (data_Cpe1_0_case3["Pitch angle"][i+1] - data_Cpe1_0_case3["Pitch angle"][i])
            Cpe1_0_real_case3[1] = data_Cpe1_0_case3["Zone G"][i] + (data_Cpe1_0_case3["Zone G"][i+1] - data_Cpe1_0_case3["Zone G"][i]) * (alpha - data_Cpe1_0_case3["Pitch angle"][i]) / (data_Cpe1_0_case3["Pitch angle"][i+1] - data_Cpe1_0_case3["Pitch angle"][i])
            Cpe1_0_real_case3[2] = data_Cpe1_0_case3["Zone H"][i] + (data_Cpe1_0_case3["Zone H"][i+1] - data_Cpe1_0_case3["Zone H"][i]) * (alpha - data_Cpe1_0_case3["Pitch angle"][i]) / (data_Cpe1_0_case3["Pitch angle"][i+1] - data_Cpe1_0_case3["Pitch angle"][i])
            Cpe1_0_real_case3[3] = data_Cpe1_0_case3["Zone I"][i] + (data_Cpe1_0_case3["Zone I"][i+1] - data_Cpe1_0_case3["Zone I"][i]) * (alpha - data_Cpe1_0_case3["Pitch angle"][i]) / (data_Cpe1_0_case3["Pitch angle"][i+1] - data_Cpe1_0_case3["Pitch angle"][i])
            Cpe1_0_real_case3[4] = data_Cpe1_0_case3["Zone J"][i] + (data_Cpe1_0_case3["Zone J"][i+1] - data_Cpe1_0_case3["Zone J"][i]) * (alpha - data_Cpe1_0_case3["Pitch angle"][i]) / (data_Cpe1_0_case3["Pitch angle"][i+1] - data_Cpe1_0_case3["Pitch angle"][i])

            Cpe1_0_real_case4[0] = data_Cpe1_0_case4["Zone F"][i] + (data_Cpe1_0_case4["Zone F"][i+1] - data_Cpe1_0_case4["Zone F"][i]) * (alpha - data_Cpe1_0_case4["Pitch angle"][i]) / (data_Cpe1_0_case4["Pitch angle"][i+1] - data_Cpe1_0_case4["Pitch angle"][i])
            Cpe1_0_real_case4[1] = data_Cpe1_0_case4["Zone G"][i] + (data_Cpe1_0_case4["Zone G"][i+1] - data_Cpe1_0_case4["Zone G"][i]) * (alpha - data_Cpe1_0_case4["Pitch angle"][i]) / (data_Cpe1_0_case4["Pitch angle"][i+1] - data_Cpe1_0_case4["Pitch angle"][i])
            Cpe1_0_real_case4[2] = data_Cpe1_0_case4["Zone H"][i] + (data_Cpe1_0_case4["Zone H"][i+1] - data_Cpe1_0_case4["Zone H"][i]) * (alpha - data_Cpe1_0_case4["Pitch angle"][i]) / (data_Cpe1_0_case4["Pitch angle"][i+1] - data_Cpe1_0_case4["Pitch angle"][i])
            Cpe1_0_real_case4[3] = data_Cpe1_0_case4["Zone I"][i] + (data_Cpe1_0_case4["Zone I"][i+1] - data_Cpe1_0_case4["Zone I"][i]) * (alpha - data_Cpe1_0_case4["Pitch angle"][i]) / (data_Cpe1_0_case4["Pitch angle"][i+1] - data_Cpe1_0_case4["Pitch angle"][i])
            Cpe1_0_real_case4[4] = data_Cpe1_0_case4["Zone J"][i] + (data_Cpe1_0_case4["Zone J"][i+1] - data_Cpe1_0_case4["Zone J"][i]) * (alpha - data_Cpe1_0_case4["Pitch angle"][i]) / (data_Cpe1_0_case4["Pitch angle"][i+1] - data_Cpe1_0_case4["Pitch angle"][i])

            
            break
    

    return Cpe1_0_real_case1, Cpe1_0_real_case2, Cpe1_0_real_case3, Cpe1_0_real_case4

def Duopitch_Cpe_90(alpha):
    
    """
    Input:
        direction:90 degree
        alpha: Pitch angle
    """
    data_Cpe10_90 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-1.4, -1.5, -1.9, -1.8, -1.6, -1.3, -1.1, -1.1, -1.1, -1.1],
        "Zone G": [-1.2, -1.2, -1.2, -1.2, -1.3, -1.3, -1.4, -1.4, -1.2, -1.2],
        "Zone H": [-1, -1, -0.8, -0.7, -0.7, -0.6, -0.8, -0.9, -0.8, -0.8],
        "Zone I": [-0.9, -0.9, -0.8, -0.6, -0.6, -0.5, -0.5, -0.5, -0.5, -0.5]
        }
    
    data_Cpe10_90 = pd.DataFrame(data_Cpe10_90)

    data_Cpe1_90 = {
        "Pitch angle": [-45, -30, -15, -5, 5, 15, 30, 45, 60, 75],
        "Zone F": [-2, -2.1, -2.5, -2.5, -2.2, -2, -1.5, -1.5, -1.5, -1.5],
        "Zone G": [-2, -2, -2, -2, -2, -2, -2, -2, -2, -2],
        "Zone H": [-1.3, -1.3, -1.2, -1.2, -1.2, -1.2, -1.2, -1.2, -1.0, -1.0],
        "Zone I": [-1.2, -1.2, -1.2, -1.2, -0.6, -0.5, -0.5, -0.5, -0.5, -0.5]
        }
    
    data_Cpe1_90 = pd.DataFrame(data_Cpe1_90)

    
   
    # Check which angle in "Pitch angle" column, the alpha is in between, and then use the two angles to interpolate the Cpe values for each zone
    Cpe10_90_real = np.zeros(4) # Cpe10 for zone F, G, H, I
    Cpe1_90_real = np.zeros(4) # Cpe1 for zone F, G, H, I

    for i in range(0, len(data_Cpe10_90["Pitch angle"])):
        if (alpha == data_Cpe10_90["Pitch angle"][i]):
            Cpe10_90_real[0] = data_Cpe10_90["Zone F"][i]
            Cpe10_90_real[1] = data_Cpe10_90["Zone G"][i]
            Cpe10_90_real[2] = data_Cpe10_90["Zone H"][i]
            Cpe10_90_real[3] = data_Cpe10_90["Zone I"][i]
            break
        if (alpha > data_Cpe10_90["Pitch angle"][i]) and (alpha < data_Cpe10_90["Pitch angle"][i+1]):
            Cpe10_90_real[0] = data_Cpe10_90["Zone F"][i] + (data_Cpe10_90["Zone F"][i+1] - data_Cpe10_90["Zone F"][i]) * (alpha - data_Cpe10_90["Pitch angle"][i]) / (data_Cpe10_90["Pitch angle"][i+1] - data_Cpe10_90["Pitch angle"][i])
            Cpe10_90_real[1] = data_Cpe10_90["Zone G"][i] + (data_Cpe10_90["Zone G"][i+1] - data_Cpe10_90["Zone G"][i]) * (alpha - data_Cpe10_90["Pitch angle"][i]) / (data_Cpe10_90["Pitch angle"][i+1] - data_Cpe10_90["Pitch angle"][i])
            Cpe10_90_real[2] = data_Cpe10_90["Zone H"][i] + (data_Cpe10_90["Zone H"][i+1] - data_Cpe10_90["Zone H"][i]) * (alpha - data_Cpe10_90["Pitch angle"][i]) / (data_Cpe10_90["Pitch angle"][i+1] - data_Cpe10_90["Pitch angle"][i])
            Cpe10_90_real[3] = data_Cpe10_90["Zone I"][i]+ (data_Cpe10_90["Zone I"][i+1] - data_Cpe10_90["Zone I"][i]) * (alpha - data_Cpe10_90["Pitch angle"][i]) / (data_Cpe10_90["Pitch angle"][i+1] - data_Cpe10_90["Pitch angle"][i])
            break
    
    for i in range(0, len(data_Cpe1_90["Pitch angle"])):
        if (alpha == data_Cpe1_90["Pitch angle"][i]):
            Cpe1_90_real[0] = data_Cpe1_90["Zone F"][i]
            Cpe1_90_real[1] = data_Cpe1_90["Zone G"][i]
            Cpe1_90_real[2] = data_Cpe1_90["Zone H"][i]
            Cpe1_90_real[3] = data_Cpe1_90["Zone I"][i]
            break
        if (alpha > data_Cpe1_90["Pitch angle"][i]) and (alpha < data_Cpe1_90["Pitch angle"][i+1]):
            Cpe1_90_real[0] = data_Cpe1_90["Zone F"][i] + (data_Cpe1_90["Zone F"][i+1] - data_Cpe1_90["Zone F"][i]) * (alpha - data_Cpe1_90["Pitch angle"][i]) / (data_Cpe1_90["Pitch angle"][i+1] - data_Cpe1_90["Pitch angle"][i])
            Cpe1_90_real[1] = data_Cpe1_90["Zone G"][i] + (data_Cpe1_90["Zone G"][i+1] - data_Cpe1_90["Zone G"][i]) * (alpha - data_Cpe1_90["Pitch angle"][i]) / (data_Cpe1_90["Pitch angle"][i+1] - data_Cpe1_90["Pitch angle"][i])
            Cpe1_90_real[2] = data_Cpe1_90["Zone H"][i] + (data_Cpe1_90["Zone H"][i+1] - data_Cpe1_90["Zone H"][i]) * (alpha - data_Cpe1_90["Pitch angle"][i]) / (data_Cpe1_90["Pitch angle"][i+1] - data_Cpe1_90["Pitch angle"][i])
            Cpe1_90_real[3] = data_Cpe1_90["Zone I"][i]+ (data_Cpe1_90["Zone I"][i+1] - data_Cpe1_90["Zone I"][i]) * (alpha - data_Cpe1_90["Pitch angle"][i]) / (data_Cpe1_90["Pitch angle"][i+1] - data_Cpe1_90["Pitch angle"][i])
            break

    return Cpe10_90_real, Cpe1_90_real

def Cal_Rougnesslength(TerrainCategory):
    TerrainParameter = {
        "Terrain_category": ["0", "I", "II", "III", "IV"],
        "z0": [0.003, 0.01, 0.05, 0.3, 1.0],
        "zmin": [1,1,2,5,10]
        }
    TerrainParameter = pd.DataFrame(TerrainParameter)
    for i in range(0, len(TerrainParameter["Terrain_category"])):
            if (TerrainCategory == TerrainParameter["Terrain_category"][i]):
                    z0 = TerrainParameter["z0"][i]
                    zmin = TerrainParameter["zmin"][i]
                    break
    
    return z0, zmin