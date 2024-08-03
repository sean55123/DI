import numpy as np
from DI import DI, DI_calculator

chemicals = ['PO', 'PC', 'PG', 'GLY', 'GC'] 
NF = np.array([4, 1, 1, 1, 1])
NH = np.array([3, 1, 0, 1, 2])
NR = np.array([2, 1, 0, 0, 0])
FP = np.array([-37, 128, 101, 176, 149.5]) 
AI = 0.75*np.array([747, 435, 420, 393, 404]) 
HC = np.array([-1917.4, -1818, -1838.19, -1654.3, -1594.2015])
 
para = np.array([
            ["A", 79.52407454, -5976.3,  0,  0, -10.686, 0.000011993, 2, 0], # PO PLXANT-1 in C & bar 
            ["A", 94.92707454, -10819,   0,  0, -12.068, 5.46E-06, 2, 0], # PC PLXANT-1 in C & bar 
            ["A", 123.9770745, -12483,   0,  0, -16.019, 6.46E-06, 2, 0], # PG PLXANT-1 in C & bar 
            ["A", 88.47307454, -13808,   0,  0, -10.088, 3.57E-19, 6, 0], # GLY PLXANT-1 in C & bar 
            ["P", 21.26683454, -10586.9, 0.0394255,  -0.0065509, 0, 0, 0, 0]]) # GC PLTDEP-1 in C & bar 

chem_rxn = {
    1: ['PO'],
    2: ['PC', 'GLY'],
    3: ['PG', 'GC']
} 

aspenfile_name = "Fuzzy 2R_HX(RUN9).bkp"

pn4 = np.maximum(np.ones(5), 0.3 * (NF + NR))
pn7, pn8 = 1.1, 1.1
LT = np.array([-111.93, 48.3, -60, 18.18, 10])
UT = np.array([209.1, 504.85, 403.25, 576.85, 182.46]) 

blocks = [
    {"instream": 3, "outstream": 2, "bname": "R1", "rxn": 1, "pn9": 1, "pn10": 1.45},
    {"instream": 2, "outstream": 2, "bname": "R2", "rxn": 1, "pn9": 1, "pn10": 1.45},
    {"instream": 1, "outstream": 1, "bname": "COOL", "rxn": 0, "pn9": 1, "pn10": 1.45},
    {"instream": 1, "outstream": 1, "bname": "V1", "rxn": 0, "pn9": 1, "pn10": 1.45},
    {"instream": 2, "outstream": 2, "bname": "HX", "rxn": 0, "pn9": 1, "pn10": 1.45},
    {"instream": 2, "outstream": 3, "bname": "C1", "rxn": 0, "pn9": 1, "pn10": 1.45},
    {"instream": 1, "outstream": 2, "bname": "C2", "rxn": 0, "pn9": 1, "pn10": 1.45},
    {"instream": 1, "outstream": 1, "bname": "P1", "rxn": 0, "pn9": 1, "pn10": 1.45},
    {"instream": 1, "outstream": 1, "bname": "P2", "rxn": 0, "pn9": 1, "pn10": 1.45},
]

model = DI(chemicals, NF, NH, NR, FP, AI, HC, para, chem_rxn, pn7, pn8, LT, UT, blocks, aspenfile_name)
DI_calculator()