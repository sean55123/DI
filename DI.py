import win32com.client as win32
import os
import time
import numpy as np

## Key in the parameters that independent to blocks
chemicals = ['PO', 'PC', 'PG', 'GLY', 'GC'] # The sequece of properties below (CO2 is safe enough to be excluded)
## Data Source: Internet
NF = np.array([4, 1, 1, 1, 1]) # NFPA red
NH = np.array([3, 1, 0, 1, 2]) # NFPA blue
NR = np.array([2, 1, 0, 0, 0]) # NFPA yellow
FP = np.array([-37, 128, 101, 176, 149.5]) # flash points for each chemical in C
AI = 0.75*np.array([747, 435, 420, 393, 404]) # 0.75*auto-ignition points in C
HC = np.array([-1917.4, -1818, -1838.19, -1654.3, -1594.2015]) # Heat of combustion in kJ/mol

## Data Source: Aspen Plus V12.1
para = np.array([
            ["A", 79.52407454, -5976.3,  0,  0, -10.686, 0.000011993, 2, 0], # PO PLXANT-1 in C & bar 
            ["A", 94.92707454, -10819,   0,  0, -12.068, 5.46E-06, 2, 0], # PC PLXANT-1 in C & bar 
            ["A", 123.9770745, -12483,   0,  0, -16.019, 6.46E-06, 2, 0], # PG PLXANT-1 in C & bar 
            ["A", 88.47307454, -13808,   0,  0, -10.088, 3.57E-19, 6, 0], # GLY PLXANT-1 in C & bar 
            ["P", 21.26683454, -10586.9, 0.0394255,  -0.0065509, 0, 0, 0, 0]]) # GC PLTDEP-1 in C & bar 
# Noted that there are only 7 para. in PLXANT, but have 8 in PLTDEP. Thus, add 0 to those PLAXANT para. sets to make the array symmetric.
# "A" represents Extended Antonie Equation, and "P" represent NIST TDE Polynomial for Liquid Vapor Pressure.
# If there is another equation used in your work, go to calculate_vp to write down the formula.

## The reactants for each reaction (for the determination of F4 calculation)
chem_rxn = {
    1: ['PO'],
    2: ['PC', 'GLY'],
    3: ['PG', 'GC']
} 

## Ambient pressure in kPa (default to be 1 atm)
AP = 101.3 

## Parameters defined by Rangaiah_2013
pn4 = np.maximum(np.ones(5), 0.3 * (NF + NR))
pn7 = pn8 = 1.1
# =============================================================================
# pn7 = 2, 1.5 and 1.1 for disasters occurring once in a year, 5 years and 20 years, respectively.
# pn8 = 2 for an area that is prone to accidents, and 1.1 otherwise.
# =============================================================================


# Open Aspen File (give the filename here, notice that the Aspen file should be in the same folder as this .py file)
aspen = win32.Dispatch('Apwn.Document')     
filepath = os.path.join(os.path.abspath('.'),'Fuzzy 2R_HX(RUN9).bkp')
# aspen.InitFromFile2(filepath)
aspen = win32.GetObject(filepath).Application
aspen.Visible = 1
aspen.SuppressDialogs = 1

time.sleep(5)

# Finding the target chemicals, stream for each block and print it out
def calculate_block_performance(instream, outstream, bname, rxn, pn9, pn10, aspen):
    iname_list = []
    oname_list = []
    MFin = np.zeros((instream, 5))
    MFout = np.zeros((outstream, 5))

    for i in range(instream):
        iname = aspen.Application.Tree.FindNode(f"\\Data\\Flowsheet\\Section\\GLOBAL\\Input\\INSTRM\\{bname}\\#{i}").value
        iname_list.append(iname)
        for j, property_name in enumerate(['PO', 'PC', 'PG', 'GLY', 'GC']):
            property_value = aspen.Application.Tree.FindNode(f'\\Data\\Streams\\{iname}\\Output\\MASSFLOW\\MIXED\\{property_name}').ValueForUnit(10, 1)
            MFin[i][j] = property_value

    for i in range(outstream):
        oname = aspen.Application.Tree.FindNode(f"\\Data\\Flowsheet\\Section\\GLOBAL\\Input\\OUTSTRM\\{bname}\\#{i}").value
        oname_list.append(oname)
        for j, property_name in enumerate(['PO', 'PC', 'PG', 'GLY', 'GC']):
            property_value = aspen.Application.Tree.FindNode(f'\\Data\\Streams\\{oname}\\Output\\MASSFLOW\\MIXED\\{property_name}').ValueForUnit(10, 1)
            MFout[i][j] = property_value

    combined_matrix = np.vstack((MFin, MFout))
    pn4_MF = pn4 * combined_matrix
    max_value = np.max(pn4_MF)
    max_positions = np.argwhere(pn4_MF == max_value)
    streams = iname_list + oname_list

    max_values = []
    for max_row, max_col in max_positions:
        chemical_name = chemicals[max_col]
        stream_name = streams[max_row]
        max_values.append((max_value, chemical_name, stream_name))
    return max_values

def calculate_and_print_block_performance(instream, outstream, bname, rxn, pn9, pn10):
    block_max_values = calculate_block_performance(instream, outstream, bname, rxn, pn9, pn10, aspen)
    print(f"{bname} Target Prediction:")
    for value, chemical_name, stream_name in block_max_values:
        print(f"Value: {value:.3f}, Target Chemical: {chemical_name}, Target Stream: {stream_name}")

# Calculation of FEDR and transfer to FEDI
def calculate_vp(T, index_TC, PP, AP, para, target_chemical, target_stream, aspen):
    LT = np.array([-111.93, 48.3, -60, 18.18, 10]) # lower temperature of extended Antoine eq. from Aspen Plus V12.1 (unit: C)
    UT = np.array([209.1, 504.85, 403.25, 576.85, 182.46]) # upper temperature of extended Antoine eq. from Aspen Plus V12.1 (unit: C)
    if  LT[index_TC] <= T <= UT[index_TC] and para[index_TC][0] == np.str_('A'):
        VP = np.exp(float(para[index_TC][1]) + float(para[index_TC][2]) / ((T +273.15)\
              + float(para[index_TC][3])) + float(para[index_TC][4]) * (T+273.15)\
              + float(para[index_TC][5]) * np.log(T+273.15)\
              + float(para[index_TC][6]) * ((T+273.15) ** float(para[index_TC][7]))) # bar
        VP = 100*VP # bar to kPa
    elif  LT[index_TC] <= T <= UT[index_TC] and para[index_TC][0] == np.str_('P'):
        VP = np.exp(float(para[index_TC][1]) + float(para[index_TC][2])/ (T + 273.15)\
                    + float(para[index_TC][3])*np.log(T + 273.15) + float(para[index_TC][4])*(T + 273.15)\
                    + float(para[index_TC][5])* (T + 273.15)**2 + float(para[index_TC][6])/(T + 273.15)**2\
                    + float(para[index_TC][7])*(T + 273.15)**6 + float(para[index_TC][8])/(T +273.15)**4) # bar
        VP = 100*VP # bar to kPa
    else:
        VP = PP * aspen.Application.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\MASSFRAC\\MIXED\\{target_chemical}').Value
    return VP

def rxn_or_not(rxn, target_chemical, F1, pn1, F, pn2, F4, pn9, pn10, pn3, pn4, index_TC, pn7, pn8):
    base_value = 4.76 * ((F1 * pn1 + F * pn2) * pn3 * pn4[index_TC] * pn7 * pn8) ** (1/3)

    for i, reaction_chemicals in chem_rxn.items():
        # if rxn == i and target_chemical in reaction_chemicals:
        if rxn == 1:
            return 4.76 * ((F1 * pn1 + F * pn2 + F4 * pn9 * pn10) * pn3 * pn4[index_TC] * pn7 * pn8) ** (1/3)

    return base_value

def calculate_FEDI(instream, outstream, bname, rxn, pn9, pn10, aspen):
    block_performance_results = calculate_block_performance(instream, outstream, bname, rxn, pn9, pn10, aspen)
    
    max_FEDI = -1  
    stream_lst = []
    FEDR_lst = []
    FEDI_lst = []
    result_lst = []

    for result in block_performance_results:
        target_chemical = result[1]
        target_stream = result[2]
        index_TC = chemicals.index(target_chemical)
        stream_lst.append(target_stream)
        
        MF = aspen.Application.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\MASSFLOW\\MIXED\\{target_chemical}').ValueForUnit(10, 1) # kg/s
        T = aspen.Application.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\RES_TEMP').ValueForUnit(22, 4) # C
        PP = aspen.Application.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\RES_PRES').ValueForUnit(20, 10) # kPa
        VF_tot = aspen.Application.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\VOLFLMX\\MIXED').ValueForUnit(12, 7) # cum/hr
        MFrac = aspen.Application.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\MASSFRAC\\MIXED\\{target_chemical}').Value # dimensionless
        
        
        VP = calculate_vp(T, index_TC, PP, AP, para, target_chemical, target_stream, aspen)
        VF = VF_tot * MFrac
        F1 = 0.1 * MF * -HC[index_TC] / 3.148
        F2 = 1.304 * 10 ** (-3) * PP * VF
        F3 = ((1.0 * 10 ** (-3)) / (T + 273)) * (PP - VP) ** 2 * VF

        # Conditional statement according to the definition from ref.Rangaiah_2013
        if VP > AP and PP > VP:
            F = F2 + F3
        elif VP > AP and PP < VP:
            F = F2
        else:
            F = F3

        # if rxn == 1 and target_chemical in chem_rxn[1]:
        #     F4 = MF * HRXN[rxn-1] / 3.148
        # else:
        #     F4 = 0
        if rxn == 1:
            Q = aspen.Application.Tree.FindNode(f'\\Data\\Blocks\\{bname}\\Output\\QCALC').ValueForUnit(13, 14) # kW = kJ/s
            if Q < 0:
                F4 = -Q / 3.148 
            else:
                F4 = 0
        else:
            F4 = 0

        if T < FP[index_TC]:
            pn1 = 1.1
        elif FP[index_TC] < T < AI[index_TC]:
            pn1 = 1.75
        else:
            pn1 = 1.95

        if VP > AP and PP > VP:
            pn2 = 1 + abs(0.6 * ((PP - VP) / PP))
        elif VP > AP and PP < VP:
            pn2 = 1 + abs(0.4 * ((PP - VP) / PP))
        elif VP < AP and PP > AP:
            pn2 = 1 + abs(0.2 * ((PP - VP) / PP))
        else:
            pn2 = 1.1

        index_pn3 = max(NF[index_TC], NR[index_TC])
        CI = aspen.Application.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\MASSFLOW\\MIXED\\{target_chemical}').ValueForUnit(10, 12)  # Chemical Inventory in tons/hr

        if index_pn3 == 4:
            pn3 = 0.0102 * CI + 0.9936
        elif index_pn3 == 3:
            pn3 = 0.0076 * CI + 0.9956
        elif index_pn3 == 2:
            pn3 = 0.0051 * CI + 0.9935
        elif index_pn3 == 0 or CI < 10:
            pn3 = 1
        else:
            pn3 = 0.0026 * CI + 0.992

        FEDR = rxn_or_not(rxn, target_chemical, F1, pn1, F, pn2, F4, pn9, pn10, pn3, pn4, index_TC, pn7, pn8)
        FEDR_lst.append(FEDR)

        if 10 <= FEDR <= 200:
            FEDI = FEDR / 2
        elif 200 < FEDR <= 300:
            FEDI = 100
        else:
            FEDI = 0
        FEDI_lst.append(FEDI)
    # print(f"VP: {VP:.4f}")
    # print(f"F1: {F1:.4f}")
    # print(f"F2: {F2:.4f}")
    # print(f"F3: {F3:.4f}")
    # print(f"F4: {F4:.4f}")
    # print(f"F: {F:.4f}")
    # print(f"pn1: {pn1:.4f}")
    # print(f"pn2: {pn2:.4f}")
    # print(f"pn3: {pn3:.4f}")
    
    
    for stream, FEDR, FEDI in zip(stream_lst, FEDR_lst, FEDI_lst):
        result_lst.append((stream, FEDR, FEDI))
    max_FEDI = max(result_lst, key=lambda x: x[1])
    print(f"{bname}_FEDI: {max_FEDI[2]:.4f}")
    return max(FEDI_lst)

def process_block(instream, outstream, bname, rxn, pn9, pn10, aspen):
    # calculate_and_print_block_performance(instream, outstream, bname, rxn, pn9, pn10) 
    calculate_block_performance(instream, outstream, bname, rxn, pn9, pn10, aspen)
    # calculate_and_print_block_performance(instream, outstream, bname, rxn, pn9, pn10)
    block_FEDI = calculate_FEDI(instream, outstream, bname, rxn, pn9, pn10, aspen)  
    return block_FEDI

# =============================================================================
# instream: the number of input stream of that unit
# outstream: the number of output stream of that unit
# bname: the specify block name in Aspen plus file
# rxn:
#     no reaction 0; with reaction 1
# 
# pn9:
#     oxidation 1.60; esterfication 1.25; nitration 1.95; reduction 1.10; halogenation 1.45; pyrolysis 1.45;
#     alkylation 1.25; hydrogenation 1.35; electrolysis 1.20; sulfonation 1.30; polymerization 1.50; aminolysis 1.40; others 1
# 
# pn10:
#     autocatalytic reaction 1.65
#     non-autocatalytic reaction occurring at above normal reaction conditions 1.45
#     non-autocatalytic reaction occurring at below normal reaction conditions 1.45
# =============================================================================

## Key in the block information for the process
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
    
total_FEDI = 0

for block in blocks:
    block_FEDI = process_block(block["instream"], block["outstream"], block["bname"], block["rxn"], block["pn9"], block["pn10"], aspen)
    total_FEDI += block_FEDI # add FEDI of every blocks together

print(f'total_FEDI: {total_FEDI:.4f}')

aspen.close()
aspen.quit()