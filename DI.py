import win32com.client as win32
import os
import time
import numpy as np

class DI:
    def __init__(self, chemicals, NF, NH, NR, FP, AI, HC, para, chem_rxn, pn7, pn8, LT, UT, blocks, aspenfile_name, AP=101.3):
        """Set up the necessary parameters for calculation.
        
        Noted for the para, there are only 7 parameters in PLXANT, while 8 in PLTDEP. 
        As a result, it would be necessary to add 0 in the PLXANT parameters list to keeps the array symmetric.
        Besides, "A" represents Extended Antonie Equation, and "P" represent NIST TDE Polynomial for Liquid Vapor Pressure.
        If there is another equation used in your work, adjust calculate_vp to keep it align with the formula in your design.
        
        Args:
            chemicals (list): The sequece of properties below (CO2 is safe enough to be excluded from calculation)
            NF (list): NFPA red
            NH (list): NFPA blue
            NR (list): NFPA yellow
            FP (list): Flash points for each chemical (oC)
            AI (list): Auto-ignition point, and 0.75 used for safety index (oC)
            HC (list): Heat of combustion (kJ/mol)
            AP (flaot): Ambient pressure (kpa, default=101.3 kpa)
            para (list): PLXANT PLTDEP data for specific chemical
            chem_rxn (tuple): Reactions
            pn7 (float): 2, 1.5 and 1.1 for disasters occurring once in a year, 5 years and 20 years, respectively.
            pn8 (float): 2 for an area that is prone to accidents, and 1.1 otherwise.
            LT (list): Lower temperature of extended Antoine eq. from Aspen Plus V12.1 (oC)
            UT (list): Upper temperature of extended Antoine eq. from Aspen Plus V12.1 (oC)
        """
        self.chemicals = chemicals
        self.NF = np.array(NF)
        self.NH = np.array(NH)
        self.NR = np.array(NR)
        self.FP = np.array(FP)
        self.AI = 0.75*np.array(AI)
        self.HC = np.array(HC)
        self.AP = np.array(AP)
        self.para = np.array(para)
        self.chem_rxn = chem_rxn
        self.pn4 = np.maximum(np.ones(5), 0.3 * (NF + NR))
        self.pn7 = pn7
        self.pn8 = pn8
        self.LT = np.array(LT)
        self.UT = np.array(UT)
        self.blocks = blocks
        self.aspenfile_name = aspenfile_name
        self.DI_calculator()
        
    def Aspen(self):
        self.aspen = win32.Dispatch('Apwn.Document')     
        filepath = os.path.join(os.path.abspath('.'), self.aspenfile_name)
        self.aspen = win32.GetObject(filepath)
        self.aspen.Visible = 1
        self.aspen.SuppressDialogs = 1

        time.sleep(1)
    
    def calculate_block_performance(self, instream, outstream, bname, rxn, pn9, pn10):
        """Here the specific chemical and stream with highest pn4 massflow would be targeted for FEDI calculation.

        Args:
            instream (str): The number of input stream of that unit
            outstream (str): The number of output stream of that unit
            bname (str): The specify block name in Aspen plus file
            rxn (int): No reaction label with 0; with reaction label with 1
            pn9 (float): Oxidation 1.60; esterfication 1.25; nitration 1.95; reduction 1.10; halogenation 1.45; pyrolysis 1.45; 
                         alkylation 1.25; hydrogenation 1.35; electrolysis 1.20; sulfonation 1.30; polymerization 1.50; aminolysis 1.40; others 1
            pn10 (float): Autocatalytic reaction 1.65
                          Non-autocatalytic reaction occurring at above normal reaction conditions 1.45
                          Non-autocatalytic reaction occurring at below normal reaction conditions 1.45

        Returns:
            max_values (list): [max_pn4 massflow, chemicals, stream name] 
        """
        iname_list = []
        oname_list = []
        MFin = np.zeros((instream, 5))
        MFout = np.zeros((outstream, 5))

        for i in range(instream):
            iname = self.aspen.Tree.FindNode(f"\\Data\\Flowsheet\\Section\\GLOBAL\\Input\\INSTRM\\{bname}\\#{i}").value
            iname_list.append(iname)
            for j, property_name in enumerate(self.chemicals):
                property_value = self.aspen.Tree.FindNode(f'\\Data\\Streams\\{iname}\\Output\\MASSFLOW\\MIXED\\{property_name}').ValueForUnit(10, 1)
                MFin[i][j] = property_value

        for i in range(outstream):
            oname = self.aspen.Tree.FindNode(f"\\Data\\Flowsheet\\Section\\GLOBAL\\Input\\OUTSTRM\\{bname}\\#{i}").value
            oname_list.append(oname)
            for j, property_name in enumerate(self.chemicals):
                property_value = self.aspen.Tree.FindNode(f'\\Data\\Streams\\{oname}\\Output\\MASSFLOW\\MIXED\\{property_name}').ValueForUnit(10, 1)
                MFout[i][j] = property_value

        combined_matrix = np.vstack((MFin, MFout))
        pn4_MF = self.pn4 * combined_matrix
        max_value = np.max(pn4_MF)
        max_positions = np.argwhere(pn4_MF == max_value)
        streams = iname_list + oname_list

        max_values = []
        for max_row, max_col in max_positions:
            chemical_name = self.chemicals[max_col]
            stream_name = streams[max_row]
            max_values.append((max_value, chemical_name, stream_name))
        return max_values
    
    def calculate_and_print_block_performance(self, instream, outstream, bname, rxn, pn9, pn10):
        block_max_values = self.calculate_block_performance(instream, outstream, bname, rxn, pn9, pn10, self.aspen)
        print(f"{bname} Target Prediction:")
        for value, chemical_name, stream_name in block_max_values:
            print(f"Value: {value:.3f}, Target Chemical: {chemical_name}, Target Stream: {stream_name}")
            
    # Calculation of FEDR and transfer to FEDI
    def calculate_vp(self, T, index_TC, PP, target_chemical, target_stream):
        """Vapor pressure is calculated at here.
           Chemical labeled with "A" will apply Extended Antonie Equation for Vapor Pressure calculation.
           Those with "P" will utilize NIST TDE Polynomial describe Liquied Vapor interaction.
           If you apply different method for vapor pressure calculation, you have to customize the equation at here.

        Args:
            T (float): Temperature
            index_TC (list): Targeted chemical label
            PP (float): Pressure 
            target_chemical (str): Targeted chemical (with highest pn4 massflow)
            target_stream (str): Targeted stream (with highest pn4 massflow)

        Returns:
            VP (float): Vapor pressure
        """
        if  (self.LT[index_TC] <= T <= self.UT[index_TC]) and (self.para[index_TC][0] == np.str_('A')):
            VP = np.exp(float(self.para[index_TC][1]) + float(self.para[index_TC][2]) / ((T +273.15)\
                + float(self.para[index_TC][3])) + float(self.para[index_TC][4]) * (T+273.15)\
                + float(self.para[index_TC][5]) * np.log(T+273.15)\
                + float(self.para[index_TC][6]) * ((T+273.15) ** float(self.para[index_TC][7]))) # bar
            VP = 100*VP # bar to kPa
        elif  (self.LT[index_TC] <= T <= self.UT[index_TC]) and (self.para[index_TC][0] == np.str_('P')):
            VP = np.exp(float(self.para[index_TC][1]) + float(self.para[index_TC][2])/ (T + 273.15)\
                        + float(self.para[index_TC][3])*np.log(T + 273.15) + float(self.para[index_TC][4])*(T + 273.15)\
                        + float(self.para[index_TC][5])* (T + 273.15)**2 + float(self.para[index_TC][6])/(T + 273.15)**2\
                        + float(self.para[index_TC][7])*(T + 273.15)**6 + float(self.para[index_TC][8])/(T +273.15)**4) # bar
            VP = 100*VP # bar to kPa
        else:
            VP = PP * self.aspen.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\MASSFRAC\\MIXED\\{target_chemical}').Value
        return VP

    def rxn_or_not(self, rxn, target_chemical, F1, pn1, F, pn2, F4, pn9, pn10, pn3, index_TC):
        base_value = 4.76 * ((F1*pn1 + F*pn2) * pn3 * self.pn4[index_TC] * self.pn7 * self.pn8) ** (1/3)

        for i, reaction_chemicals in self.chem_rxn.items():
            if rxn == 1:
                return 4.76 * ((F1*pn1 + F*pn2 + F4*pn9*pn10) * pn3 * self.pn4[index_TC] * self.pn7 * self.pn8) ** (1/3)
        return base_value

    def calculate_FEDI(self, instream, outstream, bname, rxn, pn9, pn10):
        block_performance_results = self.calculate_block_performance(instream, outstream, bname, rxn, pn9, pn10)
        
        max_FEDI = -1  
        stream_lst = []
        FEDR_lst = []
        FEDI_lst = []
        result_lst = []

        for result in block_performance_results:
            target_chemical = result[1]
            target_stream = result[2]
            index_TC = self.chemicals.index(target_chemical)
            stream_lst.append(target_stream)
            
            MF = self.aspen.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\MASSFLOW\\MIXED\\{target_chemical}').ValueForUnit(10, 1) # kg/s
            T = self.aspen.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\RES_TEMP').ValueForUnit(22, 4) # C
            PP = self.aspen.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\RES_PRES').ValueForUnit(20, 10) # kPa
            VF_tot = self.aspen.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\VOLFLMX\\MIXED').ValueForUnit(12, 7) # cum/hr
            MFrac = self.aspen.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\MASSFRAC\\MIXED\\{target_chemical}').Value # dimensionless
            
            VP = self.calculate_vp(T, index_TC, PP, target_chemical, target_stream)
            VF = VF_tot * MFrac
            F1 = 0.1 * MF * -self.HC[index_TC] / 3.148
            F2 = 1.304 * 10 ** (-3) * PP * VF
            F3 = ((1.0 * 10 ** (-3)) / (T + 273)) * (PP - VP) ** 2 * VF

            # Conditional statement according to the definition from ref.Rangaiah_2013
            if (VP > self.AP) and (PP > VP):
                F = F2 + F3
            elif (VP > self.AP) and (PP < VP):
                F = F2
            else:
                F = F3

            if rxn == 1:
                Q = self.aspen.Tree.FindNode(f'\\Data\\Blocks\\{bname}\\Output\\QCALC').ValueForUnit(13, 14) # kW = kJ/s
                if Q < 0:
                    F4 = -Q / 3.148 
                else:
                    F4 = 0
            else:
                F4 = 0

            if T < self.FP[index_TC]:
                pn1 = 1.1
            elif self.FP[index_TC] < T < self.AI[index_TC]:
                pn1 = 1.75
            else:
                pn1 = 1.95

            if (VP > self.AP) and (PP > VP):
                pn2 = 1 + abs(0.6 * ((PP - VP) / PP))
            elif (VP > self.AP) and (PP < VP):
                pn2 = 1 + abs(0.4 * ((PP - VP) / PP))
            elif (VP < self.AP) and (PP > self.AP):
                pn2 = 1 + abs(0.2 * ((PP - VP) / PP))
            else:
                pn2 = 1.1

            index_pn3 = max(self.NF[index_TC], self.NR[index_TC])
            CI = self.aspen.Tree.FindNode(f'\\Data\\Streams\\{target_stream}\\Output\\MASSFLOW\\MIXED\\{target_chemical}').ValueForUnit(10, 12)  # Chemical Inventory in tons/hr

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

            FEDR = self.rxn_or_not(rxn, target_chemical, F1, pn1, F, pn2, F4, pn9, pn10, pn3, index_TC)
            FEDR_lst.append(FEDR)

            if 10 <= FEDR <= 200:
                FEDI = FEDR / 2
            elif 200 < FEDR <= 300:
                FEDI = 100
            else:
                FEDI = 0
            FEDI_lst.append(FEDI)
        
        for stream, FEDR, FEDI in zip(stream_lst, FEDR_lst, FEDI_lst):
            result_lst.append((stream, FEDR, FEDI))
        max_FEDI = max(result_lst, key=lambda x: x[1])
        print(f"{bname}_FEDI: {max_FEDI[2]:.4f}")
        return max(FEDI_lst)

    def process_block(self, instream, outstream, bname, rxn, pn9, pn10, display=0):
        if display == 1:
            self.calculate_and_print_block_performance(instream, outstream, bname, rxn, pn9, pn10) 
        else:
            self.calculate_block_performance(instream, outstream, bname, rxn, pn9, pn10)
        block_FEDI = self.calculate_FEDI(instream, outstream, bname, rxn, pn9, pn10)  
        return block_FEDI
    
    def DI_calculator(self):
        self.Aspen()
        total_FEDI = 0
        for block in self.blocks:
            block_FEDI = self.process_block(block["instream"], block["outstream"], block["bname"], block["rxn"], block["pn9"], block["pn10"])
            total_FEDI += block_FEDI
        print(f'total_FEDI: {total_FEDI:.4f}')