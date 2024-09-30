# PyNDI
PyNDI (NTU Damage Index calculator) provides a simulator-based calculator for damage index calculation.
## Overview
The damage index calculated in this program is basically based on the FEDI (Fire and Explosion Damage Index).
Firstly, hazadous chemicals will be defined in advance. 
Secondly, the program will search the and determine which unit may provide the most huge FEDI and let it represent the whole process.
By these two steps, the operation iteration can be reduced.
Finally, this progream connect with Aspen Plus for damage index evaluation.

## Example
In the main file, necessary information, including hazadous chemicals and its properties, reaction sets, blocks for calculation, and Simulator file name should be provided in the main.py file. The detailed descriptions about the necessary information have been added in the annotation of the DI function. The tested Simulator file is provided in the .zip file.

```python
import numpy as np
from DI import DI

chemicals = ['PO', 'PC', 'PG', 'GLY', 'GC'] 
NF = [4, 1, 1, 1, 1]
NH = [3, 1, 0, 1, 2]
NR = [2, 1, 0, 0, 0]
FP = [-37, 128, 101, 176, 149.5]
AI = [747, 435, 420, 393, 404]
HC = [-1917.4, -1818, -1838.19, -1654.3, -1594.2015]
 
para = [["A", 79.52407454, -5976.3,  0,  0, -10.686, 0.000011993, 2, 0], # PO PLXANT-1 in C & bar 
        ["A", 94.92707454, -10819,   0,  0, -12.068, 5.46E-06, 2, 0], # PC PLXANT-1 in C & bar 
        ["A", 123.9770745, -12483,   0,  0, -16.019, 6.46E-06, 2, 0], # PG PLXANT-1 in C & bar 
        ["A", 88.47307454, -13808,   0,  0, -10.088, 3.57E-19, 6, 0], # GLY PLXANT-1 in C & bar 
        ["P", 21.26683454, -10586.9, 0.0394255,  -0.0065509, 0, 0, 0, 0]] # GC PLTDEP-1 in C & bar 

chem_rxn = {
    1: ['PO'],
    2: ['PC', 'GLY'],
    3: ['PG', 'GC']
} 

aspenfile_name = "Fuzzy 2R_HX(RUN9).bkp"
pn7, pn8 = 1.1, 1.1
LT = [-111.93, 48.3, -60, 18.18, 10]
UT = [209.1, 504.85, 403.25, 576.85, 182.46]

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
```

## Installation
```python
import win32com.client as win32
import os
import time
import numpy as np
```

## Acknowlege
The primary developer is Pei-Jhen Wu with support from the following contributors.
* Bor-Yih Yu (National Taiwan University)
* Hsuan-Han Chiu (Purdue Universtiy)

## References
1. AIChE, Dow's Fire & Explosion Index Hazard Classification Guide. 1994.
2. Gupta, J.P., et al., Calculation of Fire and Explosion Index (F&EI) value for the Dow Guide taking credit for the loss control measures. Journal of Loss Prevention in the Process Industries, 2003. 16(4): p. 235-241.
3. Heikkilä, A.-M., Inherent safety in process plant design: an index-based approach. 1999: VTT Technical Research Centre of Finland.
4. Teh, S.Y., et al., A hybrid multi-objective optimization framework for preliminary process design based on health, safety and environmental impact. Processes, 2019. 7(4): p. 200.
5. Sharma, S., et al., Process design for economic, environmental and safety objectives with an application to the cumene process. Multi‐Objective Optimization in Chemical Engineering: Developments and Applications, 2013: p. 449-477.


## Citing
``` sourceCode
@article{,
author = {},
journal = {},
pages = {},
title = {},
volume = {},
year = {},
doi = {}
}
```