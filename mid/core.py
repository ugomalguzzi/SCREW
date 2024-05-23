import numpy as np
import pandas as pd
import xlsxwriter
import time
import os

# pyinstaller --noconfirm --onefile --console --icon C:/repo_git/SCREW/img/maximelt.ico C:/repo_git/SCREW/core.py --distpath C:\repo_git\SCREW --workpath C:\repo_git\build_temp --specpath C:\repo_git\build_temp

# INPUT
CD = 'C52'      # [-]       Codice vite                                  
DV = 51.80      # [mm]      Diametro vite                               
PA = 45.00      # [mm]      Passo                                         
NP = 1          # [-]       Numero principi                           
DA = 0.5        # [kg/dm3]  Densità apparente alimentazione (pellet)   
IA = 44.00      # [mm]      Interasse                                
VR = 40         # [RPM]     Velocità rotazione                          
RC = 1          # [-]       Rapporto di compressione                     
LZ = 510        # [mm]      Lunghezza zona                                
DE = 1.4        # [kg/dm3]  Densità PVC                                   
PF = 8.4        # [mm]      Profondità filetto                            
CR = 18.70      # [mm]      Lunghezza cresta                              
FD = 20         # [mm]      Lunghezza fondo                             
NF = 4          # [-]       Numero fresature                                    
DF = 15         # [mm]      Diametro fresa                                     

# INPUT LIST
CD_list = ['C52', 'C52', 'C52', 'C52', 'C52', 'C52', 'C52', 'C52', 'C52', 'C52']
DV_list = [51.80, 51.80, 51.80, 51.80, 51.80, 51.80, 51.80, 51.80, 51.80, 51.80]
PA_list = [45.00, 45.00, 45.00, 45.00, 45.00, 45.00, 45.00, 45.00, 45.00, 45.00]
NP_list = [1    , 1    , 1    , 1    , 1    , 1    , 1    , 1    , 1    , 1    ]
DA_list = [0.5  , 0.5  , 0.5  , 0.5  , 0.5  , 0.5  , 0.5  , 0.5  , 0.5  , 0.5  ]
IA_list = [44.00, 44.00, 44.00, 44.00, 44.00, 44.00, 44.00, 44.00, 44.00, 44.00]
VR_list = [40   , 40   , 40   , 40   , 40   , 40   , 40   , 40   , 40   , 40   ]
RC_list = [1    , 1    , 1    , 1    , 1    , 1    , 1    , 1    , 1    , 1    ]
LZ_list = [510  , 510  , 510  , 510  , 510  , 510  , 510  , 510  , 510  , 510  ]
DE_list = [1.4  , 1.4  , 1.4  , 1.4  , 1.4  , 1.4  , 1.4  , 1.4  , 1.4  , 1.4  ]
PF_list = [8.4  , 8.4  , 8.4  , 8.4  , 8.4  , 8.4  , 8.4  , 8.4  , 8.4  , 8.4  ]
CR_list = [18.70, 18.70, 18.70, 18.70, 18.70, 18.70, 18.70, 18.70, 18.70, 18.70]
FD_list = [20   , 20   , 20   , 20   , 20   , 20   , 20   , 20   , 20   , 20   ]
NF_list = [4    , 4    , 4    , 4    , 4    , 4    , 4    , 4    , 4    , 4    ]
DF_list = [15   , 15   , 15   , 15   , 15   , 15   , 15   , 15   , 15   , 15   ]

# OUTPUT
CL   = CR * np.sqrt(np.pi**2*DV**2+PA**2)/(np.pi*DV)                                 # [mm ]     Cresta longitudinale                                         
FL   = FD * np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)/(np.pi*(DV-2*PF))                   # [mm]      Fondo longitudinale                                         
LC   = IA-DV+PF                                                                      # [mm]      Luce di calandra                                         
AF   = np.arctan((FL-CL)/(2*PF))                                                     # [rad]     Angolo di fresa                                         
TL   = (PA/2*NP) - CL - ((DV-IA)*(FL-CL)/(2*PF))                                     # [mm]      Traferro longitudinale                                         
PFc  = LC+DV-IA                                                                      # [mm]      Profondità filetto calcolata                                        
CX   = IA / DV                                                                       # [-]       Rapporto di compenetrazione X                                         
CO   = (-np.arctan(CX/np.sqrt(-CX**2+1))+1.5708)*360/np.pi                           # [deg]     Angolo compenetrazione                                         
CP   = (DV-IA)*100/DV                                                                # [%]       Compenetrazione viti %                                         
AC   = (np.pi*DV**2*(0.5-CO/720)+IA*np.sqrt(DV**2-IA**2)/2)/100                      # [cm2]     Area cilindro                                          
cCR  = CR*np.pi*DV/(np.sqrt(np.pi**2*DV**2+PA**2))                                   # [mm]      Calcolo cresta                                         
cFD  = FD*np.pi*(DV-2*PF)/np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)                       # [mm]      Calcolo fondo                                         
CT   = CR*np.sqrt(np.pi**2*DV**2+PA**2)/PA                                           # [mm]      Cresta Trasversale                                         
FT   = FD*np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)/PA                                    # [mm]      Fondo Trasversale                                         
AV   = ((np.pi*(DV-2*PF)**2/4)+NP*PF*(CT+FT)/2)/100                                  # [cm2]     Area vite                                         
EV   = ((AC*100)-2*(AV*100))*PA/1000                                                 # [cm3]     Volume vuoto al giro                                         
LF   = CR*TL*np.cos(AF)/CL                                                           # [mm]      Luce fianchi                                         
LT   = ((PA/(2*NP))-CL)                                                              # [mm]      Luce tetraedrica                                         
QM   = (EV*DA*(VR/60))/1000*3600                                                     # [kg/h]    Portata massima                                         
TP   = EV*LZ*DA/(QM*PA)                                                              # [s]       Tempo di permanenza zone piene                                         
TNP  = LZ/((VR/60)*PA)                                                               # [s]       Tempo di permanenza zone NON piene                                         
PL   = (np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)+np.sqrt(np.pi**2*DV**2+PA**2))          # [kg]      Peso calandrabile                                         
PL2  = PL*LC*LZ*CL*DE/(2*PA)/1000                                                    # [kg]      Peso calandrabile2                                         
PC   = EV*0.6/(2*NP)                                                                 # [kg]      Peso di una camera a C                                           
GT   = np.pi*(VR/60)*PF*2/LC                                                         # [1/s]     Massimo gradiente velocità di taglio                                         
CC   = DV*np.pi*(360-CO)*LZ*NP*((PA/NP)-CL)/(PA*180)/100                             # [mm2]     Superficie di contatto cilindro in una zona                                         
CV   = ((FL*np.pi*(DV-2*PF))+(2*np.pi*PF*(DV-PF)/np.cos(AF))+(CL*DV*np.pi*CO/360))   # [?]       Superficie di contatto con le viti in una zona                                         
CV2  = CV*2*LZ*NP/PA/100  

# OUTPUT LIST
CL_list  = []
FL_list  = []
LC_list  = []
AF_list  = []
TL_list  = []
PFc_list = []
CX_list  = []
CO_list  = []
CP_list  = []
AC_list  = []
cCR_list = []
cFD_list = []
CT_list  = []
FT_list  = []
AV_list  = []
EV_list  = []
LF_list  = []
LT_list  = []
QM_list  = []
TP_list  = []
TNP_list = []
PL_list  = []
PL2_list = []
PC_list  = []
GT_list  = []
CC_list  = []
CV_list  = []
CV2_list = []



for CD, DV, PA, NP, DA, IA, VR, RC, LZ, DE, PF, CR, FD, NF, DF in zip(CD_list, DV_list, PA_list, NP_list, DA_list, IA_list, VR_list, RC_list, LZ_list, DE_list, PF_list, CR_list, FD_list, NF_list, DF_list):
    CL  = CR * np.sqrt(np.pi**2*DV**2+PA**2)/(np.pi*DV)                                 # [mm ]     Cresta longitudinale                                         
    FL  = FD * np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)/(np.pi*(DV-2*PF))                   # [mm]      Fondo longitudinale                                         
    LC  = IA-DV+PF                                                                      # [mm]      Luce di calandra                                         
    AF  = np.arctan((FL-CL)/(2*PF))                                                     # [rad]     Angolo di fresa                                         
    TL  = (PA/2*NP) - CL - ((DV-IA)*(FL-CL)/(2*PF))                                     # [mm]      Traferro longitudinale                                         
    PFc  = LC+DV-IA                                                                     # [mm]      Profondità filetto                                         
    CX  = IA / DV                                                                       # [-]       Rapporto di compenetrazione X                                         
    CO  = (-np.arctan(CX/np.sqrt(-CX**2+1))+1.5708)*360/np.pi                           # [deg]     Angolo compenetrazione                                         
    CP  = (DV-IA)*100/DV                                                                # [%]       Compenetrazione viti %                                         
    AC  = (np.pi*DV**2*(0.5-CO/720)+IA*np.sqrt(DV**2-IA**2)/2)/100                      # [cm2]     Area cilindro                                          
    cCR = CR*np.pi*DV/(np.sqrt(np.pi**2*DV**2+PA**2))                                   # [mm]      Calcolo cresta                                         
    cFD = FD*np.pi*(DV-2*PF)/np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)                       # [mm]      Calcolo fondo                                         
    CT  = CR*np.sqrt(np.pi**2*DV**2+PA**2)/PA                                           # [mm]      Cresta Trasversale                                         
    FT  = FD*np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)/PA                                    # [mm]      Fondo Trasversale                                         
    AV  = ((np.pi*(DV-2*PF)**2/4)+NP*PF*(CT+FT)/2)/100                                  # [cm2]     Area vite                                         
    EV  = ((AC*100)-2*(AV*100))*PA/1000                                                 # [cm3]     Volume vuoto al giro                                         
    LF  = CR*TL*np.cos(AF)/CL                                                           # [mm]      Luce fianchi                                         
    LT  = ((PA/(2*NP))-CL)                                                              # [mm]      Luce tetraedrica                                         
    QM  = (EV*DA*(VR/60))/1000*3600                                                     # [kg/h]    Portata massima                                         
    TP  = EV*LZ*DA/(QM*PA)                                                              # [s]       Tempo di permanenza zone piene                                         
    TNP = LZ/((VR/60)*PA)                                                               # [s]       Tempo di permanenza zone NON piene                                         
    PL  = (np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)+np.sqrt(np.pi**2*DV**2+PA**2))          # [kg]      Peso calandrabile                                         
    PL2 = PL*LC*LZ*CL*DE/(2*PA)/1000                                                    # [kg]      Peso calandrabile2                                         
    PC  = EV*0.6/(2*NP)                                                                 # [kg]      Peso di una camera a C                                           
    GT  = np.pi*(VR/60)*PF*2/LC                                                         # [1/s]     Massimo gradiente velocità di taglio                                         
    CC  = DV*np.pi*(360-CO)*LZ*NP*((PA/NP)-CL)/(PA*180)/100                             # [mm2]     Superficie di contatto cilindro in una zona                                         
    CV  = ((FL*np.pi*(DV-2*PF))+(2*np.pi*PF*(DV-PF)/np.cos(AF))+(CL*DV*np.pi*CO/360))   # [?]       Superficie di contatto con le viti in una zona                                         
    CV2 = CV*2*LZ*NP/PA/100                                                             # [cm2]     Superficie di contatto con le viti in una zona                                         

    CL_list.append(CL  )
    FL_list.append(FL  )
    LC_list.append(LC  )
    AF_list.append(AF  )
    TL_list.append(TL  )
    PFc_list.append(PFc)
    CX_list.append(CX  )
    CO_list.append(CO  )
    CP_list.append(CP  )
    AC_list.append(AC  )
    cCR_list.append(cCR)
    cFD_list.append(cFD)
    CT_list.append(CT  )
    FT_list.append(FT  )
    AV_list.append(AV  )
    EV_list.append(EV  )
    LF_list.append(LF  )
    LT_list.append(LT  )
    QM_list.append(QM  )
    TP_list.append(TP  )
    TNP_list.append(TNP)
    PL_list.append(PL  )
    PL2_list.append(PL2)
    PC_list.append(PC  )
    GT_list.append(GT  )
    CC_list.append(CC  )
    CV_list.append(CV  )
    CV2_list.append(CV2)

# print("CL = " + str(CL))
# print("FL = " + str(FL))
# print("LC = " + str(LC))
# print("AF = " + str(AF))
# print("TL = " + str(TL))
# print("PF = " + str(PF))
# print("CX = " + str(CX))
# print("CO = " + str(CO))
# print("CP = " + str(CP))
# print("AC = " + str(AC))
# print("cCR = " + str(cCR))
# print("cFD = " + str(cFD))
# print("CT = " + str(CT))
# print("FT = " + str(FT))
# print("AV = " + str(AV))
# print("EV = " + str(EV))
# print("LF = " + str(LF))
# print("LT = " + str(LT))
# print("QM = " + str(QM))
# print("TP = " + str(TP))
# print("TNP = " + str(TNP))
# print("PL = " + str(PL))
# print("PL2 = " + str(PL2))
# print("PC = " + str(PC))
# print("GT = " + str(GT))
# print("CC = " + str(CC))
# print("CV = " + str(CV))
# print("CV2 = " + str(CV2))

# Create excel file
caption = "\\" + CD + time.strftime("_%Y-%m-%d_%H-%M-%S")

workbook = xlsxwriter.Workbook(os.getcwd()+caption + ".xlsx")

worksheet1 = workbook.add_worksheet()

cell_format = workbook.add_format()
cell_format.set_border()

worksheet1.insert_image('A3', "C:\\repo_git\\SCREW\\img\\GAP.png")

floating_point = workbook.add_format({'num_format': '#,##0.00'})
bordered = workbook.add_format({'border': 2})
bold = workbook.add_format({'bold': True})

# Some sample data for the table.
input_data = [
    ["Codice vite                             ", "[-]     " , "CD   "] + CD_list,
    ["Diametro vite                           ", "[mm]    " , "DV   "] + DV_list,
    ["Passo                                   ", "[mm]    " , "PA   "] + PA_list,
    ["Numero principi                         ", "[-]     " , "NP   "] + NP_list,
    ["Densità apparente alimentazione (pellet)", "[kg/dm3]" , "DA   "] + DA_list,
    ["Interasse                               ", "[mm]    " , "IA   "] + IA_list,
    ["Velocità rotazione                      ", "[RPM]   " , "VR   "] + VR_list,
    ["Rapporto di compressione                ", "[-]     " , "RC   "] + RC_list,
    ["Lunghezza zona                          ", "[mm]    " , "LZ   "] + LZ_list,
    ["Densità PVC                             ", "[kg/dm3]" , "DE   "] + DE_list,
    ["Profondità filetto                      ", "[mm]    " , "PF   "] + PF_list,
    ["Lunghezza cresta                        ", "[mm]    " , "CR   "] + CR_list,
    ["Lunghezza fondo                         ", "[mm]    " , "FD   "] + FD_list,
    ["Numero fresature                        ", "[-]     " , "NF   "] + NF_list,
    ["Diametro fresa                          ", "[mm]    " , "DF   "] + DF_list
]

output_data = [
    ["Cresta longitudinale                          " , "[mm]  " , "CL   "] + CL_list,
    ["Fondo longitudinale                           " , "[mm]  " , "FL   "] + FL_list,
    ["Luce di calandra                              " , "[mm]  " , "LC   "] + LC_list,
    ["Angolo di fresa                               " , "[rad] " , "AF   "] + AF_list,
    ["Traferro longitudinale                        " , "[mm]  " , "TL   "] + TL_list,
    ["Profondità filetto calcolata                  " , "[mm]  " , "PF   "] + PFc_list,
    ["Rapporto di compenetrazione X                 " , "[-]   " , "CX   "] + CX_list,
    ["Angolo compenetrazione                        " , "[deg] " , "CO   "] + CO_list,
    ["Compenetrazione viti %                        " , "[%]   " , "CP   "] + CP_list,
    ["Area cilindro                                 " , "[cm2] " , "AC   "] + AC_list,
    ["Calcolo cresta                                " , "[mm]  " , "cCR  "] + cCR_list,
    ["Calcolo fondo                                 " , "[mm]  " , "cFD  "] + cFD_list,
    ["Cresta Trasversale                            " , "[mm]  " , "CT   "] + CT_list,
    ["Fondo Trasversale                             " , "[mm]  " , "FT   "] + FT_list,
    ["Area vite                                     " , "[cm2] " , "AV   "] + AV_list,
    ["Volume vuoto al giro                          " , "[cm3] " , "EV   "] + EV_list,
    ["Luce fianchi                                  " , "[mm]  " , "LF   "] + LF_list,
    ["Luce tetraedrica                              " , "[mm]  " , "LT   "] + LT_list,
    ["Portata massima                               " , "[kg/h]" , "QM   "] + QM_list,
    ["Tempo di permanenza zone piene                " , "[s]   " , "TP   "] + TP_list,
    ["Tempo di permanenza zone NON piene            " , "[s]   " , "TNP  "] + TNP_list,
    ["Peso calandrabile                             " , "[kg]  " , "PL   "] + PL_list,
    ["Peso calandrabile2                            " , "[kg]  " , "PL2  "] + PL2_list,
    ["Peso di una camera a C                        " , "[kg]  " , "PC   "] + PC_list,
    ["Massimo gradiente velocità di taglio          " , "[1/s] " , "GT   "] + GT_list,
    ["Superficie di contatto cilindro in una zona   " , "[mm2] " , "CC   "] + CC_list,
    ["Superficie di contatto con le viti in una zona" , "[?]   " , "CV   "] + CV_list,
    ["Superficie di contatto con le viti in una zona" , "[cm2] " , "CV2  "] + CV2_list
]
N_PARA = len(output_data)

input_columns = [     
    {"header": "INPUT"},
    {"header": "Units"},
    {"header": "Code"},
    {"header": "Zone 1"},
    {"header": "Zone 2"},
    {"header": "Zone 3"},
    {"header": "Zone 4"},
    {"header": "Zone 5"},
    {"header": "Zone 6"},
    {"header": "Zone 7"},
    {"header": "Zone 8"},
    {"header": "Zone 9"},
    {"header": "Zone 10"}
]
N_ZONE = len(input_columns)

output_columns = [     
    {"header": "OUTPUT"},
    {"header": "Units"},
    {"header": "Code"},
    {"header": "Zone 1"},
    {"header": "Zone 2"},
    {"header": "Zone 3"},
    {"header": "Zone 4"},
    {"header": "Zone 5"},
    {"header": "Zone 6"},
    {"header": "Zone 7"},
    {"header": "Zone 8"},
    {"header": "Zone 9"},
    {"header": "Zone 10"}
]

worksheet1.write("D1", caption)

worksheet1.set_column("D:D", 42)
worksheet1.set_column("E:P", 12, floating_point)

input_options =  {
        "data": input_data,
        "columns": input_columns,
        "style": 'Table Style Light 10'
    }

output_options =  {
        "data": output_data,
        "columns": output_columns,
        "style": 'Table Style Light 13'
    }

where_input = "D4:"+chr(ord("D")+len(input_columns)-1)+str(len(input_data)+4)
worksheet1.add_table(where_input, input_options)

where_output = "D"+str(len(input_data)+4+3)+":"+chr(ord("D")+len(input_columns)-1)+str(len(input_data)+4+3+len(output_data))
worksheet1.add_table(where_output, output_options)


workbook.close()

print(pd.read_excel('C52_2024-03-20_15-48-06.xlsx', index_col=3, header=3).keys())