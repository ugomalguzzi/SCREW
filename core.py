import numpy as np

# LEGENDA
# AF = Angolo di Fresa
# LC = Luce di Calandra
# PF = Profondità Filetto
# CL = Cresta Longitudinale
# LC = Fondo Longitudinale
# AC = Area Cilndro
# CR = Calcolo Cresta
# FD = Calcolo Fondo
# CT = Cresta Trasversale
# FT = Fondo Trasversale
# AV = Calcolo Area Vite
# EV = Volume Vuoto al Giro
# LF = Luce tra i Fianchi
# LT = Luce Tetraedrica
# TL = Traferro Longitudinale
# QM = Maximum Output
# TP = Tempo di residenza zone piene
# TNP = Tempo di residenza zone non piene
# PL= Peso Calandrabile
# PC = Peso di una camera a "C"
# GT = Shear rate sec-1
# CC = superficie di contatto cilindro in una zona
# CV = superf. Contatto con le viti in una zona

# INPUT
DV = 51.80      # [mm]      Diametro vite 
PA = 45.00      # [mm]      Passo 
NP = 1          # [-]       Numero principi
DA = 0.5        # [kg/dm3]  Densità apparente
IA = 44.00      # [mm]      Interasse 
VR = 40         # [RPM]     Velocità rotazione
RC = 1          # [-]       Rapporto di compressione
LZ = 510        # [mm]      Lunghezza zona
DE = 1.4        # [kg/dm3]  Densità PVC
PF = 8.4        # [mm]      Profondità filetto
CR = 18.70      # [mm]      Lunghezza cresta
FD = 20         # [mm]      Lunghezza fondo

# OUTPUT
CL = CR * np.sqrt(np.pi**2*DV**2+PA**2)/(np.pi*DV)
FL = FD * np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)/(np.pi*(DV-2*PF))
LC = IA-DV+PF
AF = np.arctan((FL-CL)/(2*PF))
TL = (PA/2*NP) - CL - ((DV-IA)*(FL-CL)/(2*PF))
PF = LC+DV-IA

print("CL = " + str(CL))
print("FL = " + str(FL))
print("LC = " + str(LC))
print("AF = " + str(AF))
print("TL = " + str(TL))
print("PF = " + str(PF))

CX = IA / DV        # Rapporto di compenetrazione X
CO = (-np.arctan(CX/np.sqrt(-CX**2+1))+1.5708)*360/np.pi # Angolo compenetrazione
CP = (DV-IA)*100/DV     # Compenetrazione viti %

print("CX = " + str(CX))
print("CO = " + str(CO))
print("CP = " + str(CP))

AC = (np.pi*DV**2*(0.5-CO/720)+IA*np.sqrt(DV**2-IA**2)/2)/100 # [cm2] Area cilindro 

print("AC = " + str(AC))

cCR = CR*np.pi*DV/(np.sqrt(np.pi**2*DV**2+PA**2)) # Calcolo cresta
cFD = FD*np.pi*(DV-2*PF)/np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)
CT = CR*np.sqrt(np.pi**2*DV**2+PA**2)/PA # Cresta Trasversale
FT = FD*np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)/PA # Fonso Trasversale

print("cCR = " + str(cCR))
print("cFD = " + str(cFD))
print("CT = " + str(CT))
print("FT = " + str(FT))

AV = ((np.pi*(DV-2*PF)**2/4)+NP*PF*(CT+FT)/2)/100 # [cm2] Area vite

print("AV = " + str(AV))

EV = ((AC*100)-2*(AV*100))*PA/1000 # [cm3] Volume vuoto al giro

print("EV = " + str(EV))

LF = CR*TL*np.cos(AF)/CL # Luce fianchi
LT = ((PA/(2*NP))-CL)# Luce tetraedrica

print("LF = " + str(LF))
print("LT = " + str(LT))

QM = (EV*DA*(VR/60))/1000*3600 # [kg/h] Portata massima

print("QM = " + str(QM))

TP = EV*LZ*DA/(QM*PA) # [s] Tempo di permanenza zone piene
TNP = LZ/((VR/60)*PA) # [s] Tempo di permanenza zone NON piene

print("TP = " + str(TP))
print("TNP = " + str(TNP))

PL = (np.sqrt(np.pi**2*(DV-2*PF)**2+PA**2)+np.sqrt(np.pi**2*DV**2+PA**2)) # Peso calandrabile
PL2 = PL*LC*LZ*CL*DE/(2*PA)/1000 
PC = EV*0.6/(2*NP) # Peso di una camera a "C"

print("PL = " + str(PL))
print("PL2 = " + str(PL2))
print("PC = " + str(PC))

GT = np.pi*(VR/60)*PF*2/LC # [1/s] Massimo gradiente velocità di taglio

print("GT = " + str(GT))

CC = DV*np.pi*(360-CO)*LZ*NP*((PA/NP)-CL)/(PA*180)/100 # Superficie di contatto cilindro in una zona
CV = ((FL*np.pi*(DV-2*PF))+(2*np.pi*PF*(DV-PF)/np.cos(AF))+(CL*DV*np.pi*CO/360)) # Superficie di contatto con le viti in una zona
CV2 = CV*2*LZ*NP/PA/100

print("CC = " + str(CC))
print("CV = " + str(CV))
print("CV2 = " + str(CV2))

