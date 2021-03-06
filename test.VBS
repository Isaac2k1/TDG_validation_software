'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2016-03-14 18:08:57
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
' Labelling as subscript
'Call Subsequence()
' Initialization
Dim t,i,maxindex, DIAdem_reference, MyFolders, QRangeMin, QRangeMax, sCurveName, sFileName, sGroup, sGname, sReferenceTimeAxis, sTDGtimeAxis, sTDGamplitudeAxis, sNumber, iIndex, iAmplidude, iCurves
Dim I1k, I2k, Id, T1AC, T2AC, TDC, Ttot, f0, Df, Ioff, phi, DR
Dim aI1k, aI2k, aId, aT1AC, aT2AC, aTDC, aTtot, af0, aDf, aIoff, aphi, aDR
'--- Procedure to assign values to calculate SHACI curves
aI1k  = array(1,     0.8,  1.2,    1,   1,  1.25, 1.11)
aI2k  = array(1,     0.7,  1.5,    1,   1,   2.5, 1.43)
aId   = array(1,       1,    1,    1,   1,  1.25,    1)
aT1AC = array(400,   400,  400,  400, 400,   400,   16)
aT2AC = array(16,     16,   16,   16,  16,   400,  400)
aTDC  = array(45,    120,   45,   45,  45,    80,   80)
aTtot = array(214, 214.7,  214,  214, 214, 197.3, 1080)
af0   = array(50,     50,   50,   50,  50,    50,   50)
aDf   = array(0,     2.5, -2.5,    0,   0,    -3,   -7)
aIoff = array(0,       0,    0, -100,   0,   -20,    0)
aphi  = array(90,     90,   90,   90,  90,   -45,   45)
aDR   = array(90,     90,   90,   90,  30,    90,   90)
'--- add time offset to time axis
'Call MatScalAdd("[1]/t1",-0.2,0)


for i=0 to 6
I1k  = aI1k(i)
I2k  = aI2k(i)
Id   = aId(i)
T1AC = aT1AC(i)
T2AC = aT2AC(i)
TDC  = aTDC(i)
Ttot = aTtot(i)
f0   = af0(i)
Df   = aDf(i)
Ioff = aIoff(i)
phi  = aphi(i)
DR   = aDR(i)
msgbox "SHACI"&i+1&"   "&I1k&"    "&I2k&"    "&Id&"    "&T1AC&"    "&T2AC&"    "&TDC&"    "&Ttot&"    "&f0&"    "&Df&"    "&Ioff&"    "&phi&"    "&DR
next 