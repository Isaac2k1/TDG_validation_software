'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2017-06-06 16:27:05
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

Dim iCurves, sGname, sGroup, sColumn, i, sCurveName, StartTime, NoOfPeaks, timeOpeak, RR1, RR2, curDT, curDY
RR1 = array(0,0,0,0,0,0,0,0,0,0,0,0)
RR2 = array(0,0,0,0,0,0,0,0,0,0,0,0)
iCurves = 2
sGname = "ACVE"
sGroup  = sGroup = Str(i+1,"d")
sColumn = 20
NoOfPeaks = 4
StartTime = 5

 for i=1 to iCurves
  msgbox("DIAdem"&i)
    sGroup = Str(i+1,"d")
    sCurveName = "[1]/"&sGname&"-A"&Str(i,"d") ' for TDG use this: "[" & sGroup & "]/RMS", "minimum"
    'msgbox("[" & sGroup & "]/time === [" & sGroup & "]/amplitude === [" & sGroup & "]/timeOpeak === [" & sGroup & "]/peaks")
    'Call ChnPeakFind("[" & sGroup & "]/time","[" & sGroup & "]/amplitude","[" & sGroup & "]/timeOpeak","[" & sGroup & "]/peaks",6,"Max.Peaks","Time")
    'select only the part of the curve after start time
    Call ChnTimeAreaCopy("[" & sGroup & "]/time","[" & sGroup & "]/amplitude","[" & sGroup & "]/time2nd","[" & sGroup & "]/amplitude2nd", StartTime, StartTime+0.1)
    'finde 4 peaks right to start time
    Call ChnPeakFind("[" & sGroup & "]/time2nd","[" & sGroup & "]/amplitude2nd","[" & sGroup & "]/timeOpeak2nd","[" & sGroup & "]/peaks2nd",NoOfPeaks,"Max.Peaks","Time")
    curDT = CHD(4,"[" & sGroup & "]/timeOpeak2nd") - CHD(1,"[" & sGroup & "]/timeOpeak2nd")
msgbox(curDT)
    curDY = CHD(4,"[" & sGroup & "]/peaks2nd")     - CHD(1,"[" & sGroup & "]/peaks2nd")
msgbox(curDY)
    RR2(i) = curDY / curDT
    'msgbox(curDT &"  "& curDY &"  "& RR1)
    'jump 4 timesteps to the left of StartTime
    'move StartTime to the left by four samples
msgbox("StartTime= "& StartTime &" RR1="& RR2(i))
    ', "RR1", i ,"= ", RRR)
    'select only the part of the curve after start time
    Call ChnTimeAreaCopy("[" & sGroup & "]/time","[" & sGroup & "]/amplitude","[" & sGroup & "]/time1st","[" & sGroup & "]/amplitude1st",StartTime-curDT, StartTime+0.1)
    'finde 4 peaks right to start time
    Call ChnPeakFind("[" & sGroup & "]/time1st","[" & sGroup & "]/amplitude1st","[" & sGroup & "]/timeOpeak1st","[" & sGroup & "]/peaks1st",NoOfPeaks,"Max.Peaks","Time")
    curDT = CHD(2,"[" & sGroup & "]/timeOpeak1st") - CHD(1,"[" & sGroup & "]/timeOpeak1st")
msgbox(curDT)
    curDY = CHD(2,"[" & sGroup & "]/peaks1st")     - CHD(1,"[" & sGroup & "]/peaks1st")
msgbox(curDY)
    RR1(i) = curDY / curDT   
msgbox(curDT &"  "& curDY &"  "& RR1(i))

    'Excel.Worksheets(sGname).Cells(3+i,sColumn).Value   = RR1(i)
    'Excel.Worksheets(sGname).Cells(3+i,sColumn+2).Value = RR2(i)
    next


   'calculation of Rate of rise1
   for i=1 to iCurves
    msgbox("TDG"&i)
    sGroup = Str(i+1,"d")
    sCurveName = "[1]/"&sGname&"-A"&Str(i,"d") ' for TDG use this: "[" & sGroup & "]/RMS", "minimum"
    'msgbox("[" & sGroup & "]/time === [" & sGroup & "]/amplitude === [" & sGroup & "]/timeOpeak === [" & sGroup & "]/peaks")

    'Call ChnPeakFind("[" & sGroup & "]/time","[" & sGroup & "]/amplitude","[" & sGroup & "]/timeOpeak","[" & sGroup & "]/peaks",6,"Max.Peaks","Time")
    'select only the part of the curve after start time
    Call ChnTimeAreaCopy("[1]/t"&Str(i,"d"),"[1]/"&sGname&"-A"&Str(i,"d"),"[1]/time2nd"&Str(i,"d"),"[1]/amplitude2nd"&Str(i,"d"), StartTime, StartTime+0.1)
    'finde 4 peaks right to start time
    Call ChnPeakFind("[1]/time2nd"&Str(i,"d"),"[1]/amplitude2nd"&Str(i,"d"),"[1]/timeOpeak2nd"&Str(i,"d"),"[1]/peaks2nd"&Str(i,"d"),NoOfPeaks,"Max.Peaks","Time")
    curDT = CHD(4,"[1]/timeOpeak2nd"&Str(i,"d")) - CHD(1,"[1]/timeOpeak2nd"&Str(i,"d"))
msgbox(curDT)
    curDY = CHD(4,"[1]/peaks2nd"&Str(i,"d"))     - CHD(1,"[1]/peaks2nd"&Str(i,"d"))
msgbox(curDY)

    RR2(i) = curDY / curDT
msgbox(curDT &"  "& curDY &"  "& RR2(i))
    'jump 4 timesteps to the left of StartTime
    'select only the part of the curve after start time
    Call ChnTimeAreaCopy("[1]/t"&Str(i,"d"),"[1]/"&sGname&"-A"&Str(i,"d"),"[1]/time1st"&Str(i,"d"),"[1]/amplitude1st"&Str(i,"d"), StartTime-curDT, StartTime+0.1)
    'finde 4 peaks right to start time
    Call ChnPeakFind("[1]/time1st"&Str(i,"d"),"[1]/amplitude1st"&Str(i,"d"),"[1]/timeOpeak1st"&Str(i,"d"),"[1]/peaks1st"&Str(i,"d"),NoOfPeaks,"Max.Peaks","Time")
    curDT = CHD(2,"[1]/timeOpeak1st"&Str(i,"d")) - CHD(1,"[1]/timeOpeak1st"&Str(i,"d"))
    curDY = CHD(2,"[1]/peaks1st"&Str(i,"d"))     - CHD(1,"[1]/peaks1st"&Str(i,"d"))
    RR1(i) = curDY / curDT   
msgbox(curDT &"  "& curDY &"  "& RR1(i))

    'Excel.Worksheets(sGname).Cells(3+i,sColumn).Value   = RR1(i)
    'Excel.Worksheets(sGname).Cells(3+i,sColumn+2).Value = RR2(i)
   next
