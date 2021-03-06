'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created  : 2017-07-12 17:55:55
'-- Modified 
'-- Author: Marco Mailand
'-- Comment: collection of important subroutines for IEC 61083-4 TDG
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
'Labelling as subscript
'Call Subsequence()
' Initialization
Dim t, i,text, maxindex, DIAdem_reference, MyFolders, MyExcelFile, QRangeMin, QRangeMax, Excel, NoOfPeaks, StartTime
Dim absDTmin, absDTmax, curDTmin, curDTmax, absDYmin, absDYmax, curDYmin, curDYmax, sCell, sColumn, RR1, RR2, curDT, curDY, RRR
Dim sCurveName, sCurveName1, sCurveName2, sCurveName3, sCurveName4, sCurveName5, sCurveName6, sCurveName7, sCurveName8, sCurveName9
Dim sFileName, sGroup, sGname, sReferenceTimeAxis, sTDGtimeAxis, sTDGamplitudeAxis, sNumber, iIndex, iAmplidude, iCurves, iCurve, iRef
Dim sFormula, sSymbols, sValues, sUnit, DR, Pts, t0, dt, A,B,C,M,N,O,P,Q,S,G,D,V,F
Dim sFormula1, sONOFF, sEXPamp, sFREQch, sDCpart
RR1 = array(0,0,0,0,0,0,0,0,0,0,0,0)
RR2 = array(0,0,0,0,0,0,0,0,0,0,0,0)
StartTime = 5
NoOfPeaks = 20
'Dim I1k, I2k, Id, T1AC, T2AC, TDC, Ttot, f0, Df, Ioff, Phi, tt1, tt2
'--- Procedure to set SHACI curve parameters
'--- Curve # is equal to case, parameters contain also time axis
'--- Procedure to assign values to calculate SHACI curves
Sub SetSHACI(iCurve,sSymbols,sValues,sUnit,Pts,t0,dt)
'These Parameters are
' tt1 - signal start time (preset, give by standard)
' tt2 - signal stop time (to be adjusted to fit the end of the curve data)
' Ttot - should somehow be proportional to tt2-tt1, but is to be adjusted
'      sSymbols = array("I1k", "I2k", "Id", "T1AC", "T2AC", "TDC", "Ttot", "f0", "Df", "Ioff", "Phi", "DR", "sUnit", "tt1",    "tt2", "Pts", "t0", "dt")
       sSymbols = array("A"  , "B"  , "C" , "O"   , "P"   , "Q"  , "S"   , "G" , "D" , "V"   , "F"  , "DR", "sUnit", "M"  ,    "N"  , "Pts", "t0", "dt")
select case iCurve
'Table A.8 - Reference values from '42WG20 - 19 IEC 61083-4 Sept2012.docx'
 Case 1 sValues = array( 1000,  1000, 1000,    0.4,  0.016, 0.045, 0.2147,   50,    0,      0,  pi/2,   90,     "A",  0.02, 0.215042, 11501,  0.0, 2E-5)
 Case 2 sValues = array(  800,   700, 1000,    0.4,  0.016, 0.120, 0.1947,   50,  2.5,      0,  pi/2,   90,     "A",  0.02, 0.214756, 11501,  0.0, 2E-5)
 Case 3 sValues = array( 1200,  1500, 1000,    0.4,  0.016, 0.045, 0.1940,   50, -2.5,      0,  pi/2,   90,     "A",  0.02, 0.214704, 11501,  0.0, 2E-5)
 Case 4 sValues = array( 1000,  1000, 1000,    0.4,  0.016, 0.045, 0.2140,   50,    0,   -100,  pi/2,   90,     "A",  0.02, 0.215042, 11501,  0.0, 2E-5)
 Case 5 sValues = array( 1000,  1000, 1000,    0.4,  0.016, 0.045, 0.2140,   50,    0,      0,  pi/2,   30,     "A",  0.02, 0.215042, 11501,  0.0, 2E-5)
 Case 6 sValues = array( 1250,  2500, 1250,    0.4,    0.4, 0.080, 0.1773,   50,   -3,    -20, -pi/4,   90,     "A",  0.02,  0.19792, 11501,  0.0, 2E-5)
 Case 7 sValues = array( 1110,  1430, 1000,  0.016,    0.4, 0.080,  1.060,   50,   -7,      0,  pi/4,   90,     "A",  0.02,  1.08126, 57001,  0.0, 2E-5)
End Select
   sUnit        = array(  "A",   "A",  "A",    "s",    "s",   "s",    "s", "Hz", "Hz",    "A", "deg",  "%",     "A",   "s",      "s",   "1",  "s",  "s")
Pts = sValues(15)
t0  = sValues(16)
dt  = sValues(17)
'' Create a msgbox for debugging of parameter handling
'text=""
'for i= 0 to 17
'text=text&sSymbols(i)&" = "&sValues(i)&" "&sUnit(i)& vbCrLf 
'next
'msgbox text

End Sub
'--- Procedure to assign root folder for actual curve group ---
Sub InitMyFolders(sGname)
  MyFolders = "C:\Users\chmamai\OneDrive for Business\IEC\TDG_ASCII\"&sGname&"\"
 'MyFolders = "Y:\OneDrive - ABB\IEC\TDG_ASCII\"&sGname&"\"
  MyExcelFile = "C:\Users\chmamai\OneDrive for Business\IEC\TDG_validation_software\IEC61083_4.xlsx"
End Sub

'--- Convert calculated channels to waveform channels ---
'Call Subsequence()
Sub ConvCH2wfm(iCurves,sGname)
  for i=1 to iCurves
   sNumber = Str(i,"d")
   sCurveName = "[1]/"&sGname&"-A"&sNumber
   Call ChnToWfChn("[1]/t"&sNumber,sCurveName,0,"WfXRelative") '... X,ChnNoStr,XChnDelete,WfXStartTimeMode 
  next

'--- Read in the TDG output files and transform them into waveform ---
 for i=1 to iCurves
  sNumber = Str(i,"d")
  sCurveName = "[1]/"&sGname&"-A"&sNumber
  sFileName = MyFolders&sGname&"-A"&sNumber&".txt"
  iIndex = "["&Str(i+1,"d")&"]/time"
  iAmplidude = "["&Str(i+1,"d")&"]/amplitude"
  'Call MsgBox(iIndex&"_"&iAmplidude, vbInformation)
  'call msgBox(sFileName)
  Call DataFileLoad(sFileName,"TDG_TXT_DataPlugin","Load") '... DataFilename,FileImportFilter,ImportAction
  Call ChnToWfChn(iIndex,iAmplidude,0,"WfXRelative") '... X,ChnNoStr,XChnDelete,WfXStartTimeMode 
 next
End Sub

'--- Calculate differeneces of time and amplitude axes of TDG curves and DIAdem curves
'--- Reason is to validate TDG results with DIAdem results
'Call Subsequence()
Sub MinMaxOFdiff(iCurves,sGname)
'--- Subtract DIAdem_reference time axis 't' from each groups time axis and determine MIN & MAX of delta of all time axes
'--- Subtract DIAdem_reference amplitude axis 'y' from each groups amplitude axis and determine MIN & MAX of delta of all amplitude axes
'--- Set minima and maxima to highest or lowest values and display the four results. If very small then DIAdem and TDG calculate similar curves.
absDTmin=1e99
absDTmax=0
absDYmin=1e99
absDYmax=0
 for i=1 to iCurves
  sNumber = Str(i,"d")
  sGroup = Str(i+1,"d")
  sCurveName = "[1]/"&sGname&"-A"&Str(i,"d")
  sReferenceTimeAxis = "[1]/t"&sNumber  '&sNumber
  sTDGtimeAxis = "["& sGroup &"]/time"
  sTDGamplitudeAxis = "["& sGroup &"]/amplitude"
  ' Calculate difference of time axes between DIAdem curve and TDG curve
  Call MatSub(sReferenceTimeAxis,sTDGtimeAxis)
  'rename default result with proper name
  Data.Root.ChannelGroups(maxindex).Channels("MatrixSubtractedZ1").Name = "dTime"&sNumber
  ' Determine the minimum and maximum of each difference and find lowest minimum and highest maximum
  curDTmin=Data.Root.ChannelGroups(maxindex).Channels("dTime"&sNumber).Properties("minimum").Value
  if curDTmin < absDTmin then absDTmin = curDTmin
  curDTmax=Data.Root.ChannelGroups(maxindex).Channels("dTime"&sNumber).Properties("maximum").Value
  if curDTmax > absDTmax then absDTmax = curDTmax
  ' Calculate difference of amplitude axes between DIAdem curve and TDG curve
    Call MatSub(sCurveName,sTDGamplitudeAxis)
  Data.Root.ChannelGroups(maxindex).Channels("MatrixSubtractedZ1").Name = "dAmpl"&sNumber
  ' Determine the minimum and maximum of each difference and find lowest minimum and highest maximum
  curDYmin=Data.Root.ChannelGroups(maxindex).Channels("dAmpl"&sNumber).Properties("minimum").Value
  if curDYmin < absDYmin then absDYmin = curDYmin
  curDYmax=Data.Root.ChannelGroups(maxindex).Channels("dAmpl"&sNumber).Properties("maximum").Value
  if curDYmax > absDYmax then absDYmax = curDYmax
 next
text = "minimum DT (s): "& absDTmin & vbCrLf & "maximum DT (s): "& absDTmax & vbCrLf & "minimum DY (A): "& absDYmin & vbCrLf & "maximum DT (A): "& absDYmax
msgbox text
End Sub

'=======================================================
'--- Calculate values and fill them into Excel workbook
'--- This is the new routine for both, the DIAdem curves or the TDG curves, thus can switch the type of addressing by iRef
'--- Parameters:
'--- iCalc   - number of subroutine to be executed
'--- iRef    - type of referencing of curve in DIAdem data portal: 1 - DIAdem curves, 2 - TDG curves
'--- sGname  - group name, name of the TDG curve familiy
'--- sColumn - Excel colum to receive data
Sub eval2excel(iCalc,iRef,iCurves,sGname,sColumn)

 Select Case iCalc
  
  Case 1 ' initialize Excel - only available in case iRef = 1 - must be executed only once in the beginning, before first Excel access
   Select Case iRef
    Case 1
     Set Excel = CreateObject("Excel.Application")
     Excel.Visible = TRUE
     Call Excel.Workbooks.Open(MyExcelFile) 
     Excel.Worksheets(sGname).Activate
     for i=1 to iCurves
      Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = sGname&"-A"&Str(i,"d")
     next
    Case 2
   End Select
  
  Case 2 ' calculation of Upp/sqrt(8) 
     for i=1 to iCurves
      Select Case iRef
      Case 1
       sGroup = "[1]/" ' DIAdem curves always calculated at first position
       sCurveName = sGroup & sGname & "-A"&Str(i,"d")
      Case 2
       sGroup = "[" & Str(i+1,"d")& "]" ' Group string of TDG curves is incremented with iCurves
       sCurveName = sGroup & "/amplitude"
      End Select
      curDYmin=ChnPropGet(sCurveName, "minimum")
      curDYmax=ChnPropGet(sCurveName, "maximum")
      absDYmax=(curDYmax-curDYmin)/sqrt(8)
      Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = absDYmax
     next
  
   Case 3 'calculation of true Urms
   for i=1 to iCurves
    Select Case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & "t" & Str(i,"d")
      sCurveName1 = sGroup & sGname & "-A"&Str(i,"d")
      sCurveName2 = sGroup & "RMS"
     Case 2
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "time"
      sCurveName1 = sGroup & "amplitude"
      sCurveName2 = sGroup & "RMS" & Str(i,"d")
     End Select 
      Call ChnRMS(sCurveName, sCurveName1, sCurveName2, 100)
      curDYmin = ChnPropGet(sCurveName2, "minimum")
      curDYmax = ChnPropGet(sCurveName2, "maximum")
      if curDYmin <> curDYmax then call msgbox("Error on RMS calculation")
      Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = curDYmax
   next
  
   Case 4 ' calculation of Urms_ac - is now the same code as Urms - this is certainly wrong!
   for i=1 to iCurves
    Select Case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & "t" & Str(i,"d")
      sCurveName1 = sGroup & sGname & "-A"&Str(i,"d")
      sCurveName2 = sGroup & "RMS"
     Case 2 
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "time"
      sCurveName1 = sGroup & "amplitude"
      sCurveName2 = sGroup & "RMS" & Str(i,"d")
     End Select
      Call ChnRMS(sCurveName, sCurveName1, sCurveName2, 100)
      curDYmin = ChnPropGet(sCurveName2, "minimum")
      curDYmax = ChnPropGet(sCurveName2, "maximum")
      if curDYmin <> curDYmax then call msgbox("Error on RMS calculation")
      Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = curDYmax
   next
  
   Case 5 'calculation of THD
   for i=1 to iCurves
    Select Case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & sGname & "-A"&Str(i,"d")
     Case 2
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "amplitude"
    End Select
    Call ChnSINAD(,sCurveName,0,10)
    Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = 100*SINADTHDPlusNoise
   next

   Case 6 'calculation of frequency
    for i=1 to iCurves
     Select Case iRef
      Case 1
       sGroup = "[1]/" ' DIAdem curves always calculated at first position
       sCurveName  = sGroup & sGname & "-A"&Str(i,"d")
      Case 2
       sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
       sCurveName  = sGroup & "amplitude"
     End Select
     Call ChnSINAD(,sCurveName,0,10) 
     Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = SINADDetectedFrequency
    next
  
   Case 7 ' calculation of Uss
   for i=1 to iCurves
    Select case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & sGname &"-A" & Str(i,"d")
     Case 2 
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "amplitude"
     End Select
    curDYmin=ChnPropGet(sCurveName, "minimum")
    curDYmax=ChnPropGet(sCurveName, "maximum")
    absDYmax=(curDYmax-curDYmin)
    Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = absDYmax    
   next
  
   Case 8 ' calculation of Rate of rise1 and 2
   for i=1 to iCurves
    Select case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & "t" & Str(i,"d")
      sCurveName1 = sGroup & sGname & "-A"& Str(i,"d")
      sCurveName2 = sGroup & "time2nd" & Str(i,"d")
      sCurveName3 = sGroup & "amplitude2nd" & Str(i,"d")
      sCurveName4 = sGroup & "timeOpeak2nd" & Str(i,"d")
      sCurveName5 = sGroup & "peaks2nd" & Str(i,"d")
      sCurveName6 = sGroup & "time1st" & Str(i,"d")
      sCurveName7 = sGroup & "amplitude1st" & Str(i,"d")
      sCurveName8 = sGroup & "timeOpeak1st" & Str(i,"d")
      sCurveName9 = sGroup & "peaks1st" & Str(i,"d")
     Case 2 
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "time"
      sCurveName1 = sGroup & "amplitude"
      sCurveName2 = sGroup & "time2nd"
      sCurveName3 = sGroup & "amplitude2nd"
      sCurveName4 = sGroup & "timeOpeak2nd"
      sCurveName5 = sGroup & "peaks2nd"
      sCurveName6 = sGroup & "time1st"
      sCurveName7 = sGroup & "amplitude1st"
      sCurveName8 = sGroup & "timeOpeak1st"
      sCurveName9 = sGroup & "peaks1st"
     End Select
    'select only the part of the curve after start time
'   Call ChnTimeAreaCopy("[1]/t"&Str(i,"d"),"[1]/"&sGname&"-A"&Str(i,"d"),"[1]/time2nd"&Str(i,"d"),"[1]/amplitude2nd"&Str(i,"d"), StartTime, StartTime+0.2)
    Call ChnTimeAreaCopy(sCurveName, sCurveName1, sCurveName2, sCurveName3, StartTime, StartTime+0.2)
    'find 4 peaks right to start time
'   Call ChnPeakFind("[1]/time2nd"&Str(i,"d"),"[1]/amplitude2nd"&Str(i,"d"),"[1]/timeOpeak2nd"&Str(i,"d"),"[1]/peaks2nd"&Str(i,"d"),NoOfPeaks,"Max.Peaks","Time")
    Call ChnPeakFind(sCurveName2, sCurveName3, sCurveName4, sCurveName5, NoOfPeaks,"Max.Peaks","Time")
'   curDT = CHD(4,"[1]/timeOpeak2nd"&Str(i,"d")) - CHD(1,"[1]/timeOpeak2nd"&Str(i,"d"))
    curDT = CHD(4,sCurveName4) - CHD(1,sCurveName4)
'   curDY = CHD(4,"[1]/peaks2nd"&Str(i,"d"))     - CHD(1,"[1]/peaks2nd"&Str(i,"d"))
    curDY = CHD(4,sCurveName5)     - CHD(1,sCurveName5)
    RR2(i) = curDY / curDT
    Excel.Worksheets(sGname).Cells(3+i,sColumn+2).Value = RR2(i)
    'select only the part of the curve after start time
'   Call ChnTimeAreaCopy("[1]/t"&Str(i,"d"),"[1]/"&sGname&"-A"&Str(i,"d"),"[1]/time1st"&Str(i,"d"),"[1]/amplitude1st"&Str(i,"d"), StartTime-curDT, StartTime+0.2)
    Call ChnTimeAreaCopy(sCurveName, sCurveName1, sCurveName6, sCurveName7, StartTime-curDT, StartTime+0.2)
    'find 4 peaks right to start time
'   Call ChnPeakFind("[1]/time1st"&Str(i,"d"),"[1]/amplitude1st"&Str(i,"d"),"[1]/timeOpeak1st"&Str(i,"d"),"[1]/peaks1st"&Str(i,"d"),NoOfPeaks,"Max.Peaks","Time")
    Call ChnPeakFind(sCurveName6, sCurveName7, sCurveName8, sCurveName9, NoOfPeaks,"Max.Peaks","Time")
    curDT = CHD(2,sCurveName8) - CHD(1,sCurveName8)
    curDY = CHD(2,sCurveName9) - CHD(1,sCurveName9)
    RR1(i) = curDY / curDT   
    Excel.Worksheets(sGname).Cells(3+i,sColumn).Value   = RR1(i)
   next
  
   Case 9 ' calculation of Rate of ????
   for i=1 to iCurves
    Select case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & "t" & Str(i,"d")
      sCurveName1 = sGroup & sGname & "-A"& Str(i,"d")
      sCurveName2 = sGroup & "time2nd" & Str(i,"d")
      sCurveName3 = sGroup & "amplitude2nd" & Str(i,"d")
      sCurveName4 = sGroup & "timeOpeak2nd" & Str(i,"d")
      sCurveName5 = sGroup & "peaks2nd" & Str(i,"d")
      sCurveName6 = sGroup & "time1st" & Str(i,"d")
      sCurveName7 = sGroup & "amplitude1st" & Str(i,"d")
      sCurveName8 = sGroup & "timeOpeak1st" & Str(i,"d")
      sCurveName9 = sGroup & "peaks1st" & Str(i,"d")
     Case 2 
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "time"
      sCurveName1 = sGroup & "amplitude"
      sCurveName2 = sGroup & "time2nd"
      sCurveName3 = sGroup & "amplitude2nd"
      sCurveName4 = sGroup & "timeOpeak2nd"
      sCurveName5 = sGroup & "peaks2nd"
      sCurveName6 = sGroup & "time1st"
      sCurveName7 = sGroup & "amplitude1st"
      sCurveName8 = sGroup & "timeOpeak1st"
      sCurveName9 = sGroup & "peaks1st"
     End Select
    ' This is not yet fully programmed, because calculation of Rate of rise 2 is calculated in Case 8 too.
    Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = "RR2"
   next
 
   Case 10 ' calculation of Up/sqrt(2)
   for i=1 to iCurves
    Select case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName1 = sGroup & sGname & "-A"& Str(i,"d")
     Case 2 
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName1 = sGroup & "amplitude"
     End Select
    curDYmax=ChnPropGet(sCurveName1, "maximum")/sqrt(2)
'    msgbox(sCurveName1 & " = " & curDYmax)
    Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = curDYmax
   next
  
   Case 11 ' calculation of RMS envelope
   for i=1 to iCurves
    Select case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & "t" & Str(i,"d")
      sCurveName1 = sGroup & sGname & "-A"& Str(i,"d")
      sCurveName2 = sGroup & "time2nd" & Str(i,"d")
      sCurveName3 = sGroup & "amplitude2nd" & Str(i,"d")
      sCurveName4 = sGroup & "timeOpeak2nd" & Str(i,"d")
      sCurveName5 = sGroup & "peaks2nd" & Str(i,"d")
      sCurveName6 = sGroup & "time1st" & Str(i,"d")
      sCurveName7 = sGroup & "amplitude1st" & Str(i,"d")
      sCurveName8 = sGroup & "timeOpeak1st" & Str(i,"d")
      sCurveName9 = sGroup & "peaks1st" & Str(i,"d")
     Case 2 
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "time"
      sCurveName1 = sGroup & "amplitude"
      sCurveName2 = sGroup & "time2nd"
      sCurveName3 = sGroup & "amplitude2nd"
      sCurveName4 = sGroup & "timeOpeak2nd"
      sCurveName5 = sGroup & "peaks2nd"
      sCurveName6 = sGroup & "time1st"
      sCurveName7 = sGroup & "amplitude1st"
      sCurveName8 = sGroup & "timeOpeak1st"
      sCurveName9 = sGroup & "peaks1st"
     End Select

    sGroup = Str(i+1,"d")
    sCurveName = "[1]/"&sGname&"-A"&Str(i,"d") ' for TDG use this: "[" & sGroup & "]/apmplitude", "minimum"
    Call ChnRMS("[1]/t" & Str(i,"d"),"[1]/"&sGname&"-A"&Str(i,"d"),"[1]/RMS"& Str(i,"d"),0.1)
    curDYmax = ChnPropGet("[1]/RMS"& Str(i,"d"), "maximum")
'    msgbox(curDYmax)
    Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = curDYmax
   next
 
   Case 12 ' calculation of (ripple)
   for i=1 to iCurves
    Select case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & "t" & Str(i,"d")
      sCurveName1 = sGroup & sGname & "-A"& Str(i,"d")
      sCurveName2 = sGroup & "time2nd" & Str(i,"d")
      sCurveName3 = sGroup & "amplitude2nd" & Str(i,"d")
      sCurveName4 = sGroup & "timeOpeak2nd" & Str(i,"d")
      sCurveName5 = sGroup & "peaks2nd" & Str(i,"d")
      sCurveName6 = sGroup & "time1st" & Str(i,"d")
      sCurveName7 = sGroup & "amplitude1st" & Str(i,"d")
      sCurveName8 = sGroup & "timeOpeak1st" & Str(i,"d")
      sCurveName9 = sGroup & "peaks1st" & Str(i,"d")
     Case 2 
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "time"
      sCurveName1 = sGroup & "amplitude"
      sCurveName2 = sGroup & "time2nd"
      sCurveName3 = sGroup & "amplitude2nd"
      sCurveName4 = sGroup & "timeOpeak2nd"
      sCurveName5 = sGroup & "peaks2nd"
      sCurveName6 = sGroup & "time1st"
      sCurveName7 = sGroup & "amplitude1st"
      sCurveName8 = sGroup & "timeOpeak1st"
      sCurveName9 = sGroup & "peaks1st"
     End Select
    sGroup = Str(i+1,"d")
    sCurveName = "[1]/"&sGname&"-A"&Str(i,"d") ' for TDG use this: "[" & sGroup & "]/apmplitude", "minimum"
    curDYmin = Val(ChnPropGet("[1]/"&sGname&"-A"&Str(i,"d"), "minimum"))
    curDYmax = Val(ChnPropGet("[1]/"&sGname&"-A"&Str(i,"d"), "maximum"))
    curDY = curDYmax + curDYmin
    msgbox(curDY)
    curDY = (curDYmax - curDYmin) / (curDYmax + curDYmin) / 2.0 * 100.0
    Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = curDY
   next
  
  Case 13 'calculation of Up, select the bigger absolute value of either the minimum or the maximum
   for i=1 to iCurves
    Select case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & "t" & Str(i,"d")
      sCurveName1 = sGroup & sGname & "-A"& Str(i,"d")
      sCurveName2 = sGroup & "time2nd" & Str(i,"d")
      sCurveName3 = sGroup & "amplitude2nd" & Str(i,"d")
      sCurveName4 = sGroup & "timeOpeak2nd" & Str(i,"d")
      sCurveName5 = sGroup & "peaks2nd" & Str(i,"d")
      sCurveName6 = sGroup & "time1st" & Str(i,"d")
      sCurveName7 = sGroup & "amplitude1st" & Str(i,"d")
      sCurveName8 = sGroup & "timeOpeak1st" & Str(i,"d")
      sCurveName9 = sGroup & "peaks1st" & Str(i,"d")
     Case 2 
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "time"
      sCurveName1 = sGroup & "amplitude"
      sCurveName2 = sGroup & "time2nd"
      sCurveName3 = sGroup & "amplitude2nd"
      sCurveName4 = sGroup & "timeOpeak2nd"
      sCurveName5 = sGroup & "peaks2nd"
      sCurveName6 = sGroup & "time1st"
      sCurveName7 = sGroup & "amplitude1st"
      sCurveName8 = sGroup & "timeOpeak1st"
      sCurveName9 = sGroup & "peaks1st"
     End Select
    
    'DIAdem
    
   'Excel.Worksheets(sGname).Range(sColumn &Str(3+i,"d")).Value = curDYmax
   'Excel.Worksheets(sGname).Cells(3+i,sColumn).Value
   next
 
 Case 14
  'calculation of ?
   for i=1 to iCurves
    Select case iRef
     Case 1
      sGroup = "[1]/" ' DIAdem curves always calculated at first position
      sCurveName  = sGroup & "t" & Str(i,"d")
      sCurveName1 = sGroup & sGname & "-A"& Str(i,"d")
      sCurveName2 = sGroup & "time2nd" & Str(i,"d")
      sCurveName3 = sGroup & "amplitude2nd" & Str(i,"d")
      sCurveName4 = sGroup & "timeOpeak2nd" & Str(i,"d")
      sCurveName5 = sGroup & "peaks2nd" & Str(i,"d")
      sCurveName6 = sGroup & "time1st" & Str(i,"d")
      sCurveName7 = sGroup & "amplitude1st" & Str(i,"d")
      sCurveName8 = sGroup & "timeOpeak1st" & Str(i,"d")
      sCurveName9 = sGroup & "peaks1st" & Str(i,"d")
     Case 2 
      sGroup = "[" & Str(i+1,"d")& "]/" ' Group string of TDG curves is incremented with iCurves
      sCurveName  = sGroup & "time"
      sCurveName1 = sGroup & "amplitude"
      sCurveName2 = sGroup & "time2nd"
      sCurveName3 = sGroup & "amplitude2nd"
      sCurveName4 = sGroup & "timeOpeak2nd"
      sCurveName5 = sGroup & "peaks2nd"
      sCurveName6 = sGroup & "time1st"
      sCurveName7 = sGroup & "amplitude1st"
      sCurveName8 = sGroup & "timeOpeak1st"
      sCurveName9 = sGroup & "peaks1st"
     End Select
    sGroup = Str(i+1,"d")
    sCurveName = "[1]/"&sGname&"-A"&Str(i,"d") ' for TDG use this: "[" & sGroup & "]/apmplitude", "minimum"
    'DIAdem
    
   'Excel.Worksheets(sGname).Range(sColumn &Str(3+i,"d")).Value = curDYmax
   'Excel.Worksheets(sGname).Cells(3+i,sColumn).Value
   next

   End Select  ' of DIAdem2excel
 End Sub ' eval2excel 
'========================================================
'=== Fill in date and time of test, to be called last ===
  
Sub datetime(iCurves,sGname,sColumn)
' MsgBox "The name of the active sheet is " & Excel.Worksheets(sGname).Name
  for i=1 to iCurves
   Excel.Worksheets(sGname).Cells(3+i,sColumn).Value = Year(now) &"-"& month(now) &"-"& day(now)
   Excel.Worksheets(sGname).Cells(3+i,sColumn+1).Value = Time
  next
End Sub

'=== Save and Close Excel, to be called at the very end ===

Sub SaveAndCloseExcel
dim result
result = MsgBox ("Save Workbook?", vbYesNo, "SAVE")

Select Case result
 Case vbYes
   Excel.ActiveSheet.Parent.Saved = TRUE
   Call Excel.ActiveWorkbook.Save
 Case vbNo
  ' Excel.DisplayAlerts = FALSE
End Select

result = MsgBox ("Close Excel?", vbYesNo, "CLOSE")

Select Case result
 Case vbYes
   Call Excel.Quit
   Set Excel = Nothing
 Case vbNo
   Excel.DisplayAlerts = FALSE
End Select
End Sub