'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-14 12:27:23
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
'DCV-A1...3
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude("Function_Library")
'-------------------------------------------------------------------------------
iCurves = 3
sGname = "DCV"
'-------------------------------------------------------------------------------Call Data.Root.Clear()
Call Data.Root.Clear()
Call Data.Root.ChannelGroups.Add("DIAdem_reference").Activate()
Call ChnLinGenImp("t1",10001,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""DCV-A1"")= 1000+0*Ch(""t1"")")
Call ChnLinGenImp("t2",10001,-0.1,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""DCV-A2"")= 1000+29.4*(sin(100*pi*Ch(""t2""))+ sin(100*pi*Ch(""t2"")*2)/2+ sin(100*pi*Ch(""t2"")*3)/3+ sin(100*pi*Ch(""t2"")*4)/4+ sin(100*pi*Ch(""t2"")*5)/5+ sin(100*pi*Ch(""t2"")*6)/6+ sin(100*pi*Ch(""t2"")*7)/7+ sin(100*pi*Ch(""t2"")*8)/8+ sin(100*pi*Ch(""t2"")*9)/9)") 
'                              (1000+29.4*(sin(100*pi*     t    )+ sin(100*pi*     t    *2)/2+ sin(100*pi*     t    *3)/3+ sin(100*pi*     t    *4)/4+ sin(100*pi*     t    *5)/5+ sin(100*pi*     t    *6)/6+ sin(100*pi*     t    *7)/7+ sin(100*pi*     t    *8)/8+ sin(100*pi*     t    *9)/9))
Call ChnLinGenImp("t3",750001,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""DCV-A3"")= 1000*((abs(5>Ch(""t3""))*0.9*Ch(""t3"")/5)+(abs(5<=Ch(""t3""))*abs(Ch(""t3"")<=10)*(0.9+0.1*(Ch(""t3"")-5)/5)))") 
'                               1000*(   H(5-     t    )*0.9*     t    /5 +(1-H(5-     t   ))*  H(  10-  t       )*(0.9+0.1*(     t    -5)/5))

'--- Define path to TDG output files ---
Call InitMyFolders(sGname)

'--- Convert calculated channels to waveform channels ---
Call ConvCH2wfm(iCurves,sGname)

'--- Create new channel group and determine its index
Call Data.Root.ChannelGroups.Add("Min/Max_calculations").Activate()
maxindex= (data.Root.ChannelGroups.Count)
'msgbox(maxindex)

'--- Subtract DIAdem_reference time axis 't' from each groups time axis and determine MIN & MAX
Call MinMaxOFdiff(iCurves,sGname)

'--- Fill results of TDG reference curves into Excel Workbook
'--- parameters of call are: Case,Reference to group in DataPortal, number of curves, group name of curves, Excel column number
'--- Use sub eval2excel(iCalc,iRef,iCurves,sGname,sColumn)

Call eval2excel(1,1,iCurves,sGname,1) ' init
Call eval2excel(4,1,iCurves,sGname,2) ' Urms
Call eval2excel(12,1,iCurves,sGname,4)' ripple
'Call eval2excel(6,1,iCurves,sGname,6) ' frequency -- this doesn't work :-(


'--- Fill results of DIAdem curves into Excel Workbook
'do not call init
Call eval2excel(4,2,iCurves,sGname,8) ' Urms
Call eval2excel(12,2,iCurves,sGname,10)' ripple
'Call eval2excel(6,2,iCurves,sGname,12) ' frequency -- this doesn't work :-(

'final finishing
Call datetime(iCurves,sGname,14)      ' date & time
Call SaveAndCloseExcel               ' final finishing