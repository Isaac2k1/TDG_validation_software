'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-14 12:27:23
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
'DCI-A1...2
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude("Function_Library")
'-------------------------------------------------------------------------------
iCurves = 2
sGname = "DCI"
'-------------------------------------------------------------------------------Call Data.Root.Clear()
Call Data.Root.Clear()
Call Data.Root.ChannelGroups.Add("DIAdem_reference").Activate()
Call ChnLinGenImp("t1",10001,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""DCI-A1"")= 1000+0*Ch(""t1"")") ' Calculates y channel
'                               
Call ChnLinGenImp("t2",10001,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""DCI-A2"")= 1000+29.4*(sin(100*pi*Ch(""t2""))+ sin(100*pi*Ch(""t2"")*2)/2+ sin(100*pi*Ch(""t2"")*3)/3+ sin(100*pi*Ch(""t2"")*4)/4+ sin(100*pi*Ch(""t2"")*5)/5+ sin(100*pi*Ch(""t2"")*6)/6+ sin(100*pi*Ch(""t2"")*7)/7+ sin(100*pi*Ch(""t2"")*8)/8+ sin(100*pi*Ch(""t2"")*9)/9)") 
'                               1000+29.4*(sin(100*pi*     t2   )+ sin(100*pi*     t    *2)/2+ sin(100*pi*     t    *3)/3+ sin(100*pi*     t    *4)/4+ sin(100*pi*     t    *5)/5+ sin(100*pi*     t    *6)/6+ sin(100*pi*     t    *7)/7+ sin(100*pi*     t    *8)/8+ sin(100*pi*     t    *9)/9))

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

'--- Fill results of DIAdem curves into Excel Workbook
'do not call init
Call eval2excel(4,2,iCurves,sGname,6) ' Urms
Call eval2excel(12,2,iCurves,sGname,8)' ripple
'final finishing
Call datetime(iCurves,sGname,10)      ' date & time
Call SaveAndCloseExcel               ' final finishing