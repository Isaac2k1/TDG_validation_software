'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-14 12:27:23
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
'ACV-A1...6
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude("Function_Library")
'-------------------------------------------------------------------------------
iCurves = 6
sGname = "ACV"
'-------------------------------------------------------------------------------Call Data.Root.Clear()
Call Data.Root.Clear()
Call Data.Root.ChannelGroups.Add("DIAdem_reference").Activate()
Call ChnLinGenImp("t1",7501,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACV-A1"")= 1000*sqrt(2)*sin(2*50*pi*Ch(""t1""))") ' Calculates sinus of x channel
'                               1000*sqrt(2)*sin(2*50*pi*     t    )
Call ChnLinGenImp("t2",6251,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACV-A2"")= 1000*sqrt(2)*sin(2*60*pi*Ch(""t2"")) + 500") 
'                               1000*sqrt(2)*sin(2*60*pi*     t    ) + 500
Call ChnLinGenImp("t3",7501,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACV-A3"")= 1000*sqrt(2)*(sin(2*50*pi*Ch(""t3"")) + 0.05*sin(8*50*pi*Ch(""t3"")+pi*2/3))") 
'                               1000*sqrt(2)*(sin(2*50*pi*     t    ) + 0.05*sin(8*50*pi*     t    +pi*2/3))
Call ChnLinGenImp("t4",6251,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACV-A4"")= 1000*sqrt(2)*(sin(2*60*pi*Ch(""t4"")) + 0.07*sin(2*60*pi*Ch(""t4"")*5+pi*22/18))") 
'                               1000*sqrt(2)*(sin(2*60*pi*     t    ) + 0.07*sin(2*60*pi*     t    *5+pi*22/18))
Call ChnLinGenImp("t5",7501,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACV-A5"")= 1000*sqrt(2)*(sin(2*50*pi*Ch(""t5"")) + 0.05*sin(100*pi*Ch(""t5"")*3-pi) + 0.05*sin(100*pi*Ch(""t5"")*5)+0.02*sin(100*pi*Ch(""t5"")*7-pi)+0.01*sin(100*pi*Ch(""t5"")*9))") 
'                               1000*sqrt(2)*(sin(2*50*pi*     t    ) + 0.05*sin(100*pi*     t    *3-pi) + 0.05*sin(100*pi*     t    *5)+0.02*sin(100*pi*     t    *7-pi)+0.01*sin(100*pi*     t    *9))
Call ChnLinGenImp("t6",6251,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACV-A6"")= 1000*sqrt(2)*(sin(2*60*pi*Ch(""t6"")) + 0.02*sin(120*pi*Ch(""t6"")*3-pi) + 0.02*sin(120*pi*Ch(""t6"")*5)+0.0015*sin(120*pi*Ch(""t6"")*7-pi)+0.009*sin(120*pi*Ch(""t6"")*9))") 
'                               1000*sqrt(2)*(sin(2*60*pi*     t    ) + 0.02*sin(120*pi*     t    *3-pi) + 0.02*sin(120*pi*     t    *5)+0.0015*sin(120*pi*     t    *7-pi)+0.009*sin(120*pi*     t    *9))
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
Call eval2excel(2,1,iCurves,sGname,2) ' Up
Call eval2excel(3,1,iCurves,sGname,4) ' Urms
Call eval2excel(4,1,iCurves,sGname,6) ' Urmsac
Call eval2excel(5,1,iCurves,sGname,8) ' THD
Call eval2excel(6,1,iCurves,sGname,10) ' frequency

'--- Fill results of DIAdem curves into Excel Workbook
'do not call init
Call eval2excel(2,2,iCurves,sGname,12) ' Up
Call eval2excel(3,2,iCurves,sGname,14) ' Urms
Call eval2excel(4,2,iCurves,sGname,16) ' Urmsac
Call eval2excel(5,2,iCurves,sGname,18) ' THD
Call eval2excel(6,2,iCurves,sGname,20) ' frequency
Call datetime(iCurves,sGname,22) ' date & time
Call SaveAndCloseExcel