'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-14 12:27:23
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
'ACVE-A1...2
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude("Function_Library")
'-------------------------------------------------------------------------------
iCurves = 2
sGname = "ACVE"
'-------------------------------------------------------------------------------Call Data.Root.Clear()
Call Data.Root.Clear()
Call Data.Root.ChannelGroups.Add("DIAdem_reference").Activate()
Call ChnLinGenImp("t1",750001,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACVE-A1"")= 1000*sqrt(2)*sin(100*pi*Ch(""t1"")) *(0.9*Ch(""t1"")/5*abs(Ch(""t1"")<5) + abs(Ch(""t1"")>5)*abs(Ch(""t1"")<10)*(0.9+0.1*(Ch(""t1"")-5)/5))") 
'ACVE-A1	                       1000*sqrt(2)*sin(100*pi*     t    ) *(0.9*     t    /5*  H(5-   t      ) +   H(     t    -5)*  H(10-  t       )*(0.9+0.1*(     t    -5)/5))
Call ChnLinGenImp("t2",750001,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACVE-A2"")= 1000*sqrt(2)*sin(100*pi*Ch(""t2"")) *Ch(""t2"")/10*abs(Ch(""t2"")<10)") 
'ACVE-A2	                       1000*sqrt(2)*sin(100*pi*     t    ) *     t    /10*  H(10-  t       )

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
Call eval2excel(1,1,iCurves,sGname,1)  ' init
Call eval2excel(10,1,iCurves,sGname,2) ' Up/sqrt(2)
Call eval2excel(11,1,iCurves,sGname,4) ' Urms
Call eval2excel(8,1,iCurves,sGname,6)  ' Rate of rise1&2
Call eval2excel(6,1,iCurves,sGname,10) ' frequency

'--- Fill results of DIAdem curves into Excel Workbook
'do not call init
Call eval2excel(10,2,iCurves,sGname,12) ' Up/sqrt(2)
Call eval2excel(11,2,iCurves,sGname,14) ' Urms
Call eval2excel(8,2,iCurves,sGname,16)  ' Rate of rise1&2
Call eval2excel(6,2,iCurves,sGname,20)  ' frequency
Call datetime(iCurves,sGname,22)        ' date & time
Call SaveAndCloseExcel                  ' final finishing