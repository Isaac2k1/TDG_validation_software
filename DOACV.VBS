'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-14 12:27:23
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
'DOACV-A1
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude("Function_Library")
'-------------------------------------------------------------------------------
iCurves = 2
sGname = "DOACV"
'-------------------------------------------------------------------------------Call Data.Root.Clear()
Call Data.Root.Clear()
Call Data.Root.ChannelGroups.Add("DIAdem_reference").Activate()
Call ChnLinGenImp("t1",263,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""DOACV-A1"")= 1000*sqrt(2)*cos(2000*pi*Ch(""t1""))*(1-180*Ch(""t1""))") 
'                                 1000*sqrt(2)*sin(2000*pi*     t    )*(1-180*     t      )
Call ChnLinGenImp("t2",763,-.01,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""DOACV-A2"")= abs(0<Ch(""t2"")) * 1000*sqrt(2)*sin(2000*pi*Ch(""t2"")+pi/2)*(1-180*Ch(""t2"")) + (1-abs(0<Ch(""t2""))) *1000*sqrt(2)")
'                                 H  (       t    ) *(1000*sqrt(2)*sin(2000*pi*     t    +pi/2)*(1-180*     t   )) + (1-H  (       t    )) *1000*sqrt(2)

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
Call eval2excel(6,1,iCurves,sGname,4) ' frequency  - this gives a wrong value because of DC part
' damping factor
                                           
'--- Fill results of DIAdem curves into Excel Workbook          
'do not call init                                               
Call eval2excel(2,2,iCurves,sGname,8) ' Up                    
Call eval2excel(6,2,iCurves,sGname,10) ' frequency  - this gives a wrong value because of DC part
' damping factor

Call datetime(iCurves,sGname,14) ' date & time  
'Call SaveAndCloseExcel        