'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-14 12:27:23
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
'  ACI-A1...5
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude("Function_Library")
'-------------------------------------------------------------------------------
iCurves = 5
sGname = "ACI"
'-------------------------------------------------------------------------------
Call Data.Root.Clear()
Call Data.Root.ChannelGroups.Add("DIAdem_reference").Activate()
Call ChnLinGenImp("t1",7501,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACI-A1"")= 1000*sqrt(2)*sin(2*50*pi*Ch(""t1""))") ' Calculates sinus of x channel
Call ChnLinGenImp("t2",6251,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACI-A2"")= 1000*sqrt(2)*sin(2*60*pi*Ch(""t2"")) + 500")
Call ChnLinGenImp("t3",7501,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACI-A3"")= 1000*sqrt(2)*(sin(2*50*pi*Ch(""t3"")) + 0.05*sin(8*50*pi*Ch(""t3"")+pi*2/3))") 
Call ChnLinGenImp("t4",6251,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACI-A4"")= 1000*sqrt(2)*(sin(2*60*pi*Ch(""t4"")) + 0.07*sin(2*60*pi*Ch(""t4"")*5+pi*22/18))") 
Call ChnLinGenImp("t5",7501,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACI-A5"")= 1000*sqrt(2)*(sin(2*50*pi*Ch(""t5"")) + 0.02*sin(100*pi*Ch(""t5"")*3-pi) + 0.02*sin(100*pi*Ch(""t5"")*5)+0.0015*sin(100*pi*Ch(""t5"")*7-pi)+0.009*sin(100*pi*Ch(""t5"")*9))") 

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

'--- Fill results of reference curves of TDG into Excel Workbook
'--- Use sub eval2excel(iCalc,iRef,iCurves,sGname,sColumn)
Call eval2excel(1,1,iCurves,sGname,1) ' init
Call eval2excel(7,1,iCurves,sGname,2) ' Iss
Call eval2excel(3,1,iCurves,sGname,4) ' Urms
Call eval2excel(5,1,iCurves,sGname,6) ' THD
Call eval2excel(6,1,iCurves,sGname,8) ' frequency

'--- Fill results of DIAdem curves into Excel Workbook
'do not call init
Call eval2excel(7,2,iCurves,sGname,10) ' Iss
Call eval2excel(3,2,iCurves,sGname,12) ' Urms
Call eval2excel(5,2,iCurves,sGname,14) ' THD
Call eval2excel(6,2,iCurves,sGname,16) ' frequency
Call datetime(iCurves,sGname,18) ' date & time
Call SaveAndCloseExcel

'Call Data.Root.ChannelGroups(2).Activate()
'QRangeMin = -1000 * sqrt(2)
'QRangeMax =  1000 * sqrt(2)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_5bit",32,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_6bit",64,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_7bit",128,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_8bit",256,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_9bit",512,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_10bit",1024,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_11bit",2048,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_12bit",4096,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_13bit",8192,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_14bit",16384,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_15bit",32768,QRangeMin, QRangeMax)
'Call ChnQuantize("[1]/ACI-A1","/ACI-A1_16bit",65536,QRangeMin, QRangeMax)
