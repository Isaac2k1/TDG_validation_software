'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-14 12:27:23
'-- Author:  MM
'-- Comment: IEC 61083-4 TDG validation
'-- /Users/chmamai/OneDrive - ABB/IEC/TDG_validation_software/SHACI.VBS
'-------------------------------------------------------------------------------
'SHACI-A1...7
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude("Function_Library")
'-------------------------------------------------------------------------------
iCurves = 7
sGname = "SHACI"
'-------------------------------------------------------------------------------
Call Data.Root.Clear()
Call Data.Root.ChannelGroups.Add("DIAdem_reference").Activate()

sONOFF  ="abs(M<=ch(""t""))*abs(ch(""t"")<N)"
sEXPamp ="((B-A)*exp((M-ch(""t""))/P) + (A-C)*exp((M-ch(""t""))/O) + C)"
sFREQch ="sin(2*pi*(ch(""t"") - M)*(G + (ch(""t"") - M)/S*D) - F)"
sDCpart =" B *sin(F)*exp((M -1 *ch(""t""))/Q)"

sFormula1 = sONOFF & "*" & "(" & sEXPamp & "*" & sFREQch & "+" & sDCpart & ") + V"


'Call ChnLinGenImp("t1",11501,0.0,2E-5,"s") ' Generates x channel
for iCurve = 1 to iCurves

    call setshaci(iCurve,sSymbols,sValues,sUnit,Pts,t0,dt)
    
    call ChnLinGenimp("t"&icurve,pts,t0,dt,"s") ' generates x channel
    call ChnLinGenimp("t",Pts,t0,dt,"s") ' generates another copy of x channel, just for the calculate command
    'sFormula = "ch(""shaci-aa"")= abs(tt1 <ch(""t1""))* abs(ch(""t1"")<0.215042) * (((i2k-i1k)*exp((tt1-ch(""t1""))/t2ac) + (i1k-id)*exp((tt1-ch(""t1""))/t1ac) + id) * sin(2*pi*(ch(""t1"")- tt1)*(f0+(ch(""t1"")-tt1)/ttot*df)-phi) + i2k*sin(phi)*exp(-1*ch(""t1"")       /  tdc)) + ioff "
    'sFormula = "ch(""shaci-ab"")=(abs(0.02<ch(""t1""))*(abs(ch(""t1"")<0.215042))*    (1e3*sin(2*pi*(ch(""t1"")-0.02)* 50                        - pi/2) + 1e3*         exp((-1*ch(""t1"")+0.02)/45e-3)))"
    'define sFormula for the calculate command
    sFormula = "ch(""SHACI-A" & icurve & """)=" & sFormula1
'    sFormula = sFormula & sFormula1
'    sFormula = sFormula & "abs(M<=ch(""t""))*abs(ch(""t"")<N)"                            'on/off part
'    sFormula = sFormula & " * ("
'    sFormula = sFormula & "((B-A)*exp((M-ch(""t""))/P) + (A-C)*exp((M-ch(""t""))/O) + C)" 'exponential amplitude change
'    sFormula = sFormula & " * "
'    sFormula = sFormula & "sin(2*pi*(ch(""t"") - M)*(G + (ch(""t"") - M)/S*D) - F)"       'frequency changing
'    sFormula = sFormula & " + "
'    sFormula = sFormula & " B *sin(F)*exp((M -1 *ch(""t""))/Q)"                           'decaying dc offset 
'    sFormula = sFormula & ")"
'    sFormula = sFormula & "+V"                                                            'offset

' msgbox sFormula
    call    calculate(sFormula,ssymbols,svalues,sunit) 
    'the x channel is now no longer needed
    call  chndelete("t")
next 

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
Call eval2excel(4,1,iCurves,sGname,4) ' Urmsac                     
'Call eval2excel(4,1,iCurves,sGname,6) ' Urmsac                   
'Call eval2excel(5,1,iCurves,sGname,8) ' THD                      
'Call eval2excel(6,1,iCurves,sGname,10) ' frequency                
                                           
'--- Fill results of DIAdem curves into Excel Workbook          
'do not call init                                               
Call eval2excel(2,2,iCurves,sGname,12) ' Up                    
Call eval2excel(4,2,iCurves,sGname,14) ' Urmsac                 
'Call eval2excel(4,2,iCurves,sGname,16) ' Urmsac                
'Call eval2excel(5,2,iCurves,sGname,18) ' THD                   
'Call eval2excel(6,2,iCurves,sGname,20) ' frequency             
Call datetime(iCurves,sGname,22) ' date & time  
'Call SaveAndCloseExcel                                          