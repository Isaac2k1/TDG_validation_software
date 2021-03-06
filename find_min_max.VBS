'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2017-05-22 14:33:56
'-- Author: Adrian Kress
'-- Comment: v1.1:  2017-05-22: Find max and min for each channel in each group
'--                             and call a Excelsheet with the shotname and the maxs and mins
'--                             for each channel
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

' checks if there are some shots
if (GroupCount < 1) then
    call msgbox("There aren't any shots in the DataPortal")
else 
    dim i, j
    
    'loops through each group in Data-Portal
    for i=1 to GroupCount
        
        'create or open a EXCEL-file with shotname
        Dim Excel
        Set Excel = CreateObject("Excel.Application")
        Excel.Visible = TRUE
        Call Excel.Workbooks.Add()
        Excel.ActiveWorkbook.Windows(1).Caption = GroupPropGet(i,"name")
        
        'add 'Min' and 'Max' column to the 
        Excel.ActiveSheet.Cells(1,3).Value = "Min"
        Excel.ActiveSheet.Cells(1,4).Value = "Max"
        
        'for each channel in group
        for j=1 to GroupChnCount(i)
          ' saves the name of the channel to a line in the file
          Excel.ActiveSheet.Cells(1+j,1).Value = ChnPropGet("[" & i & "]/[" & j & "]", "name")
          
          ' returns max and min and save it in the sheet
          Excel.ActiveSheet.Cells(1+j,3).Value = Str((ChnPropGet("[" & i & "]/[" & j & "]", "minimum")))
          Excel.ActiveSheet.Cells(1+j,4).Value = Str((ChnPropGet("[" & i & "]/[" & j & "]", "maximum")))
        next
        Excel.ActiveSheet.Parent.Saved = FALSE
        Set Excel = Nothing
    next
end if