'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2017-05-30 17:16:51
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

        Dim MyExcelFile
        Dim Excel
MyExcelFile = "C:\Users\chmamai\OneDrive for Business\IEC\TDG_validation_software\IEC61083_4.xlsx"
        Set Excel = CreateObject("Excel.Application")
        Excel.Visible = TRUE
        Call Excel.Workbooks.Open(
