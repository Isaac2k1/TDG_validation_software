'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 19.12.2012 22:40:25
'-- Ersteller: 
'-- Kommentar: 
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Script.
Call MsgBoxSetPos(10, 10, 30, 20)
Call MsgButtonTextSet("Step&1", "Step&2", "Step&3")
Select Case MsgBoxDisp("Select next step","MB_YESNOCANCEL")
  Case "IDYes"
    Call MsgBoxDisp("Step 1", "MB_OK", "MsgTypeNote",, 5)
  Case "IDNo"
    Call MsgBoxDisp("Step 2", "MB_OK", "MsgTypeInformation",, 5)
  Case  "IDCancel"
    Call MsgBoxDisp("Step 3", "MB_OK", "MsgTypeWarning",, 5)
End Select
MsgNotModal = TRUE
Call MsgBoxDisp("Example completed", "MB_NOBUTTON", "MsgTypeNote",, 10,MsgNotModal)
