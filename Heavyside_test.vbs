'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-26 23:24:35
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
dim Temp, TempC, t
Call Data.Root.Clear()
'Function H(t)
  'if t > 0 then H(t)=1 else H(t)=0
  Call Calculate ("Ch(""H"")= Sign (0.04-Ch(""time""))")
  Call Calculate ("Ch(""G"")= Sign (Ch(""time"")-0.04)")
'End function

