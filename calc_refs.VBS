'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-14 12:27:23
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
'ACV-A1...6
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
dim t,i,maxindex, DIAdem_reference, MyFolders(), QRangeMin, QRangeMax, sCurveName, sFileName,sNumber, iIndex,iAmplidude
maxindex= (data.Root.ChannelGroups.Count)+1
Call Data.Root.ChannelGroups.Add("DIAdem_reference").Activate()
Call ChnLinGenImp("time",7501,0,2E-5,"s") ' Generates x channel
Call Calculate("Ch(""ACV-A1"")= 1000*sqrt(2)*sin(2*50*pi*Ch(""time""))") ' Calculates sinus of x channel
Call Calculate("Ch(""ACV-A2"")= 1000*sqrt(2)*sin(2*50*pi*Ch(""time"")) + 500") 
Call Calculate("Ch(""ACV-A3"")= 1000*sqrt(2)*(sin(2*50*pi*Ch(""time"")) + 0.05*sin(8*50*pi*Ch(""time"")+pi*2/3))") 
Call Calculate("Ch(""ACV-A4"")= 1000*sqrt(2)*(sin(2*50*pi*Ch(""time"")) + 0.07*sin(2*50*pi*Ch(""time"")*5+pi*22/18))") 
Call Calculate("Ch(""ACV-A5"")= 1000*sqrt(2)*(sin(2*50*pi*Ch(""time"")) + 0.05*sin(100*pi*Ch(""time"")*3-pi) + 0.05*sin(100*pi*Ch(""time"")*5)+0.02*sin(100*pi*Ch(""time"")*7-pi)+0.01*sin(100*pi*Ch(""time"")*9))") 
Call Calculate("Ch(""ACV-A6"")= 1000*sqrt(2)*(sin(2*50*pi*Ch(""time"")) + 0.02*sin(100*pi*Ch(""time"")*3-pi) + 0.02*sin(100*pi*Ch(""time"")*5)+0.0015*sin(100*pi*Ch(""time"")*7-pi)+0.009*sin(100*pi*Ch(""time"")*9))") 
