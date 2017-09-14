'------------------------------------------------------------------------------
'The DIAdem 11.3 DataPlugin Wizard generated this script on 10/14/2011 13:14.
'
' This DataPlugin can be used to import ASCII output files of the TDG v0.3 pro-
' gram into DIAdem. It is adapted for the STL-header only. Other header formats
' were not available at the time of writing. 
'
' Author: Marco Mailand
' e-mail: marco.mailand@ch.abb.com
' Comp. : ABB Switzerland Ltd. 
'
'------------------------------------------------------------------------------
Option Explicit
Const eIgnore=101
Const eChannel=1
Const eChannelGroup=2

Sub ReadStore(File)
  Dim AreaTypes
  Dim AreaLineCounts
  Dim DataBeginLine
  Dim GroupPropNames
  Dim GroupPropDataTypes
  Dim GroupPropValues
  Dim GroupName
  Dim ChannelPropNames
  Dim ChannelPropDataTypes
  Dim ChannelPropValues
  Dim ChannelNames
  Dim Group
  Dim ChannelTypes
  Dim Channels
  '----------------------------------------------------
  Call SetLineFormatter(File)
  '----------------------------------------------------
  AreaTypes      = Array(eChannelGroup,eChannel,eChannelGroup)
  AreaLineCounts = Array(7,3,2)
  DataBeginLine  = GetDataBeginLine(File,AreaTypes,AreaLineCounts)
  '----------------------------------------------------
  GroupPropNames     = Array("header rows","Delimiter","Data rows","creation date","creation time","creation source","comments","resolution bits","dynamic range %")
  GroupPropDataTypes = Array(eString,eString,eString,eString,eString,eString,eString,eString,eString)
  GroupPropValues = GetGroupPropValues(File,AreaTypes,AreaLineCounts,GroupName,GroupPropNames,GroupPropDataTypes,NULL,1)
  '----------------------------------------------------
  ChannelPropNames     = Array("Scale factor","unit_string","name")
  ChannelPropDataTypes = Array(eString,eString,eString)
  ChannelPropValues = GetChannelPropValues(File,AreaTypes,AreaLineCounts,ChannelNames,ChannelPropNames,ChannelPropDataTypes,NULL,1)
  '----------------------------------------------------
  Set Group = CreateGroup(File,Root,GroupName,GroupPropNames,GroupPropValues)
  '----------------------------------------------------
  ChannelTypes = Array(eR64,eR64)
  Set Channels = CreateChannels(File,DataBeginLine,Group,ChannelNames,ChannelTypes,ChannelPropNames,ChannelPropValues,False)
End Sub

'----------------------------------------------------
Sub SetLineFormatter(File)
  File.Formatter.LineFeeds = GetLineFeed(File)
  File.Formatter.CommentSign = ""
  File.Formatter.IgnoreEmptyLines = True
End Sub
'----------------------------------------------------
Sub SetGroupPropFormatter(File)
  File.Formatter.Delimiters = ";"
  File.Formatter.DecimalPoint = "."
  File.Formatter.ThousandSeparator = ""
  File.Formatter.ExponentSeparator = "e"
  File.Formatter.StringSign = ""
  File.Formatter.TimeFormat = "D.M.YY hh:mm:ss.ffffff"
  File.Formatter.TrimCharacters = ""
  File.Formatter.NoValueSign = ""
End Sub
'----------------------------------------------------
Sub SetChannelPropFormatter(File)
  File.Formatter.Delimiters = ";"
  File.Formatter.DecimalPoint = "."
  File.Formatter.ThousandSeparator = ""
  File.Formatter.ExponentSeparator = "e"
  File.Formatter.StringSign = ""
  File.Formatter.TimeFormat = "D.M.YY hh:mm:ss.ffffff"
  File.Formatter.TrimCharacters = ""
  File.Formatter.NoValueSign = ""
End sub
'----------------------------------------------------
Sub SetChannelValueBlockFormatter(File)
  File.Formatter.Delimiters = ";"
  File.Formatter.DecimalPoint = "."
  File.Formatter.ThousandSeparator = ""
  File.Formatter.ExponentSeparator = "e"
  File.Formatter.StringSign = ""
  File.Formatter.TimeFormat = "D.M.YY hh:mm:ss.ffffff"
  File.Formatter.TrimCharacters = ""
  File.Formatter.NoValueSign = ""
End Sub

'----------------------------------------------------
Function GetLineFeed(File)
  Dim BuffSize
  Dim CurrPos
  Dim sgBuffer
  Dim CharSize
  Dim CharCount
  If (File.Formatter.CharacterFormat = eUTF16) Then
    CharSize = 2
  Else
    CharSize = 1
  End If
  CharCount = File.Size/CharSize
  BuffSize = 200
  If (BuffSize > CharCount) Then BuffSize = File.Size   
  CurrPos = File.Position
  sgBuffer = File.GetCharacters(BuffSize)
  If (InStr(sgBuffer,vbCRLF) > 0) Then   ' Check for CR+LF
    GetLineFeed = vbCRLF
  ElseIf (InStr(sgBuffer,vbCR) > 0) Then ' Check for CR
    GetLineFeed = vbCR
  ElseIf (InStr(sgBuffer,vbLF) > 0) Then ' Check for LF
    GetLineFeed = vbLF
  Else
    GetLineFeed = vbNewLine
  End If  
  File.Position = CurrPos
End Function
'----------------------------------------------------
Function GetValue(ValueArray,Index,ValueDefault)
  If Index<=UBound(ValueArray) Then
    If IsObject(ValueArray(Index)) Then
      Set GetValue = ValueArray(Index)
    Else
      GetValue = ValueArray(Index)
    End If
  Else
    GetValue = ValueDefault
  End If
End Function
'----------------------------------------------------
Function GetDataBeginLine(File,AreaTypes,ByRef AreaLineCounts)
  Dim Index
  Dim DataBeginLine
  Dim AreaLineCount
  Dim LineText
  Dim CompareText
  Dim SkipOffset
  Dim CurrPos

  CurrPos = File.Position
  SkipOffset = 0
  DataBeginLine = 1
  If UBound(AreaLineCounts)>=0 Then
    For Index=0 To UBound(AreaLineCounts)
      If VarType(AreaLineCounts(Index))=vbInteger Then
        Call File.SkipLines(AreaLineCounts(Index)-SkipOffset)
        SkipOffset = 0
      Else
        AreaLineCount = 0
        CompareText = LCase(AreaLineCounts(Index))
        LineText = LCase(File.GetNextLine)
        While (LineText<>CompareText) And (File.Position<File.Size)
          AreaLineCount = AreaLineCount+1
          LineText = LCase(File.GetNextLine)
        Wend
        AreaLineCounts(Index) = AreaLineCount
        SkipOffset = 1
      End If
      DataBeginLine = DataBeginLine + AreaLineCounts(Index)
    Next
  End If
  File.Position = CurrPos
  GetDataBeginLine = DataBeginLine
End Function
'----------------------------------------------------
Function GetGroupPropValues(File,AreaTypes,AreaLineCounts,ByRef GroupName,ByRef GroupPropNames,GroupPropDataTypes,PropNameColumn,PropValueColumn)
  Const PropNameDefault = "Noname"
  Dim PropDataTypeDefault : PropDataTypeDefault = eString
  Dim GroupPropValues()
  Dim GroupPropNamesEx()
  Dim AreaIndex
  Dim PropIndex
  Dim LineIndex
  Dim PropName
  Dim PropDataType
  Dim PropValueString
  Dim lFilePositionOld
  Dim CurrPos

  CurrPos = File.Position
  Call SetGroupPropFormatter(File)

  GroupName = File.Info.FileName
  ReDim GroupPropNamesEx(-1)
  ReDim GroupPropValues(-1)
  PropIndex = 0
  AreaIndex = 0
  While AreaIndex<=UBound(AreaTypes)
    If AreaTypes(AreaIndex)=eChannelGroup Then
      If AreaLineCounts(AreaIndex)>0 Then
        Redim Preserve GroupPropNamesEx(PropIndex+AreaLineCounts(AreaIndex)-1)
        Redim Preserve GroupPropValues(PropIndex+AreaLineCounts(AreaIndex)-1)
        For LineIndex=1 To AreaLineCounts(AreaIndex)
          lFilePositionOld = File.Position
          If IsNull(PropNameColumn) Then
            PropName = GetValue(GroupPropNames,PropIndex,PropNameDefault&"_"&(PropIndex-UBound(GroupPropNames)))
            PropDataType = GetValue(GroupPropDataTypes,PropIndex,PropDataTypeDefault)
            Call File.SkipValues(PropValueColumn-1)
          Else
            Call File.SkipValues(PropNameColumn-1)
            PropName = File.GetNextStringValue(eString)
            Call File.SkipValues(PropValueColumn-PropNameColumn-1)
            PropValueString = File.GetNextStringValue(eString)
            If IsNumeric(PropValueString) Then
              PropDataType = eR64
            ElseIf IsDate(PropValueString) Then
              PropDataType = eTime
            Else
              PropDataType = eString
            End If
            File.Position = lFilePositionOld
            Call File.SkipValues(PropValueColumn-1)
          End If
          GroupPropNamesEx(PropIndex) = PropName
          if LCase(PropName)="name" Then
            GroupPropValues(PropIndex) = File.GetNextStringValue(eString)
            GroupName = GroupPropValues(PropIndex)
          ElseIf PropDataType=eTime Then
            On Error Resume Next
            Set GroupPropValues(PropIndex) = File.GetNextStringValue(PropDataType)
            On Error Goto 0
          Else
            GroupPropValues(PropIndex) = File.GetNextStringValue(PropDataType)
          End If
          File.Position = lFilePositionOld
          Call File.GetNextLine
          PropIndex = PropIndex+1
        Next
      End If
    Else
      Call File.SkipLines(AreaLineCounts(AreaIndex))
    End If
    AreaIndex= AreaIndex+1
  Wend
  File.Position = CurrPos

  GroupPropNames = GroupPropNamesEx
  GetGroupPropValues = GroupPropValues
End Function
'----------------------------------------------------
Function GetChannelPropValues(File,AreaTypes,AreaLineCounts,ByRef ChannelNames,ChannelPropNames,ChannelPropDataTypes,PropNameColumn,PropValueStartColumn)
  Const PropNameDefault = "Noname"
  Dim PropDataTypeDefault : PropDataTypeDefault = eString
  Dim ChannelPropValues()
  Dim ChannelPropNamesEx()
  Dim AreaIndex
  Dim PropIndex
  Dim ValueIndex
  Dim LineIndex
  Dim PropName
  Dim PropDataType
  Dim PropValue
  Dim PropValues()
  Dim PropValueString
  Dim lFilePositionOld
  Dim CurrPos

  CurrPos = File.Position
  Call SetChannelPropFormatter(File)

  ChannelNames = Array()
  ReDim ChannelPropNamesEx(-1)
  ReDim ChannelPropValues(-1)
  PropIndex = 0
  AreaIndex = 0
  While AreaIndex<=UBound(AreaTypes)
    If AreaTypes(AreaIndex)=eChannel Then
      If AreaLineCounts(AreaIndex)>0 Then
        Redim Preserve ChannelPropNamesEx(PropIndex+AreaLineCounts(AreaIndex)-1)
        Redim Preserve ChannelPropValues(PropIndex+AreaLineCounts(AreaIndex)-1)
        For LineIndex=1 To AreaLineCounts(AreaIndex)
          lFilePositionOld = File.Position
          If IsNull(PropNameColumn) Then
            PropName = GetValue(ChannelPropNames,PropIndex,PropNameDefault&"_"&(PropIndex-UBound(ChannelPropNames)))
            PropDataType = GetValue(ChannelPropDataTypes,PropIndex,PropDataTypeDefault)
            Call File.SkipValues(PropValueStartColumn-1)
          Else
            Call File.SkipValues(PropNameColumn-1)
            PropName = File.GetNextStringValue(eString)
            Call File.SkipValues(PropValueStartColumn-PropNameColumn-1)
            PropValueString = File.GetNextStringValue(eString)
            If IsNumeric(PropValueString) Then
              PropDataType = eR64
            ElseIf IsDate(PropValueString) Then
              PropDataType = eTime
            Else
              PropDataType = eString
            End If
            File.Position = lFilePositionOld
            Call File.SkipValues(PropValueStartColumn-1)
          End If
          If LCase(PropName)="name" Then
            PropDataType = eString
          End If
          If PropDataType=eTime Then
            ValueIndex = 0
            On Error Resume Next
            Set PropValue = File.GetNextStringValue(PropDataType)
            While (Not IsEmpty(PropValue)) And (Err.Number=0)
              Redim Preserve PropValues(ValueIndex)
              Set PropValues(ValueIndex) = PropValue
              ValueIndex = ValueIndex+1
              Set PropValue = File.GetNextStringValue(PropDataType)
            Wend
            On Error Goto 0
          Else
            ValueIndex = 0
            PropValue = File.GetNextStringValue(PropDataType)
            While Not IsEmpty(PropValue)
              Redim Preserve PropValues(ValueIndex)
              PropValues(ValueIndex) = PropValue
              ValueIndex = ValueIndex+1
              PropValue = File.GetNextStringValue(PropDataType)
            Wend
          End If
          ChannelPropNamesEx(PropIndex) = PropName
          ChannelPropValues(PropIndex) = PropValues
          If LCase(PropName)="name" Then
            ChannelNames = PropValues
          End If
          File.Position = lFilePositionOld
          Call File.GetNextLine
          PropIndex = PropIndex+1
        Next
      End If
    Else
      Call File.SkipLines(AreaLineCounts(AreaIndex))
    End If
    AreaIndex= AreaIndex+1
  Wend
  File.Position = CurrPos

  ChannelPropNames = ChannelPropNamesEx
  GetChannelPropValues = ChannelPropValues
End Function
'----------------------------------------------------
Function CreateGroup(File,Root,GroupName,GroupPropNames,GroupPropValues)
  Dim Group
  Dim Index

  Set Group = Root.ChannelGroups.Add(GroupName)
  Index=0
  While Index<=UBound(GroupPropNames)
    If LCase(GroupPropNames(Index))<>"name" Then
      On Error Resume Next
      Call Group.Properties.Add(GroupPropNames(Index),GroupPropValues(Index))
      If Err Then
        On Error Goto 0
        Call Group.Properties.Add(GroupPropNames(Index)&"1",GroupPropValues(Index))
      End If
      On Error Goto 0
    End If
    Index = Index+1
  Wend
  Set CreateGroup = Group
End Function
'----------------------------------------------------
Function CreateChannels(File,DataBeginLine,Group,ChannelNames,ChannelTypes,ChannelPropNames,ChannelPropValues,Waveform)
  Dim ChannelDataTypeDefault : ChannelDataTypeDefault = eR64
  Const ChannelNameDefault = "Noname"
  Dim Index
  Dim ChannelValueBlock
  Dim BlockIndex
  Dim ChannelCount
  Dim TempValue
  Dim ChannelName
  Dim ChannelDataType
  Dim Channel

  Call SetChannelValueBlockFormatter(File)
  Call File.SkipLines(DataBeginLine-1)
  Set ChannelValueBlock = File.GetStringBlock()

  BlockIndex = 0
  ChannelCount = 0
  TempValue = File.GetNextStringValue(eString)
  While Not IsEmpty(TempValue)
    ChannelDataType = GetValue(ChannelTypes,BlockIndex,ChannelDataTypeDefault)

    If IsNull(ChannelDataType) Then
      Set Channel = ChannelValueBlock.Channels.Add("temp", eString)
    ElseIf ChannelDataType=eIgnore Then
      Set Channel = ChannelValueBlock.Channels.Add("temp", eString)
      ChannelCount = ChannelCount+1
    Else
      ChannelName = GetValue(ChannelNames,ChannelCount,ChannelNameDefault)
      Set Channel = ChannelValueBlock.Channels.Add(ChannelName,ChannelDataType)
      Call Group.Channels.AddDirectAccessChannel(Channel)

      If Waveform Then
        Call Channel.Properties.Add("wf_start_offset",CDbl(1))
        Call Channel.Properties.Add("wf_increment",CDbl(1))
        Call Channel.Properties.Add("wf_xname","")
        Call Channel.Properties.Add("wf_xunit_string","")
        Call Channel.Properties.Add("wf_samples",CLng(1))
        Call Channel.Properties.Add("wf_time_pref","relative")
        Call Channel.Properties.Add("wf_start_time",CreateTime(0,1,1,0,0,0,0,0,0))
      End If
      If UBound(ChannelPropNames)>=0 Then
        For Index=0 To UBound(ChannelPropNames)
          If LCase(ChannelPropNames(Index))<>"name" Then
            On Error Resume Next
            Call Channel.Properties.Add(ChannelPropNames(Index),GetValue(ChannelPropValues(Index),ChannelCount,""))
            If Err Then
              On Error Goto 0
              Call Channel.Properties.Add(ChannelPropNames(Index)&"1",GetValue(ChannelPropValues(Index),ChannelCount,""))
            End If
            On Error Goto 0
          End If
        Next
      End If
      ChannelCount = ChannelCount+1
    End If
    
    BlockIndex=BlockIndex+1
    TempValue = File.GetNextStringValue(eString)    
  Wend 

  Set CreateChannels = Group.Channels
End Function
