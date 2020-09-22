Attribute VB_Name = "modVBKeyWords"
Option Explicit

' **********************************************************************
'
' Module:         modVBKeyWords
' Author:         Sharp Dressed Codes
' Web:            http://sharpdressedcodes.com
' Built:          20th January, 2008
' Purpose:        Parses Visual Basic Keywords in a RichTextBox Control
'                 to make them look like the VB IDE.
' Comments:
' Bugs:           Add constants & types needs finishing off, but it's fine for now.
'
' **********************************************************************


Public Const COLOR_VBCOMMENT As Long = 32768       ' green
Public Const COLOR_VBKEYWORD As Long = 8388608     ' blue
Public Const COLOR_VBTEXT As Long = 0              ' black

Public Type ExtrasType
  Value As String
  Type As DeclareTypes
End Type

Public VBKeyWords As Collection
Public Extras As Collection

Private StartIndex As Integer

Public Sub ParseVBKeyWords(ByVal rtb As RichTextBox, ByVal strData As String, ByVal boolForceLongs As Boolean, ByVal boolAddConstants As Boolean, Optional ByVal boolAppend As Boolean)
  
  Dim e As Byte ' 0 = none, 1 = prefix, 2 = suffix
  Dim I As Integer, J As Integer, Pos As Integer, Index As Long
  Dim arrLines() As String, arrWords() As String
  Dim strTemp As String, strKeyWord As String, strComment As String
  Dim boolFlag As Boolean, lst As ListBox, et As ExtrasType
  Dim xyz As String
  
  If Not boolAppend Then Set Extras = New Collection
  
  If LenB(strData) = 0 Then Exit Sub
  
  Do While InStr(strData, Space$(3))
    strData = Replace$(strData, Space$(3), Space$(2))
  Loop
  
  If VBKeyWords Is Nothing Then LoadVBKeyWords
  
  With rtb
    
    .Locked = False
    
    If Not boolAppend Then
      .Text = vbNullString
      .SelStart = 0
    End If
    
    Pos = InStr(strData, "= ")
    If Pos > 0 Then
      strTemp = Mid$(strData, Pos + 2)
      Pos = InStr(strTemp, Space$(1))
      If Pos > 0 Then strTemp = Left$(strTemp, Pos - 1)
      Pos = InStr(strTemp, vbCrLf)
      If Pos > 0 Then strTemp = Left$(strTemp, Pos - 1)
    End If
    
    If (boolForceLongs) And (Left$(LCase$(strData), 6) = "const ") And (InStr(LCase$(strData), " as ") = 0) Then
      If (IsNumeric(strTemp)) Or (Right(strTemp, 1) = "&") Then
        strData = Replace$(strData, " = ", " As Long = ")
      ElseIf Left$(strTemp, 1) = Chr$(34) Then
        strData = Replace$(strData, " = ", " As String = ")
      End If
    End If
    
    If InStr(strData, vbCrLf) = 0 Then
      
      strTemp = strData
      Pos = InStr(strTemp, "'")
      If Pos > 0 Then
        strComment = Mid$(strTemp, Pos)
        strTemp = Left$(strTemp, Pos - 1)
      End If
      
      If InStr(strTemp, Space$(1)) = 0 Then
        ReDim arrWords(0) As String
        arrWords(0) = strTemp
        GoSub PrintLine
      Else
        arrWords = Split(strTemp, Space$(1))
        GoSub PrintLine
      End If
      
    Else
    
      arrLines = Split(strData, vbCrLf)
      
      For I = 0 To UBound(arrLines)
        
        Pos = InStr(arrLines(I), "'")
        If Pos > 0 Then
          strComment = Mid$(arrLines(I), Pos)
          arrLines(I) = Left$(arrLines(I), Pos - 1)
        End If
        
        If InStr(arrLines(I), Space$(1)) = 0 Then
          ReDim arrWords(0) As String
          arrWords(0) = arrLines(I)
          GoSub PrintLine
        Else
          arrWords = Split(arrLines(I), Space$(1))
          GoSub PrintLine
        End If
        
      Next
      
    End If
    
    'If (boolAddConstants) And (boolFlag) And (Extras.Count) Then
    'If (boolAddConstants) And (boolFlag) And (Extras.Count) And (Not boolAppend) Then
    If (boolAddConstants) And (Extras.Count) And (Not boolAppend) Then
    'If (StartIndex > 0) And (boolAddConstants) And (Extras.Count) And (Not boolAppend) Then
      'For I = 1 To Extras.Count
      For I = StartIndex To Extras.Count
        boolFlag = False
        StartIndex = 0 '1
        ParseVBKeyWords rtb, Extras.Item(I), boolForceLongs, boolAddConstants, True
      Next
    ElseIf (boolAppend) And (StartIndex > 0) Then
      For I = StartIndex To Extras.Count
        boolFlag = False
        StartIndex = 0
        ParseVBKeyWords rtb, Extras.Item(I), boolForceLongs, boolAddConstants, True
      Next
    End If
    
    .SelColor = COLOR_VBTEXT  ' reset colour
    .Locked = True
    
  End With
  
  Exit Sub
  
PrintLine:
  
  With rtb
    
    'remove empty entries, so arrwords(0) is always the first word...
    
'    Dim colTemp As New Collection
'
'    For j = 0 To UBound(arrWords)
'      colTemp.Add arrWords(j)
'    Next
'
'    For j = colTemp.Count To 1 Step -1
'      If LenB(colTemp(j)) = 0 Then colTemp.Remove j
'    Next
'
'    ReDim arrWords(colTemp.Count - 1) As String
'
'    For j = 1 To colTemp.Count
'      arrWords(j - 1) = colTemp(j)
'    Next
    
    For J = 0 To UBound(arrWords)
      
      'If LenB(arrWords(j)) = 0 Then GoTo NextWord
      
      e = 0
      strKeyWord = vbNullString
      
      On Error Resume Next
      
      strKeyWord = VBKeyWords.Item(LCase$(arrWords(J)))
      If LenB(strKeyWord) = 0 Then
        strKeyWord = VBKeyWords.Item(Mid$(LCase$(arrWords(J)), 2))
        If LenB(strKeyWord) Then
          e = 1
        Else
          strKeyWord = VBKeyWords.Item(Left$(LCase$(arrWords(J)), Len(arrWords(J)) - 1))
          If LenB(strKeyWord) Then e = 2
        End If
      End If
      
      On Error GoTo 0
      
      'If (Not boolFlag) And (boolAddConstants) Then boolFlag = ((LCase$(strKeyWord) = "declare") Or (LCase$(strKeyWord) = "type"))
      If (Not boolFlag) And (boolAddConstants) Then
        boolFlag = (LCase$(strKeyWord) = "declare")
        If Not boolFlag Then boolFlag = (LCase$(strKeyWord) = "type")
      End If
      
      If J > 0 Then .SelText = Space$(1)
      
      Select Case e
        
        Case 0 ' none
          
          'If (boolAddConstants) And (boolFlag) Then
          If boolAddConstants Then
            If LenB(strKeyWord) = 0 Then
              'If j > 0 Then
              If J > 2 Then
            
                If LCase$(arrWords(J - 1)) = "as" Then
                  
                  'look for Types
                  
                  Index = InStrList(frmCom.lstType, "type " & arrWords(J))
                  If Index > -1 Then
                    Set lst = frmCom.lstType
                  Else
                    Index = InStrList(frmCom.lstType, "type " & Mid$(arrWords(J), 2))
                    If Index > -1 Then
                      Set lst = frmCom.lstType
                    Else
                      Index = InStrList(frmCom.lstType, "type " & Left$(arrWords(J), Len(arrWords(J)) - 1))
                      If Index > -1 Then
                        Set lst = frmCom.lstType
                      End If
                    End If
                  End If
                  
                  If lst Is Nothing Then
                    'look for enums
                    Index = InStrList(frmCom.lstEnum, "enum " & arrWords(J))
                    If Index > -1 Then
                      Set lst = frmCom.lstEnum
                    Else
                      Index = InStrList(frmCom.lstEnum, "enum " & Mid$(arrWords(J), 2))
                      If Index > -1 Then
                        Set lst = frmCom.lstEnum
                      Else
                        Index = InStrList(frmCom.lstEnum, "enum " & Left$(arrWords(J), Len(arrWords(J)) - 1))
                        If Index > -1 Then
                          Set lst = frmCom.lstEnum
                        End If
                      End If
                    End If
                  End If
                  
                  If lst Is Nothing Then
                    ' look for constants - could be in brackets ()
                    
                    'Index = FindMatch(frmCom.lstConst, "const " & arrWords(j))
                    Index = InStrList(frmCom.lstConst, "const " & arrWords(J))
                    If Index > -1 Then
                      Set lst = frmCom.lstConst
                    Else
                      Index = InStrList(frmCom.lstConst, "const " & Mid$(arrWords(J), 2))
                      If Index > -1 Then
                        Set lst = frmCom.lstConst
                      Else
                        Index = InStrList(frmCom.lstConst, "const " & Left$(arrWords(J), Len(arrWords(J)) - 1))
                        If Index > -1 Then
                          Set lst = frmCom.lstConst
                        'Else
                          'Index = FindMatch(frmCom.lstConst, "const " & Mid$(Left$(arrWords(j), Len(arrWords(j)) - 1), 2))
                          'If Index > -1 Then Set lst = frmCom.lstConst
                        End If
                      End If
                    End If
                  End If
                  
                End If 'If LCase$(arrWords(j - 1)) = "as" Then
                
              'Else 'If (j > 0) Then
              ElseIf J = 2 Then
                
                ' Privilege(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
                ' Privilege(ANYSIZE_ARRAY-1)
                ' Privilege(ANYSIZE_ARRAY+1)
                ' Privilege(ANYSIZE_ARRAY or)
                ' Privilege(ANYSIZE_ARRAY to )

                Pos = InStr(arrWords(J), "(")
                If Pos Then xyz = Mid$(arrWords(J), Pos + 1)
                If Pos Then Pos = InStr(xyz, ")")
                If (LenB(xyz)) And (Pos > 0) Then xyz = Left$(xyz, Pos - 1)
                
                'Debug.Print "xyz: " & xyz
                
                If LenB(xyz) Then
                  Index = InStrList(frmCom.lstConst, "const " & xyz)
                  If Index > -1 Then
                    Set lst = frmCom.lstConst
                  Else
                    Index = InStrList(frmCom.lstEnum, "enum " & xyz)
                    If Index > -1 Then Set lst = frmCom.lstEnum
                  End If
                End If
                
                xyz = vbNullString
                
              End If 'If (j > 0) Then
            
            End If 'lenb(keyword)=0
            
            Dim ii As Long
            If Not lst Is Nothing Then
'              If Extras.Count Then
'                For ii = 1 To Extras.Count
'                  If Extras.Item(ii) = lst.List(Index) Then
'                    Debug.Print lst.List(Index); " already here"
'                    GoTo AlreadyHere
'                  End If
'                Next
'              End If
              On Error Resume Next
              xyz = vbNullString
              xyz = Extras.Item(lst.List(Index))
              On Error GoTo 0
              If LenB(xyz) Then
                Debug.Print lst.List(Index); " already here"
               'GoTo AlreadyHere
               'Return
               'Set lst = Nothing
               'GoTo NextWord
              End If
              Err.Clear
              On Error Resume Next
              Extras.Add lst.List(Index), lst.List(Index)
              'If Err Then Extras.Remove Extras.Count
              If Err Then
                GoTo AlreadyHere
              End If
              On Error GoTo 0
              If Not boolAppend Then
                StartIndex = 1
              Else
                'If StartIndex = 1 Then StartIndex = Extras.Count
                If StartIndex < 1 Then StartIndex = Extras.Count
              End If
              
AlreadyHere:

              Set lst = Nothing
            End If
           
          End If 'booladdconstants
                
          .SelColor = IIf(LenB(strKeyWord), COLOR_VBKEYWORD, COLOR_VBTEXT)
          '.SelText = UCase$(Left$(arrWords(J), 1)) & Mid$(arrWords(J), 2)
          .SelText = arrWords(J)
        
        Case 1 ' prefix
          
          .SelColor = COLOR_VBTEXT
          .SelText = Left$(arrWords(J), 1)
          .SelColor = COLOR_VBKEYWORD
          '.SelText = UCase$(Mid$(arrWords(J), 2, 1)) & Mid$(arrWords(J), 3)
          .SelText = Mid$(arrWords(J), 2)
        
        Case 2 ' suffix
        
          .SelColor = COLOR_VBKEYWORD
          .SelText = UCase$(Left$(arrWords(J), 1)) & Mid$(arrWords(J), 2, Len(arrWords(J)) - 2) 'Left$(arrWords(j), Len(arrWords(j)) - 1)
          .SelColor = COLOR_VBTEXT
          .SelText = Right$(arrWords(J), 1)
          
      End Select

NextWord:

    Next
    
    If LenB(strComment) Then
      .SelColor = COLOR_VBCOMMENT
      .SelText = strComment
      strComment = vbNullString
    End If
      
    .SelText = vbCrLf
    
  End With
  
  Return
  
End Sub

Private Sub LoadVBKeyWords()
  
  ' this needs more...
  
  Set VBKeyWords = New Collection
  
  With VBKeyWords
    .Add "option", "option"
    .Add "explicit", "explicit"
    .Add "compare", "compare"
    .Add "base", "base"
    .Add "text", "text"
    .Add "database", "database"
    .Add "dim", "dim"
    .Add "global", "global"
    .Add "if", "if"
    .Add "then", "then"
    .Add "#if", "#if"
    .Add "end", "end"
    .Add "#end", "#end"
    .Add "select", "select"
    .Add "case", "case"
    .Add "else", "else"
    .Add "elseif", "elseif"
    .Add "do", "do"
    .Add "exit", "exit"
    .Add "loop", "loop"
    .Add "until", "until"
    .Add "while", "while"
    .Add "wend", "wend"
    .Add "open", "open"
    .Add "close", "close"
    .Add "input", "input"
    .Add "binary", "binary"
    .Add "access", "access"
    .Add "read", "read"
    .Add "write", "write"
    .Add "raiseevent", "raiseevent"
    .Add "for", "for"
    .Add "to", "to"
    .Add "next", "next"
    .Add "lbound", "lbound"
    .Add "ubound", "ubound"
    .Add "set", "set"
    .Add "get", "get"
    .Add "let", "let"
    .Add "nothing", "nothing"
    .Add "null", "null"
    .Add "empty", "empty"
    .Add "friend", "friend"
    .Add "private", "private"
    .Add "public", "public"
    .Add "const", "const"
    .Add "type", "type"
    .Add "enum", "enum"
    .Add "declare", "declare"
    .Add "withevents", "withevents"
    .Add "with", "with"
    .Add "function", "function"
    .Add "sub", "sub"
    .Add "property", "property"
    .Add "optional", "optional"
    .Add "byval", "byval"
    .Add "byref", "byref"
    .Add "as", "as"
    .Add "any", "any"
    .Add "variant", "variant"
    .Add "long", "long"
    .Add "integer", "integer"
    .Add "byte", "byte"
    .Add "boolean", "boolean"
    .Add "string", "string"
    .Add "true", "true"
    .Add "false", "false"
    .Add "cstr", "cstr"
    .Add "clng", "clng"
    .Add "cint", "cint"
    .Add "cvar", "cvar"
    .Add "ccur", "ccur"
    .Add "csng", "csng"
    .Add "cbyte", "cbyte"
    .Add "cbool", "cbool"
    .Add "cdate", "cdate"
    .Add "cdbl", "cdbl"
    .Add "cdec", "cdec"
    .Add "lib", "lib"
    .Add "alias", "alias"
    .Add "on", "on"
    .Add "error", "error"
    .Add "goto", "goto"
    .Add "resume", "resume"
    .Add "gosub", "gosub"
    .Add "redim", "redim"
    .Add "preserve", "preserve"
    .Add "implements", "implements"
    .Add "static", "static"
    .Add "put", "put"
    .Add "line", "line"
    .Add "print", "print"
    .Add "is", "is"
    .Add "typeof", "typeof"
    .Add "foreach", "foreach"
    .Add "new", "new"
    .Add "debug", "debug"
    .Add "assert", "assert"
  End With
  
End Sub
