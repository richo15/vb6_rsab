VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'=====
'JNode
'=====
'
'LICENSE:
'
'JNode Class (JNode.cls)
'
'Version 1.2
'
'A parser/serializer/DOM class for JSON data interchange written
'in Visual Basic 6.0 (usable in Office VBA with little or no
'modification).
'
'
'Copyright 2015 Robert D. Riemersma, Jr.
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'
'NOTES:
'
'   o No emphasis on performance.
'
'   o Error checking is reasonbly thorough, but most errors will
'     raise a simple "error 5" with no extra information.
'
'   o JNode stores/returns Empty as the JSON null value.
'
'   o JNode stores/returns Double for all numeric values.
'
'   o Val() and Str$() are used to ensure "." is used as decimal point
'     no matter what the current language and regional settings are.
'
'   o The first Value assignment or retrieval of an member sets the
'     JSON type to "object" or "array" based on the ID passed (String
'     implies a Key for an object, Integer or Long implies an Index
'     for an array).
'
'   o Yes, attempting to retrieve a non-existant Value item causes it
'     to be created with the value Empty.
'
'   o Alternatively the MakeObject or MakeArray methods can be called
'     to set which variety of JNode we are.  However these calls will
'     discard any current contents of the JNode.
'
'   o Remove method only supports removing members of JSON objects.
'     This is more commonly required, but if removal from arrays is
'     also needed that could be added.
'
'   o Requires a reference to:
'
'           Microsoft Scripting Runtime (for Dictionary).
'
'Version 1.1 changes:
'
'   o Added encoding for LF, CR, TAB, BS, FF, other ASCII control
'     characters, characters from DEL to &HA0 which are commonly
'     considered control characters and probably should be encoded in
'     JSON strings as well.
'
'Version 1.2 changes:
'
'   o Handle parsing of "empty" JSON arrays and objects properly.  Was
'     blowing up with error 5.
'

'Members will hold either a:
'
'    o Variant containing an array, or
'
'    o Scripting.Dictionary, with CompareMode = vbBinaryCompare (the
'      default for Dictionary objects when created).
'
'... depending on the variety of JNode.
Private Members As Variant

'These are used to support two inline "string builders."  The S-
'prefixed builder items are used for the "secondary string builder."
Private Buf As String
Private BufCursor As Long
Private SBuf As String
Private SBufCursor As Long

'Pseudo-constants used in encoding string values:
Private DEL As String
Private H9F As String

Private SerializeWhitespace As Boolean
Private mWhitespace As String

Public Property Get Count() As Long
    If IsObject Then
        Count = Members.Count
    ElseIf IsEmpty(Members) Then
        Count = 0
    Else
        Count = UBound(Members) + 1
    End If
End Property

Public Property Get Exists(ByVal ID As Variant) As Boolean
    Dim VarTypeOfID As VbVarType
    
    VarTypeOfID = VBA.VarType(ID)
    If IsObject Then
        If VarTypeOfID <> vbString Then Err.Raise 5
            
        Exists = Members.Exists(ID)
    Else
        If VarTypeOfID <> vbInteger And VarTypeOfID <> vbLong Then Err.Raise 5
        
        Exists = LBound(Members) <= ID And ID <= UBound(Members)
    End If
End Property

Public Property Get Keys() As Variant
    If IsObject Then
        Keys = Members.Keys()
    Else
        Err.Raise 5 'JSON "arrays" do not have keys.
    End If
End Property

Public Property Get IsArray() As Boolean
    If Not IsEmpty(Members) Then
        IsArray = Not VBA.IsObject(Members)
    End If
End Property

Public Property Get IsObject() As Boolean
    IsObject = VBA.IsObject(Members)
End Property

Public Property Get JSON( _
    Optional ByVal Whitespace As String, _
    Optional ByVal Level As Integer) As String
    
    'Note that the Level argument is meant for internal use.
    Dim MyKeys As Variant
    Dim I As Long
    
    mWhitespace = Whitespace
    SerializeWhitespace = Len(mWhitespace) > 0
    If IsObject Then
        BufCat "{": If SerializeWhitespace Then BufCatWhitespace Level + 1
        MyKeys = Keys
        For I = 0 To UBound(MyKeys)
            BufCat SerializeValue(MyKeys(I), 0)
            BufCat ":": If SerializeWhitespace Then BufCat " "
            BufCat SerializeValue(Members(MyKeys(I)), Level + 1)
            If I < UBound(MyKeys) Then
                BufCat ",": If SerializeWhitespace Then BufCatWhitespace Level + 1
            Else
                If SerializeWhitespace Then BufCatWhitespace Level
            End If
        Next
        BufCat "}"
    ElseIf IsArray Then
        BufCat "[": If SerializeWhitespace Then BufCatWhitespace Level + 1
        For I = 0 To UBound(Members)
            BufCat SerializeValue(Members(I), Level + 1)
            If I < UBound(Members) Then
                BufCat ",": If SerializeWhitespace Then BufCatWhitespace Level + 1
            Else
                If SerializeWhitespace Then BufCatWhitespace Level
            End If
        Next
        BufCat "]"
    Else
        Err.Raise 5
    End If
    JSON = BufExtract
End Property

Public Property Let JSON( _
    Optional ByVal Whitespace As String, _
    Optional ByVal Level As Integer, _
    ByRef RHS As String)
    
    'Note that Whitespace and Level shouldn't be used and here only defined in
    'order to match the Property Get signature above.
    Dim LenOfRHS As Long
    Dim Cursor As Long
    
    LenOfRHS = Len(RHS)
    Cursor = 1
    JSONNested RHS, Cursor, LenOfRHS
    
    For Cursor = Cursor + 1 To LenOfRHS
        Select Case Mid$(RHS, Cursor, 1)
            Case " ", vbCr, vbLf, vbTab, vbBack, vbFormFeed: 'Do nothing.
            Case Else: Err.Raise 5 'Non-whitepace after CloseDelim.
        End Select
    Next
End Property

'Default property:
Public Property Get Value(ByVal ID As Variant) As Variant
Attribute Value.VB_UserMemId = 0
    Dim VarTypeOfID As VbVarType
    
    VarTypeOfID = VBA.VarType(ID)
    If IsEmpty(Members) Then
        If VarTypeOfID = vbString Then
            Set Members = New Scripting.Dictionary
        Else
            If VarTypeOfID <> vbInteger And VarTypeOfID <> vbLong Then Err.Raise 5
            Members = Array()
        End If
    End If
    If IsObject Then
        If VarTypeOfID <> vbString Then Err.Raise 5
    Else
        If VarTypeOfID <> vbInteger And VarTypeOfID <> vbLong Then Err.Raise 5
        If ID < 0 Then Err.Raise 5
        If ID > UBound(Members) Then ReDim Preserve Members(ID)
    End If
    If VBA.IsObject(Members(ID)) Then
        Set Value = Members(ID)
    Else
        Value = Members(ID)
    End If
End Property

Public Property Let Value(ByVal ID As Variant, ByVal RHS As Variant)
    Dim VarTypeOfID As VbVarType
    
    VarTypeOfID = VBA.VarType(ID)
    If IsEmpty(Members) Then
        If VarTypeOfID = vbString Then
            Set Members = New Scripting.Dictionary
        Else
            If VarTypeOfID <> vbInteger And VarTypeOfID <> vbLong Then Err.Raise 5
            Members = Array()
        End If
    End If
    If VBA.IsObject(RHS) Then
        If Not TypeOf RHS Is JNode Then
            Err.Raise 5
        End If
    End If
    If IsObject Then
        If VarTypeOfID <> vbString Then Err.Raise 5
    Else
        If VarTypeOfID <> vbInteger And VarTypeOfID <> vbLong Then Err.Raise 5
        If ID < 0 Then Err.Raise 5
        If ID > UBound(Members) Then ReDim Preserve Members(ID)
    End If
    If VBA.VarType(RHS) = vbNull Then RHS = Empty
    If VBA.IsObject(RHS) Then
        Set Members(ID) = RHS
    Else
        Select Case VBA.VarType(RHS)
            Case vbInteger, vbLong, vbSingle, vbCurrency, vbDecimal, vbByte
                Members(ID) = CDbl(RHS)
            Case Else
                Members(ID) = RHS
        End Select
    End If
End Property

Public Property Set Value(ByVal ID As Variant, ByVal RHS As Variant)
    'Delegate this operation to the Property Let procedure:
    Value(ID) = RHS
End Property

Public Property Get VarType(ByVal ID As Variant) As VbVarType
    VarType = VBA.VarType(Members(ID))
End Property

Public Sub MakeArray()
    Members = Array()
End Sub

Public Sub MakeObject()
    Set Members = New Scripting.Dictionary
End Sub

Public Sub Remove(ByVal ID As String)
    If IsObject Then
        Members.Remove ID
    Else
        Err.Raise 5
    End If
End Sub

'============================== Recursing Parser Routine ===========================
'Called by Property Let JSON for the heavy lifting of parsing... and also called by
'itself on any new JNode members of this instance as they are created.
'
'Not meant to be called by application code using JNode objects.  Public Hidden
'instead of Friend since we can't hide Friend members.
Public Sub JSONNested(ByRef Text As String, ByRef Cursor As Long, ByVal LenOfText As Long)
Attribute JSONNested.VB_MemberFlags = "40"
    Dim ThisWasStarted As Boolean
    Dim ThisIsObject As Boolean
    Dim Ch As String
    Dim CloseDelim As String
    Dim ID As String
    Dim Count As Long
    Dim Value As Variant
    Dim JNode As JNode
    
    Members = Empty
    For Cursor = Cursor To LenOfText
        Select Case Mid$(Text, Cursor, 1)
            Case "{"
                MakeObject
                CloseDelim = "}"
                ThisIsObject = True
                ThisWasStarted = True
                Exit For
            Case "["
                MakeArray
                CloseDelim = "]"
                ThisWasStarted = True
                Exit For
        End Select
    Next
    If IsEmpty(Members) Then Err.Raise 5 'No OpenDelim found.
    
    For Cursor = Cursor + 1 To LenOfText
        Ch = Mid$(Text, Cursor, 1)
        Select Case Ch
            Case """"
                Value = ParseString(Text, Cursor, LenOfText)
            Case "0" To "9", "-"
                Value = ParseNumber(Text, Cursor, LenOfText)
            Case "{", "["
                Set JNode = New JNode 'Need early-bound reference to use Friend members.
                JNode.JSONNested Text, Cursor, LenOfText
                Set Value = JNode
            Case "}", "]"
                GoTo ProcessCloseDelim
            Case "f", "t", "n"
                Value = ParseWordValue(Text, Cursor, LenOfText)
            Case Else
                GoTo GetNextChar
        End Select
        
        If ThisIsObject Then
            If Len(ID) > 0 Then
                If VBA.VarType(Value) = vbObject Then
                    Set Members.Item(ID) = Value
                Else
                    Members.Item(ID) = Value
                End If
                ID = vbNullString
                Ch = SkipToCommaOrCloseDelim(Text, Cursor, LenOfText, CloseDelim)
            Else
                ID = Value
                SkipToColon Text, Cursor, LenOfText
                GoTo GetNextChar
            End If
        Else
            If Count > UBound(Members) Then ReDim Preserve Members(Count + 10)
            If VBA.VarType(Value) = vbObject Then
                Set Members(Count) = Value
            Else
                Members(Count) = Value
            End If
            Count = Count + 1
            Ch = SkipToCommaOrCloseDelim(Text, Cursor, LenOfText, CloseDelim)
        End If
        
ProcessCloseDelim:
        If Ch = "}" Then
            If ThisIsObject And ThisWasStarted Then
                Exit Sub
            Else
                Err.Raise 5
            End If
        ElseIf Ch = "]" Then
            If Not ThisIsObject And ThisWasStarted Then
                If Count - 1 < UBound(Members) Then ReDim Preserve Members(Count - 1)
                Exit Sub
            Else
                Err.Raise 5
            End If
        End If
        
GetNextChar:
    Next
    Err.Raise 5 'Fell off the end of Text.
End Sub

'============================== String Builder Routines ============================
Private Sub BufClear()
    Buf = vbNullString
    BufCursor = 1
End Sub

Private Sub BufCat(ByVal Text As String)
    Const CHUNK As Long = 128
    Dim LenOfText As Long
    
    LenOfText = Len(Text)
    If Len(Buf) - BufCursor + 1 < LenOfText Then
        If CHUNK > LenOfText Then
            Buf = Buf & Space$(CHUNK)
        Else
            Buf = Buf & Space$(LenOfText)
        End If
    End If
    Mid$(Buf, BufCursor, LenOfText) = Text
    BufCursor = BufCursor + LenOfText
End Sub

Private Sub BufCatWhitespace(ByVal Level As Integer)
    Dim I As Integer
    
    BufCat vbNewLine
    For I = 1 To Level
        BufCat mWhitespace
    Next
End Sub

Private Function BufExtract() As String
    BufExtract = Left$(Buf, BufCursor - 1)
    BufClear
End Function

Private Sub SBufClear()
    SBuf = vbNullString
    SBufCursor = 1
End Sub

Private Sub SBufCat(ByVal Text As String)
    Const CHUNK As Long = 64
    Dim LenOfText As Long
    
    LenOfText = Len(Text)
    If Len(SBuf) - SBufCursor + 1 < LenOfText Then
        If CHUNK > LenOfText Then
            SBuf = SBuf & Space$(CHUNK)
        Else
            SBuf = SBuf & Space$(LenOfText)
        End If
    End If
    Mid$(SBuf, SBufCursor, LenOfText) = Text
    SBufCursor = SBufCursor + LenOfText
End Sub

Private Function SBufExtract() As String
    SBufExtract = Left$(SBuf, SBufCursor - 1)
    SBufClear
End Function

'=========================== Serialization Helper Routines =========================
Private Function SerializeValue(ByVal Value As Variant, ByVal Level As Integer) As String
    Dim ValueString As String
    Dim I As Long
    Dim Ch As String
    
    Select Case VBA.VarType(Value)
        Case vbEmpty
            SerializeValue = "null"
        Case vbBoolean
            If Value Then SerializeValue = "true" Else SerializeValue = "false"
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            Ch = LTrim$(Str$(Abs(Value)))
            If Left$(Ch, 1) = "." Then Ch = "0" & Ch
            If Value < 0 Then SerializeValue = "-" & Ch Else SerializeValue = Ch
        Case vbString
            ValueString = Value
            SBufCat """"
            For I = 1 To Len(ValueString)
                Ch = Mid$(ValueString, I, 1)
                Select Case Ch
                    Case "\":        Ch = "\\"
                    Case """":       Ch = "\"""
                    Case vbLf:       Ch = "\n"
                    Case vbCr:       Ch = "\r"
                    Case vbTab:      Ch = "\t"
                    Case vbBack:     Ch = "\b"
                    Case vbFormFeed: Ch = "\f"
                    Case Is < " ", DEL To H9F
                        Ch = "\u" & Right$("000" & Hex$(AscW(Ch)), 4)
                    'Case Else 'Do nothing.
                End Select
                SBufCat Ch
            Next
            SBufCat """"
            SerializeValue = SBufExtract()
        Case vbObject
            SerializeValue = Value.JSON(mWhitespace, Level)
    End Select
End Function

'============================== Parsing Helper Routines ============================
Private Sub SkipToColon( _
    ByRef Text As String, _
    ByRef Cursor As Long, _
    ByVal LenOfText As Long)
    
    For Cursor = Cursor + 1 To LenOfText
        If Mid$(Text, Cursor, 1) = ":" Then Exit Sub
    Next
    Err.Raise 5 'Fell off end of Text.
End Sub

Private Function SkipToCommaOrCloseDelim( _
    ByRef Text As String, _
    ByRef Cursor As Long, _
    ByVal LenOfText As Long, _
    ByVal CloseDelim As String) As String
    
    For Cursor = Cursor + 1 To LenOfText
        SkipToCommaOrCloseDelim = Mid$(Text, Cursor, 1)
        If SkipToCommaOrCloseDelim = "," Or Mid$(Text, Cursor, 1) = CloseDelim Then Exit Function
    Next
    Err.Raise 5 'Fell off end of Text.
End Function

Private Function ParseString( _
    ByRef Text As String, _
    ByRef Cursor As Long, _
    ByVal LenOfText As Long) As String
    
    Dim Ch As String
    
    For Cursor = Cursor + 1 To LenOfText
        Ch = Mid$(Text, Cursor, 1)
        Select Case Ch
            Case """": ParseString = SBufExtract(): Exit Function
            Case "\":
                Cursor = Cursor + 1
                Ch = Mid$(Text, Cursor, 1)
                Select Case LCase$(Ch) 'We'll make this forgiving though lowercase is proper.
                    Case "\", "/": Ch = Ch
                    Case """":      Ch = """"
                    Case "n":      Ch = vbLf
                    Case "r":      Ch = vbCr
                    Case "t":      Ch = vbTab
                    Case "b":      Ch = vbBack
                    Case "f":      Ch = vbFormFeed
                    Case "u":      Ch = ParseHexChar(Text, Cursor, LenOfText)
                    'Case Else: 'Do nothing.
                End Select
        End Select
        SBufCat Ch
    Next
    Err.Raise 5 'Fell off end of the Text.
End Function

Private Function ParseHexChar( _
    ByRef Text As String, _
    ByRef Cursor As Long, _
    ByVal LenOfText As Long) As String
    
    Const ASCW_OF_ZERO As Long = &H30&
    Dim Length As Long
    Dim Ch As String
    Dim DigitValue As Long
    Dim Value As Long

    For Cursor = Cursor + 1 To LenOfText
        Ch = Mid$(Text, Cursor, 1)
        Select Case Ch
            Case "0" To "9", "A" To "F", "a" To "f"
                Length = Length + 1
                If Length > 4 Then Exit For
                If Ch > "9" Then
                    DigitValue = (AscW(Ch) And &HF&) + 9
                Else
                    DigitValue = AscW(Ch) - ASCW_OF_ZERO
                End If
                Value = Value * &H10& + DigitValue
            Case Else
                Exit For
        End Select
    Next
    If Length = 0 Then Err.Raise 5 'No hex digits at all.
    Cursor = Cursor - 1
    ParseHexChar = ChrW$(Value)
    Exit Function
End Function

Private Function ParseNumber( _
    ByRef Text As String, _
    ByRef Cursor As Long, _
    ByVal LenOfText As Long) As Double
    
    Dim Start As Long
    
    Start = Cursor
    For Cursor = Cursor + 1 To LenOfText
        Select Case LCase$(Mid$(Text, Cursor, 1))
            Case "0" To "9", ".", "e", "+", "-": 'Do nothing.
            Case Else: Exit For
        End Select
    Next
    ParseNumber = Val(Mid$(Text, Start, Cursor - Start))
    Cursor = Cursor - 1
End Function

Private Function ParseWordValue( _
    ByRef Text As String, _
    ByRef Cursor As Long, _
    ByVal LenOfText As Long) As Variant
    
    Dim Start As Long
    Dim WordValue As String
    
    Start = Cursor
    For Cursor = Cursor + 1 To LenOfText
        Select Case Mid$(Text, Cursor, 1)
            Case "a" To "u": 'Do nothing.
            Case Else: Exit For
        End Select
    Next
    WordValue = Mid$(Text, Start, Cursor - Start)
    Cursor = Cursor - 1
    Select Case WordValue
        Case "false": ParseWordValue = False
        Case "true":  ParseWordValue = True
        Case "null":  ParseWordValue = Empty
        Case Else: Err.Raise 5
    End Select
End Function

Private Sub Class_Initialize()
    DEL = ChrW$(&H7F&)
    H9F = ChrW$(&H9F&)
    BufClear
    SBufClear
End Sub
