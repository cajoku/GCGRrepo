VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Key_ As String
Private Val_ As Variant
Private Type_ As dataType
Private Dbl_ As Double
Private Generation_ As Integer
Private Parent_ As Variant
Public Obj As Collection
Private Const cL As String = ":"
Private Const cM As String = ","

Private Enum dataType
    OBJECTjs = 1
    ARRAYjs = 2
    STRINGjs = 3
    NUMBERjs = 4
    BOOLjs = 5
    NULLjs = 6
End Enum

Private Sub Class_Initialize()

    Set Parent_ = Application

End Sub

Property Let Key(keyVal As String)

    Dim str As String
    str = Replace(keyVal, Chr(34), "\" & Chr(34))
    str = Replace(str, "\\" & Chr(34), "\" & Chr(34))
    Key_ = str

End Property

Property Get JSONstr(Optional omitKey As Boolean) As String
    Dim prefix As String
    Dim suffix As String
    Dim inDent As String
    Dim txt As String
    Dim str As String
    Dim i As Integer
    Dim var
    
    On Error Resume Next
    inDent = tabCt(myI())
    Select Case Type_
        Case 1
            'object
            If omitKey Then
                prefix = inDent & "{" & vbLf
            Else
                prefix = inDent & q(Key_) & cL & vbLf & inDent & "{" & vbLf
            End If
            suffix = vbLf & inDent & "}," & vbLf
            For Each var In Obj
                txt = txt & var.JSONstr
            Next
            txt = prefix & Left(txt, Len(txt) - 1) & suffix
            
        Case 2
            'array
            prefix = inDent & q(Key_) & cL & vbLf & inDent & "[" & vbLf
            suffix = vbLf & inDent & "]," & vbLf
            If UBound(Val_) = 0 And Val_(0) Is Nothing Then
                txt = ""
            Else
                For i = LBound(Val_) To UBound(Val_)
                    txt = txt & Val_(i).JSONstr(True)
                Next
                txt = Left(txt, Len(txt) - 2)
            End If
            txt = prefix & txt & suffix

        Case 3
            'string
            If omitKey Then
                prefix = inDent
            Else
                prefix = inDent & q(Key_) & cL
            End If
            suffix = cM & vbLf
            txt = prefix & q(Val_) & suffix
        
        Case 4
            'number
            If omitKey Then
                prefix = inDent
            Else
                prefix = inDent & q(Key_) & cL
            End If
            suffix = cM & vbLf
            txt = prefix & Val_ & suffix
            
        Case 5
            'bool
            prefix = inDent & q(Key_) & cL
            suffix = cM & vbLf
            If Val_ Then str = "TRUE" Else str = "FALSE"
            txt = prefix & str & suffix
                        
        Case 6
            'null
            prefix = inDent & q(Key_) & cL
            suffix = cM & vbLf
            txt = prefix & "Null" & suffix
            
    End Select
    
    JSONstr = txt


End Property

Property Get Value() As Variant

    Value = Val_

End Property

Property Set Value(val As Variant)
    
    Type_ = OBJECTjs
    If TypeName(val) = "Collection" Then
        Set Val_ = val
        Set Obj = val
    Else
        Set Obj = New Collection
        'Obj.Add val
    End If

End Property

Property Let Value(val As Variant)

    Select Case TypeName(val)
        Case "Null"
            Type_ = NULLjs
            Val_ = Null
        Case "Integer"
            Type_ = NUMBERjs
            Dbl_ = val
            Val_ = Dbl_
        Case "Double"
            Type_ = NUMBERjs
            Val_ = val
        Case "Long"
            Type_ = NUMBERjs
            Dbl_ = val
            Val_ = Dbl_
        Case "Currency"
            Type_ = NUMBERjs
            Dbl_ = val
            Val_ = Dbl_
        Case "String"
            Type_ = STRINGjs
            Val_ = Replace(val, Chr(34), "\" & Chr(34))
            'Val_ = Replace(val, "\\" & Chr(34), "\" & Chr(34))
        Case "Boolean"
            Type_ = BOOLjs
            Val_ = val
        Case "Date"
            Type_ = STRINGjs
            Val_ = val
        Case Else
            If Right(TypeName(val), 2) = "()" Then
                Type_ = ARRAYjs
                Val_ = val
            Else
                Type_ = NULLjs
                Val_ = Null
            End If
    End Select

End Property

Property Set Parent(var As Variant)
    Set Parent_ = var
End Property

Property Get Parent() As Variant
    Set Parent = Parent_
End Property

Function addArray(keySTR As String, Optional arr As Variant) As JSON
        
    If Obj Is Nothing Then
        'TODO, test if value already set
        Set Obj = New Collection
        Type_ = OBJECTjs
    End If
    
    Dim tArr(0) As JSON

    Set addArray = Me.addObj
    With addArray
        .Key = keySTR
        If Not IsMissing(arr) And Right(TypeName(arr), 2) = "()" Then
            .Value = arr
        Else
            .Value = tArr
        End If
    End With

End Function

Sub appendArray(JSONobj As JSON)
    Dim i As Integer
    
    If Type_ <> ARRAYjs Then GoTo e1
    
    Set JSONobj.Parent = Me
    
    If UBound(Val_) = 0 And Val_(0) Is Nothing Then
        Set Val_(0) = JSONobj
    Else
        i = UBound(Val_) + 1
        ReDim Preserve Val_(i)
        Set Val_(i) = JSONobj
    End If

Exit Sub
e1:
    Debug.Print e_("appendArray performed on JSON not of type Array" _
    & vbLf & "Append failed, original value retained.")

End Sub

Sub addLeaf(keySTR As String, val As Variant)
    
    If Type_ <> OBJECTjs And Type_ <> 0 Then GoTo e1
    
    If Obj Is Nothing Then
        'TODO, test if value already set
        Set Obj = New Collection
        Type_ = OBJECTjs
    End If
    
    Dim tJSON As JSON
    Set tJSON = Me.addObj
    With tJSON
        .Key = keySTR
        .Value = val
    End With
    
Exit Sub
e1:
    Debug.Print e_("addLeaf performed on JSON not of type Object" _
    & vbLf & "Append failed, original value retained.")
    
End Sub

Function addObj() As JSON

    If Obj Is Nothing Then
        Set Obj = New Collection
        Type_ = OBJECTjs
    End If

    Set addObj = New JSON
    Set addObj.Parent = Me
    Obj.Add addObj
    
End Function

Function createJSONfile(myPath As String, Optional fileName As String)
    
    Dim file As Object
    Dim fso As Object
    Dim str As String
    
    If fileName = "" Then fileName = "JSONpkg_" & Format(Now(), "mmm-dd-yy_hh_mm")
    
    If Right(myPath, 1) <> "\" Then myPath = myPath + "\"
    str = "data ={" & vbLf & JSONstr & vbLf & "}"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.createtextfile(myPath & fileName & ".js")
    file.writeline str
    file.Close
    Set file = Nothing
    Set fso = Nothing

End Function

Private Function myI() As Integer
    Dim var
    Dim i As Integer
    Const iMax As Integer = 20
    
    Set var = Me
    Do Until TypeName(var) = "Application" Or i >= iMax
        i = i + 1
        Set var = var.Parent
    Loop
    myI = i
    
End Function

Private Function tabCt(ct As Integer) As String
    Dim i As Integer
    
    For i = 1 To ct
        tabCt = tabCt + vbTab
    Next

End Function

Private Function q(str As Variant) As String
    
    q = Chr(34) & str & Chr(34)

End Function

Private Function e_(str As String) As String

    e_ = "~~~~~~~~~~~~~~~JSON CLASS ERROR~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & _
        vbLf & str & vbLf & _
        "JSON key:= " & Key_ & ",   dataType:= " & dataTypeName(Type_) & vbLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

End Function

Private Function dataTypeName(i As Long) As String
    
    dataTypeName = Array("Undefined", "Object", "Array", "String", "Number", "Bool", "Null")(i)

End Function

