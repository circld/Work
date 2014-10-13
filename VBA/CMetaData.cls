VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMetaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private RowF, RowL, ColF, ColL  As Long
Private ItemName                As String
Private Measure(1 To 4)         As Variant

Public Property Get Name() As String
    Name = ItemName
End Property

Public Property Let Name(Value As String)
    ItemName = Value
End Property

Public Property Get FirstRow() As Long
    FirstRow = RowF
End Property

Public Property Let FirstRow(Value As Long)
    RowF = Value
End Property

Public Property Get LastRow() As Long
    LastRow = RowL
End Property

Public Property Let LastRow(Value As Long)
    RowL = Value
End Property

Public Property Get FirstCol() As Long
    FirstCol = ColF
End Property

Public Property Let FirstCol(Value As Long)
    ColF = Value
End Property

Public Property Get LastCol() As Long
    LastCol = ColL
End Property

Public Property Let LastCol(Value As Long)
    ColL = Value
End Property

Public Property Get Val1() As Variant
    Val1 = Measure(1)
End Property

Public Property Let Val1(Value As Variant)
    Measure(1) = Value
End Property

Public Property Get Val2() As Variant
    Val2 = Measure(2)
End Property

Public Property Let Val2(Value As Variant)
    Measure(2) = Value
End Property

Public Property Get Val3() As Variant
    Val3 = Measure(3)
End Property

Public Property Let Val3(Value As Variant)
    Measure(3) = Value
End Property
Public Property Get Val4() As Variant
    Val4 = Measure(4)
End Property

Public Property Let Val4(Value As Variant)
    Measure(4) = Value
End Property

Public Property Get DataRange() As Range
' Automatically define DataRange when first and last rows/cols are defined
    If FirstRow <> 0 & LastRow <> 0 & FirstCol <> 0 & LastCol <> 0 Then
        Set DataRange = Range(Cells(RowF, ColF), Cells(RowL, ColL))
    Else
        MsgBox "Please fully specify FirstRow, FirstCol, LastRow, LastCol"
    End If
End Property
