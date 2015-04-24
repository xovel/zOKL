Attribute VB_Name = "modBaseCollection"
'------------------------------------------------------------------------------
'这个文件定义了一些辅助变量和函数
'一般情况下您不需要看懂，也不需要自己修改这个文件
'如果您没有特别的需要，请不要改变这个文件里面的任何内容
'------------------------------------------------------------------------------
Option Explicit
Public ControlDataCollection As New Collection

Public Sub SaveControlData(ControlName As String, ControlValue As Variant)
    If VarType(ControlValue) = vbBoolean Then
        ControlValue = CInt(ControlValue)
    End If

    Dim TempControlData As ControlData
    With TempControlData
        .Name = ControlName
        .Value = ControlValue
    End With

    ControlDataCollection.Add TempControlData, ControlName
End Sub

Public Function LoadControlData(ControlName As String) As String
    LoadControlData = ControlDataCollection(ControlName).Value
End Function

Public Sub ClearControlData()
    On Error Resume Next
    Set ControlDataCollection = Nothing
    Set ControlDataCollection = New Collection
End Sub

'以上的内容，一般情况下您不需要看懂，也不需要进行修改



