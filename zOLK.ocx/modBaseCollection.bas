Attribute VB_Name = "modBaseCollection"
'------------------------------------------------------------------------------
'����ļ�������һЩ���������ͺ���
'һ�������������Ҫ������Ҳ����Ҫ�Լ��޸�����ļ�
'�����û���ر����Ҫ���벻Ҫ�ı�����ļ�������κ�����
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

'���ϵ����ݣ�һ�������������Ҫ������Ҳ����Ҫ�����޸�



