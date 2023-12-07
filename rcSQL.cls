VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcSQL
' Version: 0.3
' Author: Rafael Fillipe Silva
' Description: ...

Public FilePath As String
Private ConnectionString As String
Private Connection As Object
Public RecordSet As Object

Public Function OpenConnection(ByVal File As String) As Boolean
    FilePath = File
    ConnectionString = ""

    Set Connection = Nothing
    Set RecordSet = Nothing

    If FilePath = "" Then
        Exit Function
    End If

    ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" & FilePath & """;Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""

    Set Connection = CreateObject("ADODB.Connection")
    Set RecordSet = CreateObject("ADODB.Recordset")

    Connection.Open ConnectionString
End Function

Public Function ConvertRange(R As Range) As String
    ConvertRange = "[" & R.Worksheet.Name & "$"

    If R.Count > 1 Then
        ConvertRange = ConvertRange & R.Address(0, 0)
    Else
        ConvertRange = ConvertRange & R.Address(0, 0) & ":" & R.Address(0, 0)
    End If

    ConvertRange = ConvertRange & "]"
End Function

Public Function Query(ByVal sql As String) As Boolean
Attribute Query.VB_UserMemId = 0
    On Error Resume Next
    RecordSet.Open sql, Connection
    Query = True
End Function

Public Sub CloseConnection()
End Sub

