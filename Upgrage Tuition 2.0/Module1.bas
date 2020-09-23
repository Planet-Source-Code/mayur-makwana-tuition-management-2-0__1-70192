Attribute VB_Name = "Module1"
Public cn As New ADODb.Connection
Public rs As New ADODb.Recordset
Public Status As Boolean
Public txtbox As Control

Public Sub Con()
Set cn = Nothing
cn.CursorLocation = adUseClient
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Tuition.mdb;Persist Security Info=False"
End Sub
Public Sub Changecolor(Action As Boolean, frm As Control)
On Error Resume Next
    If Action = True Then
        Set txtbox = Nothing
        Set txtbox = frm
        txtbox.BackColor = &HFFFFC0
    Else
        txtbox.BackColor = vbWhite
    End If
End Sub
