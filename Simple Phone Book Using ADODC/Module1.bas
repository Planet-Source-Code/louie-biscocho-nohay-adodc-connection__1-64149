Attribute VB_Name = "Module1"
' ======================================
' Codes by   :    Louie Biscocho Nohay
' Email Add:      lbnohay@yahoo.com
' Website:        www.noborsoft.cjb.net
' Tel. No.   :    +63.43.984.8338
' ======================================

Option Explicit

' connect to database using ADODC

Public Sub get_connection(ByRef objADODC As Adodc, ByVal sRecord As String, ByVal sLocation As String, ByVal bPassword As Boolean, ByVal sPassword As String)

If bPassword = True Then
   objADODC.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sLocation & "; Persist Security Info= False; Jet OLEDB: Database Password=" & sPassword
Else
   objADODC.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sLocation & "; Persist Security Info= False"
End If

objADODC.CommandType = adCmdText
objADODC.RecordSource = sRecord
objADODC.Refresh

End Sub
' inform messaging
Public Sub Inform(ByVal msg As String, ByVal sTitle As String)
MsgBox msg, vbOKOnly + vbInformation, sTitle
End Sub

' exclaim messaging
Public Sub Exclaim(ByVal msg As String, ByVal sTitle As String)
MsgBox msg, vbOKOnly + vbExclamation, sTitle
End Sub

' critical messaging
Public Sub Question(ByRef msg As String, ByVal sTitle As String)
MsgBox msg, vbYesNo + vbCritical, sTitle
End Sub

' Highlight textboxes
Public Sub HyLyt(ByRef objText As TextBox)
With objText
   .SelStart = 0
   .SelLength = Len(objText.Text)
End With
End Sub
