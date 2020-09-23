Attribute VB_Name = "mMain"
Public cConnect As String  ' "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False; User ID=;Data Source=D:\Irvin Workfolder\BohemianExpressTravel\DATA\DB.mdb;Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Locale Identifier=1033;Jet OLEDB:Registry Path='';Jet OLEDB:Database Password="
Public fMainForm As frmMain

Public Function is_empty(ByRef sText As Variant) As Boolean
If sText.Text = "" Then
    is_empty = True
    MsgBox "The field is required.Please check it!", vbExclamation, "System version 1.0"
    sText.SetFocus
Else
    is_empty = False
End If
End Function
Public Sub search_in_listview(ByRef sListView As ListView, ByVal sFindText As String)
Dim tmp_listtview As ListItem
    Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem + lvwText, lvwPartial, lvwPartial)
    If Not tmp_listtview Is Nothing Then
    tmp_listtview.EnsureVisible
    tmp_listtview.Selected = True
End If
End Sub

Public Function is_Numeric(ByRef sText As String) As Boolean
    If IsNumeric(sText) = False Then
        is_Numeric = False
        MsgBox "The field required a numeric input.Please check it!", vbExclamation
    Else
        is_Numeric = True
    End If
End Function

 Sub Main()
    Set fMainForm = New frmMain
    Set fLogin = New frmLogin
    
    cConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;User ID=;Data Source=" & App.Path & "\Data\dbBohemianTravel.mdb;Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Locale Identifier=1033;Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='danieldave';Jet OLEDB:Global Partial Bulk Ops=2"

    fMainForm.WindowState = vbMaximized
    fMainForm.Show
    fLogin.Show vbModal
 End Sub

