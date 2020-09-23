Attribute VB_Name = "modBindData"
Option Explicit
Global Rs As New ADODB.Recordset
Global Cn As New ADODB.Connection
Public Sub BindData(strRs As String)
    On Error GoTo Mod_BindData_Error
    Cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" _
                         & "Data Source=" & App.Path & "\StorePro.mdb"
    Cn.CursorLocation = adUseServer
    Cn.Open
    Rs.Open strRs, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    Exit Sub
Mod_BindData_Error:
    MsgBox "The database or table cannot be opened", vbExclamation, "Connection Error"
End Sub


