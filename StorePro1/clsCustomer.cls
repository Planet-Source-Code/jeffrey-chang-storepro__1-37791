VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Datachanged() 'A custom event

'Public members to implement properties
Public strFirst As String
Public strLast As String
Public lngCustId As Long
Public strAdd1 As String
Public strAdd2 As String
Public strPCode As String
Public strHPhone As String
Public strCPhone As String
Public dateLastSale As Date
Public curYTD As Currency
Public dateJoin As Date

'Private member for simulated inheritance
'To delegate the task in the member variable, Rs, to the
'private instance of clsRecordset; i.e. m_ObjRecordset
Private m_ObjRecordset As New clsRecordset
Private Sub Class_Initialize()
Dim strRs As String
    On Error GoTo Class_Initialize_Error
        strRs = "CUSTOMER"
        BindData (strRs)
    Exit Sub
Class_Initialize_Error:
        MsgBox Err.Description
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
        Rs.Close
        Cn.Close
    Exit Sub
End Sub
Private Sub Cust_Reload_Members()
'For displaying the contents of the recordset on the associated forms
    On Error GoTo Cust_Reload_Members_Error
        With Rs
            strFirst = .Fields("1stName_")
            strLast = .Fields("LastName_")
            lngCustId = .Fields("CustNo_")
            strAdd1 = .Fields("Add1_")
            strAdd2 = .Fields("Add2_")
            strPCode = .Fields("Pcode_")
            strHPhone = .Fields("HomePhone_")
            strCPhone = .Fields("CellPhone_")
            dateLastSale = .Fields("LastSale_")
            curYTD = .Fields("YTDSale_")
            dateJoin = .Fields("Join_")
        End With
        
        RaiseEvent Datachanged
        
        Exit Sub
Cust_Reload_Members_Error:
        MsgBox Err.Description
End Sub
Private Sub Cust_Save_Members()
'For transferring the contents on the form into the recordset
On Error GoTo Cust_Save_Members_Error
    With Rs
        .Fields("1stName_") = frmCustomer.txtFirst(0)
        .Fields("LastName_") = frmCustomer.txtLast(1)
        .Fields("CustNo_") = frmCustomer.txtCustId(0)
        .Fields("Add1_") = frmCustomer.txtAdd1(1)
        .Fields("Add2_") = frmCustomer.txtAdd2(2)
        .Fields("Pcode_") = frmCustomer.txtPCode(3)
        .Fields("HomePhone_") = frmCustomer.txtHPhone(4)
        .Fields("CellPhone_") = frmCustomer.txtCPhone(5)
        .Fields("LastSale_") = frmCustomer.txtLastSale(0)
        .Fields("YTDSale_") = frmCustomer.txtYTD(1)
        .Fields("Join_") = frmCustomer.txtJoin(2)
    End With
    Exit Sub
Cust_Save_Members_Error:
    Exit Sub
End Sub
Private Sub Cust_Clear_Members()
'For clearing the contents on the form to a blank form ready for input
On Error GoTo Cust_Clear_Members_Error
    With frmCustomer
        .txtFirst(0) = ""
        .txtLast(1) = ""
        .txtCustId(0) = ""
        .txtAdd1(1) = ""
        .txtAdd2(2) = ""
        .txtPCode(3) = ""
        .txtHPhone(4) = ""
        .txtCPhone(5) = ""
        .txtLastSale(0) = ""
        .txtYTD(1) = ""
        .txtJoin(2) = ""
    End With
    Exit Sub
Cust_Clear_Members_Error:
    Exit Sub
End Sub
Public Sub MoveNext()
    m_ObjRecordset.MoveNext Rs
    Cust_Reload_Members
End Sub
Public Sub MovePrevious()
    m_ObjRecordset.MovePrevious Rs
    Cust_Reload_Members
End Sub
Public Sub MoveLast()
    m_ObjRecordset.MoveLast Rs
    Cust_Reload_Members
End Sub
Public Sub MoveFirst()
    m_ObjRecordset.MoveFirst Rs
    Cust_Reload_Members
End Sub
Public Sub AddAsNew()
    frmCustomer.txtFirst(0).SetFocus
    m_ObjRecordset.AddAsNew Rs
    Cust_Clear_Members
End Sub
Public Sub SaveChanges()
    Cust_Save_Members
    m_ObjRecordset.SaveChanges Rs
End Sub
Public Sub DeleteCurrent()
    m_ObjRecordset.DeleteCurrent Rs
    Cust_Reload_Members
End Sub
Public Sub FindRecords()
    Dim strName As String
    strName = InputBox("Enter the full or partial First Name", "Find")
    m_ObjRecordset.FindRecords Rs, strName
    Cust_Reload_Members
End Sub
