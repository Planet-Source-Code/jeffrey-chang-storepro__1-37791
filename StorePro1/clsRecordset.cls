VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim varBookmark As Variant
Public Sub MoveNext(recRecords As ADODB.Recordset)
With recRecords
    .MoveNext
    If .EOF Then .MoveLast
End With
End Sub
Public Sub MovePrevious(recRecords As ADODB.Recordset)
With recRecords
    .MovePrevious
    If .BOF Then .MoveFirst
End With
End Sub
Public Sub MoveLast(recRecords As ADODB.Recordset)
recRecords.MoveLast
End Sub
Public Sub MoveFirst(recRecords As ADODB.Recordset)
recRecords.MoveFirst
End Sub
Public Sub AddAsNew(recRecords As ADODB.Recordset)
recRecords.AddNew
End Sub
Public Sub SaveChanges(recRecords As ADODB.Recordset)
With recRecords
    .Update
End With
End Sub
Public Sub DeleteCurrent(recRecords As ADODB.Recordset)
recRecords.Delete
On Error Resume Next
recRecords.MoveFirst
End Sub
Public Sub FindRecords(recRecords As ADODB.Recordset, strName As String)
On Error GoTo Class_FindRecords_Error
With recRecords
    .MoveFirst
    .Find " [1stName_] LIKE '" & strName & "*'"
End With
Exit Sub
Class_FindRecords_Error:
MsgBox "Cannot find matching records"
End Sub
