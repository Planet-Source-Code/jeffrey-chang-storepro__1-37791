VERSION 5.00
Begin VB.Form frmCustomer 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   495
      Index           =   3
      Left            =   7550
      TabIndex        =   31
      Top             =   4300
      Width           =   1200
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   495
      Index           =   0
      Left            =   150
      MaskColor       =   &H8000000F&
      TabIndex        =   30
      Top             =   4300
      Width           =   1200
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   495
      Index           =   4
      Left            =   150
      TabIndex        =   29
      Top             =   5300
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Index           =   8
      Left            =   7550
      TabIndex        =   28
      Top             =   5300
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Index           =   6
      Left            =   5150
      TabIndex        =   27
      Top             =   5300
      Width           =   1200
   End
   Begin VB.CommandButton cmdAddAsNew 
      Caption         =   "Add"
      Height          =   495
      Index           =   5
      Left            =   2650
      TabIndex        =   26
      Top             =   5300
      Width           =   1200
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   495
      Index           =   2
      Left            =   5150
      TabIndex        =   25
      Top             =   4300
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Previous"
      Height          =   495
      Index           =   1
      Left            =   2650
      TabIndex        =   24
      Top             =   4300
      Width           =   1200
   End
   Begin VB.Frame fraMisc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Miscellaneous"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1500
      Left            =   4500
      TabIndex        =   12
      Top             =   1300
      Width           =   2600
      Begin VB.TextBox txtJoin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1200
         TabIndex        =   23
         Top             =   1020
         Width           =   1200
      End
      Begin VB.TextBox txtYTD 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Top             =   660
         Width           =   1200
      End
      Begin VB.TextBox txtLastSale 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   14
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label lblJoin 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Join Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1020
         Width           =   1000
      End
      Begin VB.Label lblYTD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "YTD Sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   660
         Width           =   1000
      End
      Begin VB.Label lblLastSale 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Last Sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   1000
      End
   End
   Begin VB.TextBox txtLast 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   1500
      TabIndex        =   4
      Top             =   780
      Width           =   4000
   End
   Begin VB.Frame fraCustInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Customer Profile"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2600
      Left            =   150
      TabIndex        =   2
      Top             =   1300
      Width           =   4000
      Begin VB.TextBox txtCPhone 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   1380
         TabIndex        =   21
         Top             =   2100
         Width           =   2400
      End
      Begin VB.TextBox txtHPhone 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   1380
         TabIndex        =   20
         Top             =   1740
         Width           =   2400
      End
      Begin VB.TextBox txtPCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1380
         TabIndex        =   19
         Top             =   1380
         Width           =   2400
      End
      Begin VB.TextBox txtAdd2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1380
         TabIndex        =   18
         Top             =   1020
         Width           =   2400
      End
      Begin VB.TextBox txtAdd1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   17
         Top             =   660
         Width           =   2400
      End
      Begin VB.TextBox txtCustId 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1380
         TabIndex        =   6
         Top             =   300
         Width           =   2400
      End
      Begin VB.Label lblCellPhone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cell Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2100
         Width           =   1100
      End
      Begin VB.Label lblHomePhone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Home Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1740
         Width           =   1100
      End
      Begin VB.Label lblPCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Postal Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1380
         Width           =   1100
      End
      Begin VB.Label lblAdd2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Address 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   1100
      End
      Begin VB.Label lblAdd1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Address 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   1100
      End
      Begin VB.Label lblCustId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Customer No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1100
      End
   End
   Begin VB.TextBox txtFirst 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   1500
      TabIndex        =   0
      Top             =   300
      Width           =   4000
   End
   Begin VB.Label lblLast 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   750
      Width           =   1100
   End
   Begin VB.Label lblFirst 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   300
      Width           =   1100
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_Customer As clsCustomer
Attribute m_Customer.VB_VarHelpID = -1
Private Sub cmdAddAsNew_Click(Index As Integer)
    m_Customer.AddAsNew
End Sub
Private Sub cmdDelete_Click(Index As Integer)
    m_Customer.DeleteCurrent
End Sub
Private Sub cmdFind_Click(Index As Integer)
    m_Customer.FindRecords
End Sub
Private Sub cmdFirst_Click(Index As Integer)
    m_Customer.MoveFirst
End Sub
Private Sub cmdLast_Click(Index As Integer)
    m_Customer.MoveLast
End Sub
Private Sub cmdNext_Click(Index As Integer)
    m_Customer.MoveNext
End Sub
Private Sub cmdPrev_Click(Index As Integer)
    m_Customer.MovePrevious
End Sub
Private Sub cmdSave_Click(Index As Integer)
    Dim strMsg As String
    strMsg = ValidateForm
    If strMsg = "" Then
        m_Customer.SaveChanges
    Else
        MsgBox strMsg, vbExclamation, "Validation Error"
    End If
End Sub
Private Sub Form_Load()
    Set m_Customer = New clsCustomer
End Sub
Private Sub m_Customer_DataChanged()
   With m_Customer
            txtFirst(0) = .strFirst
            txtLast(1) = .strLast
            txtCustId(0) = .lngCustId
            txtAdd1(1) = .strAdd1
            txtAdd2(2) = .strAdd2
            txtPCode(3) = .strPCode
            txtHPhone(4) = .strHPhone
            txtCPhone(5) = .strCPhone
            txtLastSale(0) = .dateLastSale
            txtYTD(1) = .curYTD
            txtJoin(2) = .dateJoin
    End With
End Sub
Public Function ValidateForm()
'Handles all the form-level validation and return an error
'message if needed. In this case, we only setup this code
'to check for the required fields
Dim strSendMsg As String

With txtFirst(0)
    If Len(Trim(.Text)) = 0 Then
        strSendMsg = strSendMsg & vbCrLf & _
                   "First Name is a required field" & vbCrLf
                   .SetFocus
    End If
End With
With txtCustId(0)
    If Len(Trim(.Text)) = 0 Then
        strSendMsg = strSendMsg & vbCrLf & _
                   "A customer Number must be entered" & vbCrLf
                   .SetFocus
    End If
End With

ValidateForm = strSendMsg

End Function


