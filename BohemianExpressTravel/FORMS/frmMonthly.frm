VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMonthly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Sales"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5610
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00EDE6E0&
      Height          =   3630
      Left            =   0
      ScaleHeight     =   3570
      ScaleWidth      =   5550
      TabIndex        =   0
      Top             =   0
      Width           =   5610
      Begin VB.CommandButton btnSelect 
         Caption         =   "&SELECT DATE "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   30
         TabIndex        =   5
         Top             =   2715
         Width           =   5505
      End
      Begin MSComCtl2.MonthView mv1 
         Height          =   2370
         Left            =   30
         TabIndex        =   1
         Top             =   330
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   18284545
         CurrentDate     =   39116
      End
      Begin MSComCtl2.MonthView mv2 
         Height          =   2370
         Left            =   2835
         TabIndex        =   2
         Top             =   330
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   18284545
         CurrentDate     =   39116
      End
      Begin VB.Label lblTo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4260
         TabIndex        =   7
         Top             =   45
         Width           =   1245
      End
      Begin VB.Label lblfrom 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1545
         TabIndex        =   6
         Top             =   45
         Width           =   1140
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Date To:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2865
         TabIndex        =   4
         Top             =   75
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Date Frm:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   3
         Top             =   90
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add Microsoft Windows Control 2.6 component to attached Month View Calendar.

Private Sub btnSelect_Click()
SearchDATE
End Sub
Private Sub SearchDATE()
On Error GoTo SearchErr
Dim sSQL  As String
Dim rpt As New rptDailySales

        sSQL = "SELECT Payments.* " & _
            "From Payments " & _
            "WHERE (((Payments.PaymentDate) Between #" & lblfrom & "# And #" & lblTo & "#))"
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
        
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
    
    If rs.RecordCount = 0 Then
    rs.Close
    cn.Close
    MsgBox "No transaction found!", vbInformation

    Else

    Set rpt.DataSource = rs
    rpt.Title = "Bohemian Express Travel Sales Report"
    rpt.WindowState = vbMaximized
    rpt.Show vbModal
           
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

    End If
    Exit Sub
SearchErr:
    MsgBox "No Transaction found!", vbCritical, "Warning"
End Sub
Private Sub mv1_DateClick(ByVal DateClicked As Date)
lblfrom = mv1
End Sub

Private Sub mv2_DateClick(ByVal DateClicked As Date)
lblTo = mv2
End Sub
