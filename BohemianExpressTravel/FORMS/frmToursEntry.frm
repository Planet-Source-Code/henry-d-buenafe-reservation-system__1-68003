VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmToursEntry 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tours Entry"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   Icon            =   "frmToursEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00EDE6E0&
      Enabled         =   0   'False
      Height          =   4305
      Left            =   0
      ScaleHeight     =   4245
      ScaleWidth      =   9915
      TabIndex        =   20
      Top             =   2670
      Width           =   9975
      Begin VB.CommandButton cmdAccts 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   8895
         TabIndex        =   22
         Top             =   1020
         Width           =   915
      End
      Begin VB.CommandButton cmdAccts 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   8895
         TabIndex        =   21
         Top             =   630
         Width           =   915
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2790
         Left            =   15
         TabIndex        =   23
         Top             =   420
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   4921
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Customer ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Address"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Phone"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Due"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   5655
         TabIndex        =   37
         Top             =   3885
         Width           =   1050
      End
      Begin VB.Label lblAmountDue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7425
         TabIndex        =   36
         Top             =   3915
         Width           =   1395
      End
      Begin VB.Label txtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7425
         TabIndex        =   30
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Tour Group Members"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   105
         TabIndex        =   27
         Top             =   105
         Width           =   1785
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Total No. Of Persons"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   5640
         TabIndex        =   26
         Top             =   3270
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   5655
         TabIndex        =   25
         Top             =   3585
         Width           =   1155
      End
      Begin VB.Label lblTotalPerson 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7425
         TabIndex        =   24
         Top             =   3285
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00EDE6E0&
      Enabled         =   0   'False
      Height          =   1860
      Left            =   0
      ScaleHeight     =   1800
      ScaleWidth      =   9915
      TabIndex        =   5
      Top             =   810
      Width           =   9975
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4155
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   33
         Top             =   405
         Width           =   2700
      End
      Begin VB.TextBox txtTourPackageAmt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   31
         Top             =   735
         Width           =   1560
      End
      Begin VB.TextBox txtTourGuideName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4140
         Locked          =   -1  'True
         MaxLength       =   120
         TabIndex        =   29
         Top             =   1395
         Width           =   2700
      End
      Begin VB.CommandButton cmdLookupGuide 
         Height          =   315
         Left            =   3765
         Picture         =   "frmToursEntry.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1395
         Width           =   360
      End
      Begin VB.CommandButton cmdLookUpTourPackage 
         Height          =   315
         Left            =   3780
         Picture         =   "frmToursEntry.frx":124C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   405
         Width           =   360
      End
      Begin VB.TextBox txtDestination 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1065
         Width           =   4635
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   8
         Top             =   405
         Width           =   1560
      End
      Begin VB.TextBox txtTourGuideID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   120
         TabIndex        =   7
         Top             =   1395
         Width           =   1545
      End
      Begin VB.TextBox txtTourNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7995
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   6
         Top             =   45
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker dtTourDate 
         Height          =   315
         Left            =   7995
         TabIndex        =   11
         Top             =   405
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53870593
         CurrentDate     =   38353
      End
      Begin MSComCtl2.DTPicker dtTransDate 
         Height          =   315
         Left            =   2205
         TabIndex        =   18
         Top             =   75
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53870593
         CurrentDate     =   38353
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tour Package Amount"
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
         Left            =   135
         TabIndex        =   32
         Top             =   735
         Width           =   1875
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans Date"
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
         Left            =   135
         TabIndex        =   19
         Top             =   105
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tour Date"
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
         Index           =   8
         Left            =   6930
         TabIndex        =   17
         Top             =   435
         Width           =   840
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tour No"
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
         Index           =   7
         Left            =   6900
         TabIndex        =   16
         Top             =   60
         Width           =   645
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tour Package"
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
         Index           =   4
         Left            =   150
         TabIndex        =   15
         Top             =   405
         Width           =   1155
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
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
         Index           =   5
         Left            =   150
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tour Guide"
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
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
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
         Index           =   6
         Left            =   150
         TabIndex        =   12
         Top             =   2355
         Width           =   690
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      BackColor       =   &H00C4BFB7&
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleMode       =   0  'User
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton cmdButton 
         Cancel          =   -1  'True
         Caption         =   "&Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   3450
         Picture         =   "frmToursEntry.frx":17CE
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   15
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Browse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   2595
         Picture         =   "frmToursEntry.frx":19F0
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   15
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   15
         Picture         =   "frmToursEntry.frx":1D2E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   15
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   885
         Picture         =   "frmToursEntry.frx":2038
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   15
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   1740
         Picture         =   "frmToursEntry.frx":2342
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   15
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   4305
         Picture         =   "frmToursEntry.frx":264C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   15
         Width           =   855
      End
      Begin VB.Label lblLogo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5640
         TabIndex        =   38
         Top             =   255
         Width           =   4050
      End
   End
End
Attribute VB_Name = "frmToursEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents frm As frmSelectTourPackage
Attribute frm.VB_VarHelpID = -1
Private WithEvents frmSelTG As frmSelectTourGuide
Attribute frmSelTG.VB_VarHelpID = -1
Private WithEvents frmSelCustomer As frmSelectCustomer2
Attribute frmSelCustomer.VB_VarHelpID = -1
Private WithEvents frmBrowse As frmBrowseTours
Attribute frmBrowse.VB_VarHelpID = -1
Private WithEvents frmPay As frmPayments
Attribute frmPay.VB_VarHelpID = -1

Dim cnMain As ADODB.Connection
Dim rsMaster As ADODB.Recordset
Dim rsDetails As ADODB.Recordset

Dim lNewRec As Boolean


Private Function GetLastRefNo() As Long
    Dim sSQL As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
                
    sSQL = "SELECT sysLastNo.LastNo From sysLastNo " & _
           "WHERE (((sysLastNo.Key)='TOURNO')) "
        
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        GetLastRefNo = rs("LastNo").Value + 1
    Else
        GetLastRefNo = 1
    End If
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
End Function

Private Sub IncrementRefNo()
On Error GoTo HandleError
    Dim sSQL As String
    
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    If lNewRec = False Then
        Exit Sub
    End If
    
    sSQL = "SELECT sysLastNo.* From sysLastNo " & _
           "WHERE sysLastNo.Key = 'TOURNO' "
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = cConnect
    cn.Open
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    
     If Not (rs.BOF And rs.EOF) Then
            rs("LastNo") = rs("LastNo") + 1
        Else
            rs.AddNew
            rs("Key") = "TOURNO"
            rs("LastNo") = 1
        End If
        
        rs.Update
        
        rs.Close
        cn.Close
        Set rs = Nothing
        Set cn = Nothing

    Exit Sub
HandleError:
    MsgBox Error
End Sub

Private Sub CalcTotalAmt()
    Dim pTotalAmount As Double
    Dim ptotalPerson As Integer
    Dim pTourAmount As Double
    
    pTourAmount = CDbl(txtTourPackageAmt.Text)
    ptotalPerson = ListView1.ListItems.Count
    pTotalAmount = ptotalPerson * pTourAmount
    
    lblTotalPerson.Caption = ptotalPerson

    
    txtAmount.Caption = pTotalAmount

End Sub

Private Sub ClearEntries()
    lNewRec = False
    ListView1.ListItems.Clear
    dtTransDate.Value = Date
    txtCode.Text = ""
    txtTourNo.Text = ""
    txtDescription.Text = ""
    dtTourDate.Value = Date
    txtTourPackageAmt.Text = ""
    txtDestination.Text = ""
    txtTourGuideID.Text = ""
    txtTourGuideName.Text = ""
    
    lblTotalPerson.Caption = 0
    txtAmount.Caption = 0
    
    Picture1.Enabled = False
    Picture2.Enabled = False
End Sub

Private Function ValidEntry() As Boolean
    If txtCode.Text = "" Then
        MsgBox "Please enter a valid tenant.", vbOKOnly, "Empty"
        txtCode.SetFocus
        Exit Function
    End If
    If txtReference.Text = "" Then
        MsgBox "Please enter a valid reference.", vbOKOnly, "Empty"
        txtReference.SetFocus
        Exit Function
    End If
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Please enter valid items.", vbOKOnly, "Empty"
        Exit Function
    End If
    If Not (dtDate.Value >= pvarDateStart And dtDate.Value <= pvarDateEnd) Then
        MsgBox "Date must be between " & pvarDateStart & " and " & pvarDateEnd & ".", vbOKOnly, "Invalid date"
        dtDate.SetFocus
        Exit Function
    End If
    ValidEntry = True
End Function


Private Sub cmdAccts_Click(Index As Integer)
    Select Case Index
        Case 0
            If txtCode.Text = "" Then
                MsgBox "Enter Tour Package.", vbOKOnly, "Bohemian Travel"
                Exit Sub
            Else
                frmSelCustomer.LoadList
                frmSelCustomer.Show vbModal
            End If
        Case 2
            ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
            CalcTotalAmt
    End Select
End Sub

Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case 0
                ClearEntries
                Picture1.Enabled = True
                Picture2.Enabled = True
                cmdButton(0).Enabled = False
                cmdButton(3).Enabled = False
                cmdButton(6).Enabled = True
                lNewRec = True
                dtTourDate.Value = Date
                dtTransDate.Value = Date
                txtTourNo.Text = Format(GetLastRefNo(), "00000000")
        Case 1
            frmBrowse.LoadList
            frmBrowse.Show vbModal
                
        Case 2
            'Dim frmPay As New frmPayments
            
            If txtTourNo.Text = "" Then
                MsgBox "Please Select Transaction.", vbOKOnly, "Empty"
                Exit Sub
                
            ElseIf lblAmountDue.Caption = 0 Then
                    MsgBox "Transaction is Fully Paid.", vbOKOnly, "Empty"
                    Exit Sub
            Else
                frmPay.txtTourNo.Text = txtTourNo
                frmPay.txtReceipt.Text = Format(frmPay.GetLastRefNo(), "00000000")
                frmPay.Show vbModal
            End If
                
        Case 3
            Picture1.Enabled = True
            Picture2.Enabled = True
            cmdButton(0).Enabled = False
            cmdButton(3).Enabled = False
            cmdButton(6).Enabled = True
            
        Case 6
            If ListView1.ListItems.Count = 0 Then
                MsgBox "Please enter valid items.", vbOKOnly, "Empty"
                Exit Sub
            'Else
            End If
            If MsgBox("Are you sure you want to save this transaction?", vbYesNo) = vbYes Then
                SaveEntries
                SaveDetails
                ClearEntries
                cmdButton(0).Enabled = True
     '           cmdButton(3).Enabled = False
                cmdButton(6).Enabled = False
                cmdButton(7).Enabled = True
            'End If
            End If
        Case 7
            Unload Me
    End Select
End Sub

Private Sub cmdLookupGuide_Click()
    
    frmSelTG.LoadList
    frmSelTG.Show vbModal
    
End Sub

Private Sub cmdLookUpTourPackage_Click()
    frm.LoadList
    frm.Show vbModal
End Sub

Private Sub Form_Activate()
    Me.Refresh
End Sub

Private Sub Form_Initialize()
    Set frm = New frmSelectTourPackage
    Set frmSelTG = New frmSelectTourGuide
    Set frmSelCustomer = New frmSelectCustomer2
    Set frmBrowse = New frmBrowseTours
    Set frmPay = New frmPayments
End Sub

Private Sub Form_Terminate()
    Set frm = Nothing
    Set frmSelTG = Nothing
    Set frmSelCustomer = Nothing
    Set frmBrowse = Nothing
    Set frmPay = Nothing
End Sub

Private Sub frm_OnSelect()
    txtCode.Text = frm.pPackageID
    txtTourPackageAmt.Text = frm.pPackageAmount
    txtDescription.Text = frm.pDescription
    txtDestination.Text = frm.pDesctination
End Sub

Private Sub frmBrowse_OnSelect()
    txtTourNo.Text = frmBrowse.pTourNo
    dtTransDate.Value = frmBrowse.pTransdate
    dtTourDate.Value = frmBrowse.pTourDate
    txtCode.Text = frmBrowse.pTourPackageID
    txtDescription.Text = frmBrowse.pTourPackageName
    txtTourPackageAmt.Text = frmBrowse.pTourPackageAmount
    txtDestination.Text = frmBrowse.pDestination
    txtTourGuideID.Text = frmBrowse.pTourGuideID
    txtTourGuideName.Text = frmBrowse.pName
    lblTotalPerson.Caption = frmBrowse.pNoPerson
    txtAmount.Caption = frmBrowse.pAmount
    
    FillListView
    
    CalcAmountDue
    
    cmdButton(3).Enabled = True
End Sub


Private Sub frmPayment_OnPay()
    CalcAmountDue
End Sub

Private Sub frmPay_OnPay()
    CalcAmountDue
End Sub

Private Sub frmSelCustomer_OnSelect()
    Dim li As ListItem
                
    Set li = ListView1.ListItems.Add(, , frmSelCustomer.pCustomerNo & "")
    li.ListSubItems.Add , , frmSelCustomer.pCustomerName & ""
    li.ListSubItems.Add , , frmSelCustomer.pAddress & ""
    li.ListSubItems.Add , , frmSelCustomer.pPhone & ""
    
    li.Tag = 0 'nID
    CalcTotalAmt
    
End Sub

Private Sub frmSelTG_OnSelect()
    txtTourGuideID.Text = frmSelTG.pEmployeeID
    txtTourGuideName.Text = frmSelTG.pName
    'CalcTotalAmt
End Sub

Private Sub SaveEntries()
    'On Error GoTo SaveErr
'    Dim cn As ADODB.Connection
'    Dim rs As ADODB.Recordset
    
    Dim sSQL As String
    
    sSQL = "SELECT Tours.* " & _
            "From Tours " & _
            "WHERE (((Tours.TourNo)='" & txtTourNo.Text & "'))"
    
    Set cnMain = New ADODB.Connection
    Set rsMaster = New ADODB.Recordset
    
    cnMain.ConnectionString = cConnect
    cnMain.Open
    
    rsMaster.Open sSQL, cnMain, adOpenDynamic, adLockOptimistic
        
    If lNewRec = True Then
    
        With rsMaster
            If .RecordCount > 0 Then
                MsgBox "ID Number already in use!", vbInformation, "Warning"
                Exit Sub
            Else
                .AddNew
                .Fields("TourNo") = txtTourNo.Text
                .Fields("TransDate") = dtTransDate.Value
                .Fields("TourDate") = dtTourDate.Value
                .Fields("TourPackageID") = txtCode.Text
                .Fields("Amount") = txtAmount.Caption
                .Fields("Destination") = txtDestination.Text
                .Fields("NoPerson") = lblTotalPerson.Caption
                .Fields("TourGuideID") = txtTourGuideID.Text
                '.Fields("Type") = 2
                .Update
                IncrementRefNo
            End If
        End With
    Else
            With rsMaster
                '.Fields("TourNo") = txtTourNo.Text
                .Fields("TransDate") = dtTransDate.Value
                .Fields("TourDate") = dtTourDate.Value
                .Fields("TourPackageID") = txtCode.Text
                .Fields("Amount") = txtAmount.Caption
                .Fields("Destination") = txtDestination.Text
                .Fields("NoPerson") = lblTotalPerson.Caption
                .Fields("TourGuideID") = txtTourGuideID.Text
                '.Fields("Type") = 2
                .Update
            End With
    End If
    
    rsMaster.Close
    cnMain.Close
    
    Set rsMaster = Nothing
    Set cnMain = Nothing
End Sub

Private Sub SaveDetails()
    Dim lvw As ListView
    Dim i As Long
    
    Dim sSQL As String
    Dim uSQL As String
    
    sSQL = "SELECT TourMembers.* " & _
            "FROM TourMembers "
    
    Set lvw = ListView1
    
    Set cnMain = New ADODB.Connection
    Set rsDetails = New ADODB.Recordset
    
    cnMain.ConnectionString = cConnect
    cnMain.Open
        
    rsDetails.Open sSQL, cnMain, adOpenDynamic, adLockOptimistic
        
    For i = 1 To lvw.ListItems.Count
        With lvw.ListItems(i)
            If lvw.ListItems(i).Tag = 0 Then
                rsDetails.AddNew
                rsDetails.Fields("TourNo") = txtTourNo.Text
                rsDetails.Fields("CustomerNo") = .Text
                rsDetails.Fields("CustomerName") = .ListSubItems(1).Text
                rsDetails.Fields("Address") = .ListSubItems(2).Text
                rsDetails.Fields("Phone") = .ListSubItems(3).Text
            Else
                uSQL = "SELECT TourMembers.* " & _
                        "From TourMembers " & _
                        "WHERE (((TourMembers.ID)= " & .Tag & "))"

                Dim rs As ADODB.Recordset
                Dim cn As ADODB.Connection
                
                Set cn = New ADODB.Connection
                Set rs = New ADODB.Recordset
                
                cn.ConnectionString = cConnect
                cn.Open
                
                rs.Open uSQL, cn, adOpenDynamic, adLockOptimistic
                
                rs.Fields("TourNo") = txtTourNo.Text
                rs.Fields("CustomerNo") = .Text
                rs.Fields("CustomerName") = .ListSubItems(1).Text
                rs.Fields("Address") = .ListSubItems(2).Text
                rs.Fields("Phone") = .ListSubItems(3).Text
                rs.Update
                
                rs.Close
                cn.Close
                
                Set rs = Nothing
                Set cn = Nothing
            End If
        End With
    Next
        rsDetails.Update
        rsDetails.Close
        cnMain.Close
            
        Set rsDetails = Nothing
        Set cnMain = Nothing
End Sub


Private Sub FillListView()
    Dim li As ListItem
    Dim LV As ListView
    Dim sSQL As String
            
    sSQL = "SELECT TourMembers.* " & _
            "From TourMembers " & _
            "WHERE (((TourMembers.TourNo)='" & txtTourNo.Text & "'))"
                            
    Set cnMain = New ADODB.Connection
    Set rsDetails = New ADODB.Recordset
    
    cnMain.Open cConnect
    rsDetails.Open sSQL, cnMain
    
    Set LV = ListView1
    LV.ListItems.Clear
    
    If (rsDetails.RecordCount <> 0) Or (rsDetails.RecordCount = -1) Then
        rsDetails.MoveFirst
        Do While Not rsDetails.EOF
            Set li = LV.ListItems.Add(, , rsDetails("CustomerName") & "")
            li.ListSubItems.Add , , rsDetails("CustomerNo") & ""
            li.ListSubItems.Add , , rsDetails("Address") & ""
            li.ListSubItems.Add , , rsDetails("Phone") & ""
            li.Tag = rsDetails("ID")
            rsDetails.MoveNext
        Loop
    End If
    rsDetails.Close
    cnMain.Close
End Sub

Private Sub CalcAmountDue()
    Dim sSQL As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
        
    Dim pSumOfAmount As Double
    Dim pAmountDue As Double
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    Dim rec As Integer
    
    sSQL = "SELECT Payments.TourNo, Sum(Payments.Amount) AS SumOfAmount " & _
            "From Payments " & _
            "GROUP BY Payments.TourNo " & _
            "HAVING (((Payments.TourNo)='" & txtTourNo.Text & "'))"
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn, adOpenStatic, adLockBatchOptimistic
    
    rec = rs.RecordCount
    'MsgBox rs.RecordCount
    
    If rs.RecordCount > 0 Then
        pSumOfAmount = rs.Fields("SumOfAmount")
        pAmountDue = CDbl(txtAmount.Caption) - pSumOfAmount
        lblAmountDue.Caption = pAmountDue
    Else
        lblAmountDue.Caption = txtAmount.Caption
    End If
    
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
End Sub
Private Sub txtCode_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub txtTourNo_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

