VERSION 5.00
Begin VB.Form frmAddClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Transaction"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAddClient.frx":0000
   ScaleHeight     =   4185
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnPrint 
      Height          =   525
      Left            =   1440
      Picture         =   "frmAddClient.frx":39F50
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   645
      Width           =   660
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4845
      TabIndex        =   11
      Top             =   1200
      Width           =   1410
   End
   Begin VB.CommandButton btnPayment 
      Height          =   525
      Left            =   765
      Picture         =   "frmAddClient.frx":3A13F
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   645
      Width           =   660
   End
   Begin VB.CommandButton btnSave 
      Enabled         =   0   'False
      Height          =   525
      Left            =   90
      Picture         =   "frmAddClient.frx":3A596
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   645
      Width           =   660
   End
   Begin VB.CommandButton btnExit 
      Height          =   525
      Left            =   5580
      Picture         =   "frmAddClient.frx":3A691
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   660
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   165
      Left            =   1395
      TabIndex        =   27
      Top             =   1215
      Width           =   705
   End
   Begin VB.Label lbldate 
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
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   4590
      TabIndex        =   25
      Top             =   2715
      Width           =   1575
   End
   Begin VB.Label lblDestination1 
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
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   1575
      TabIndex        =   24
      Top             =   3135
      Width           =   1575
   End
   Begin VB.Label lblTourNo 
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
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   1545
      TabIndex        =   23
      Top             =   2730
      Width           =   1575
   End
   Begin VB.Label lblpackageAmount 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """P""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   4605
      TabIndex        =   22
      Top             =   3135
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Package Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   3345
      TabIndex        =   21
      Top             =   3135
      Width           =   1350
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   150
      TabIndex        =   20
      Top             =   3120
      Width           =   1380
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   3330
      TabIndex        =   19
      Top             =   2700
      Width           =   1395
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour No:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   135
      TabIndex        =   18
      Top             =   2730
      Width           =   1215
   End
   Begin VB.Label lblIncharge 
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
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   1560
      TabIndex        =   17
      Top             =   2265
      Width           =   4665
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Incharge:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   120
      TabIndex        =   16
      Top             =   2235
      Width           =   1380
   End
   Begin VB.Label lblBalance 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """P""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   4590
      TabIndex        =   15
      Top             =   1905
      Width           =   1575
   End
   Begin VB.Label lblPayment 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """P""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   1560
      TabIndex        =   14
      Top             =   1875
      Width           =   1620
   End
   Begin VB.Label lblClient 
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
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   1560
      TabIndex        =   13
      Top             =   1560
      Width           =   4710
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Click image to search the customer accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   3780
      TabIndex        =   12
      Top             =   465
      Width           =   1635
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   165
      Left            =   5565
      TabIndex        =   10
      Top             =   4005
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   6240
      Y1              =   2595
      Y2              =   2595
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   3360
      TabIndex        =   8
      Top             =   1890
      Width           =   810
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Down Payment:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   135
      TabIndex        =   7
      Top             =   1875
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Client Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   135
      TabIndex        =   6
      Top             =   1545
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Client ID Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   3345
      TabIndex        =   5
      Top             =   1230
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   5520
      Picture         =   "frmAddClient.frx":3AAD3
      Top             =   420
      Width           =   720
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Client Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   300
      Left            =   90
      TabIndex        =   4
      Top             =   0
      Width           =   6150
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   165
      Left            =   90
      TabIndex        =   3
      Top             =   1215
      Width           =   630
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   165
      Left            =   720
      TabIndex        =   2
      Top             =   1215
      Width           =   705
   End
End
Attribute VB_Name = "frmAddClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub ClientAccount()
On Error GoTo BindErr
Dim li As ListItem
Dim LV As ListView

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    
    .Open "SELECT * FROM dbClient Where UserID='" & txtID.Text & "'", cConnect, _
        adOpenDynamic, adLockOptimistic
        
     Set lblClient.DataSource = rs 'Data Binding method
     lblClient.DataField = "name"
     Set lblIncharge.DataSource = rs
     lblIncharge.DataField = "Incharge"
          
     Exit Sub
     End With
BindErr:
    MsgBox Err.Description
End Sub
Private Sub btnPayment_Click()
    frmPayment.Show vbModal
    
    btnSave.Enabled = True
    btnPayment.Enabled = False
End Sub

Private Sub btnSave_Click()
On Error Resume Next
If lblClient = "" Then
    MsgBox "Empty field Client invalid!", vbInformation, "Invalid"
    Exit Sub
End If
If MsgBox("Are you sure you want to save this account?", vbYesNo) = vbYes Then
    Save_User
    
    frmTourIN.txtNo.Text = frmTourIN.txtNo.Text + 1
        
    frmTourIN.btnNew.Enabled = False
    frmTourIN.btnDelete.Enabled = False
    frmTourIN.btnSave.Enabled = False
        
    frmTourIN.btnUpdate.Enabled = True
    frmTourIN.btnEdit.Enabled = True
    frmTourIN.btnExit.Enabled = False
    Unload Me
End If
End Sub
Private Sub Save_User()
On Error GoTo SaveErr
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
With rs
    .Open "SELECT * FROM dbReserve", cConnect, adOpenStatic, adLockOptimistic
    
        .AddNew
        .Fields("TourNo") = lblTourNo
        .Fields("Date") = lbldate
        .Fields("Destination") = lblDestination1
        .Fields("Clientname") = lblClient
        .Fields("ClientID") = txtID.Text
        .Fields("Downpayment") = lblPayment
        .Fields("Balance") = lblBalance
        .Fields("Incharge") = lblIncharge
        .Update
    
    Exit Sub
SaveErr:
    MsgBox Err.Description
End With
End Sub

Private Sub Form_Load()
txtID.MaxLength = 6
End Sub

Private Sub lblPayment_Change()
lblBalance.Caption = Val(lblpackageAmount) - Val(lblPayment)
End Sub
Private Sub txtID_Change()
    ClientAccount
End Sub
