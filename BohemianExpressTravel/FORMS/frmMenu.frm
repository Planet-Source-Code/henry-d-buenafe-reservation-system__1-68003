VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "s"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblReservation 
      BackStyle       =   0  'Transparent
      Caption         =   "Reservation Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   1050
      TabIndex        =   16
      Top             =   4755
      Width           =   3045
   End
   Begin VB.Label lblpackage 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Package"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   1020
      TabIndex        =   15
      Top             =   4230
      Width           =   2655
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      Visible         =   0   'False
      X1              =   6420
      X2              =   6420
      Y1              =   3990
      Y2              =   4860
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      Visible         =   0   'False
      X1              =   6420
      X2              =   6675
      Y1              =   4395
      Y2              =   4395
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Visible         =   0   'False
      X1              =   6420
      X2              =   6675
      Y1              =   4860
      Y2              =   4860
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Visible         =   0   'False
      X1              =   1785
      X2              =   6435
      Y1              =   3975
      Y2              =   3975
   End
   Begin VB.Label lblMonthlySales 
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Sales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   6825
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDaySales 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales of the day"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   6825
      TabIndex        =   13
      Top             =   4185
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Visible         =   0   'False
      X1              =   6420
      X2              =   6675
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      Visible         =   0   'False
      X1              =   6420
      X2              =   6675
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      Visible         =   0   'False
      X1              =   6420
      X2              =   6420
      Y1              =   2580
      Y2              =   3450
   End
   Begin VB.Line line1 
      BorderColor     =   &H80000009&
      Visible         =   0   'False
      X1              =   1785
      X2              =   6435
      Y1              =   3450
      Y2              =   3450
   End
   Begin VB.Label lblIndividual 
      BackStyle       =   0  'Transparent
      Caption         =   "Individual Tour Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   6840
      TabIndex        =   12
      Top             =   2865
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Label lblGroup 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Tour Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   6810
      TabIndex        =   11
      Top             =   2385
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   240
      Left            =   10680
      TabIndex        =   10
      Top             =   8715
      Width           =   1245
   End
   Begin VB.Label lblSales 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   1020
      TabIndex        =   9
      Top             =   3735
      Width           =   690
   End
   Begin VB.Label lblTour 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   1005
      TabIndex        =   8
      Top             =   3255
      Width           =   780
   End
   Begin VB.Label lblClient 
      BackStyle       =   0  'Transparent
      Caption         =   "Client Account Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   990
      TabIndex        =   7
      Top             =   2715
      Width           =   2505
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gopel Group of Companies"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   240
      Left            =   105
      TabIndex        =   6
      Top             =   8745
      Width           =   2430
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bohemian Express Travel"
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
      Height          =   390
      Left            =   75
      TabIndex        =   5
      Top             =   8460
      Width           =   3660
   End
   Begin VB.Label lblTransaction 
      BackStyle       =   0  'Transparent
      Caption         =   "Incharge Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   990
      TabIndex        =   4
      Top             =   2190
      Width           =   2295
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "User Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   975
      TabIndex        =   3
      Top             =   1695
      Width           =   1770
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reservation System v1.0"
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
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   11655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   315
      TabIndex        =   1
      Top             =   825
      Width           =   1425
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   10410
      TabIndex        =   0
      Top             =   8070
      Width           =   1320
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
lbldate.Caption = Date
End Sub



Private Sub lblClient_Click()
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    lblGroup.Visible = False
   ' lblSingle.Visible = False
    lblIndividual.Visible = False
    
    Line5.Visible = False
    Line6.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    lblDaySales.Visible = False
    lblMonthlySales.Visible = False

    frmClientAcct.Show vbModal
End Sub

Private Sub lblExit_Click()
    If MsgBox("Are you sure you want to exit this transaction?", vbYesNo) = vbYes Then
    Unload Me
    End If
End Sub

Private Sub lblGroup_Click()
    GroupTour
    frmTour.Caption = "GROUP TOUR"
End Sub

Private Sub GroupTour()
On Error Resume Next
Dim li As ListItem
Dim LV As ListView

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    frmTour.lblSpecific.Caption = "GROUP"
    .Open "SELECT * FROM dbTour WHERE Specific='" & frmTour.lblSpecific & "'", cConnect, adOpenDynamic, adLockOptimistic
        
        Set LV = frmTour.ListView1
        LV.ListItems.Clear
     
        If .RecordCount <> 0 Then
        .MoveFirst
        Do While Not .EOF
    
        Set li = LV.ListItems.Add(, , rs!TourDate & "")
        li.ListSubItems.Add , , rs!Destination & ""
        li.ListSubItems.Add , , rs!TourNo & ""
        li.ListSubItems.Add , , rs!ID & ""
        li.ListSubItems.Add , , rs!NoPerson & ""
        li.ListSubItems.Add , , rs!Incharge & ""
        li.Tag = rs!ID
        .MoveNext
                
    Loop
    End If
    frmTour.Show vbModal
End With
End Sub

Private Sub lblIndividual_Click()
    IndividualTour
End Sub

Private Sub IndividualTour()
On Error Resume Next
Dim li As ListItem
Dim LV As ListView

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    frmTour.lblSpecific.Caption = "INDIVIDUAL"
    .Open "SELECT * FROM dbTour WHERE Specific='" & frmTour.lblSpecific & "'", cConnect, adOpenDynamic, adLockOptimistic
        
        Set LV = frmTour.ListView1
        LV.ListItems.Clear
     
        If .RecordCount <> 0 Then
        .MoveFirst
        Do While Not .EOF
    
        Set li = LV.ListItems.Add(, , rs!TourDate & "")
        li.ListSubItems.Add , , rs!Destination & ""
        li.ListSubItems.Add , , rs!TourNo & ""
        li.ListSubItems.Add , , rs!ID & ""
        li.ListSubItems.Add , , rs!NoPerson & ""
        li.ListSubItems.Add , , rs!Incharge & ""
        li.Tag = rs!ID
        .MoveNext
    Loop
    End If
    frmTour.Show vbModal
End With
End Sub
Private Sub lblpackage_Click()
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    lblGroup.Visible = False
  '  lblSingle.Visible = False
    lblIndividual.Visible = False
    
    Line5.Visible = False
    Line6.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    lblDaySales.Visible = False
    lblMonthlySales.Visible = False

    frmTransaction.Show vbModal
End Sub

Private Sub lblReservation_Click()
    frmReservation.Show vbModal
End Sub

Private Sub lblSales_Click()
    Line5.Visible = True
    Line6.Visible = True
    Line7.Visible = True
    Line8.Visible = True
    lblDaySales.Visible = True
    lblMonthlySales.Visible = True
End Sub

Private Sub lblTour_Click()
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    lblGroup.Visible = True
'    lblSingle.Visible = True
    lblIndividual.Visible = True
End Sub

Private Sub lblTransaction_Click()
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    lblGroup.Visible = False
   ' lblSingle.Visible = False
    lblIndividual.Visible = False
    
    Line5.Visible = False
    Line6.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    lblDaySales.Visible = False
    lblMonthlySales.Visible = False
    
    frmPilotAcct.Show vbModal
End Sub

Private Sub lblUser_Click()
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    lblGroup.Visible = False
  ' lblSingle.Visible = False
    lblIndividual.Visible = False
    
    Line5.Visible = False
    Line6.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    lblDaySales.Visible = False
    lblMonthlySales.Visible = False
    
    frmUserAcct.Show vbModal
End Sub


