VERSION 5.00
Begin VB.Form frmTourIN 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3480
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTourIN.frx":0000
   ScaleHeight     =   3480
   ScaleWidth      =   6885
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnAdd 
      Enabled         =   0   'False
      Height          =   525
      Left            =   120
      Picture         =   "frmTourIN.frx":39F50
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   60
      Width           =   2070
   End
   Begin VB.ComboBox cbDestination 
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
      Left            =   5010
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   825
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.TextBox txtIncharge 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1785
      TabIndex        =   27
      Top             =   2895
      Width           =   5040
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5010
      TabIndex        =   26
      Top             =   2565
      Width           =   1815
   End
   Begin VB.TextBox txtSpecific 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1785
      TabIndex        =   25
      Top             =   2565
      Width           =   1815
   End
   Begin VB.TextBox txtPackage 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5010
      TabIndex        =   24
      Top             =   2235
      Width           =   1815
   End
   Begin VB.TextBox txtNo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1785
      TabIndex        =   23
      Top             =   2235
      Width           =   1815
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5010
      TabIndex        =   22
      Top             =   1905
      Width           =   1815
   End
   Begin VB.TextBox txtDestination 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1785
      TabIndex        =   21
      Top             =   1905
      Width           =   1815
   End
   Begin VB.TextBox txtTour 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5010
      TabIndex        =   20
      Top             =   1575
      Width           =   1815
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1785
      TabIndex        =   19
      Top             =   1575
      Width           =   1815
   End
   Begin VB.CommandButton btnExit 
      Height          =   525
      Left            =   3705
      Picture         =   "frmTourIN.frx":3A115
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   630
      Width           =   660
   End
   Begin VB.CommandButton btnEdit 
      Enabled         =   0   'False
      Height          =   525
      Left            =   2985
      Picture         =   "frmTourIN.frx":3A557
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   630
      Width           =   660
   End
   Begin VB.CommandButton btnUpdate 
      Enabled         =   0   'False
      Height          =   525
      Left            =   2250
      Picture         =   "frmTourIN.frx":3A747
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   630
      Width           =   660
   End
   Begin VB.CommandButton btnDelete 
      Enabled         =   0   'False
      Height          =   525
      Left            =   1530
      Picture         =   "frmTourIN.frx":3A969
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   630
      Width           =   660
   End
   Begin VB.CommandButton btnSave 
      Enabled         =   0   'False
      Height          =   525
      Left            =   810
      Picture         =   "frmTourIN.frx":3AD5A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   630
      Width           =   660
   End
   Begin VB.CommandButton btnNew 
      Height          =   525
      Left            =   120
      Picture         =   "frmTourIN.frx":3AE55
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   630
      Width           =   630
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Add Client"
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
      Height          =   195
      Left            =   2250
      TabIndex        =   34
      Top             =   420
      Width           =   810
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOUR"
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
      Left            =   90
      TabIndex        =   33
      Top             =   0
      Width           =   6720
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   225
      Left            =   5040
      TabIndex        =   30
      Top             =   3240
      Width           =   1785
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Incharge"
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
      Left            =   105
      TabIndex        =   29
      Top             =   2970
      Width           =   1350
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Left            =   3720
      TabIndex        =   28
      Top             =   2655
      Width           =   1350
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Specific"
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
      TabIndex        =   18
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Package Amt."
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
      Left            =   3720
      TabIndex        =   17
      Top             =   2325
      Width           =   1350
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Client"
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
      Left            =   105
      TabIndex        =   16
      Top             =   2310
      Width           =   1350
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Tour"
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
      Left            =   3735
      TabIndex        =   15
      Top             =   2010
      Width           =   1350
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Number"
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
      Left            =   3720
      TabIndex        =   14
      Top             =   1650
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Distination Package"
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
      Left            =   105
      TabIndex        =   13
      Top             =   1980
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Number"
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
      TabIndex        =   12
      Top             =   1650
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   6105
      Picture         =   "frmTourIN.frx":3B0BD
      Top             =   30
      Width           =   720
   End
   Begin VB.Label Label9 
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
      Left            =   3645
      TabIndex        =   11
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
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
      Left            =   2955
      TabIndex        =   10
      Top             =   1215
      Width           =   705
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
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
      Left            =   2250
      TabIndex        =   9
      Top             =   1215
      Width           =   705
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
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
      Left            =   1545
      TabIndex        =   8
      Top             =   1215
      Width           =   705
   End
   Begin VB.Label Label5 
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
      Left            =   720
      TabIndex        =   7
      Top             =   1215
      Width           =   855
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New"
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
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   1215
      Width           =   810
   End
End
Attribute VB_Name = "frmTourIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnAdd_Click()
On Error Resume Next
    If lblDate > txtDate Then
    MsgBox "Error Transaction Date!", vbInformation, "Date"
    Else
        AddThis
        frmAddClient.Show vbModal
    End If
End Sub
Private Sub AddThis()
    frmAddClient.lblTourNo.Caption = frmTourIN.txtTour.Text
    frmAddClient.lblDestination1.Caption = frmTourIN.txtDestination.Text
    frmAddClient.lblDate.Caption = frmTourIN.txtDate.Text
    frmAddClient.lblpackageAmount.Caption = frmTourIN.txtPackage.Text
End Sub

Private Sub btnDelete_Click()
If MsgBox("Are you sure you want to delete this account?", vbYesNo) = vbYes Then
 Delete
  btnDelete.Enabled = False
  btnEdit.Enabled = False
  btnNew.Enabled = True
  ClearEntries
  End If
End Sub
Private Sub Delete()
On Error GoTo DelErr
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
With rs
    .Open "SELECT *FROM dbTour WHERE ID='" & txtID.Text & "'", _
     cConnect, adOpenDynamic, adLockPessimistic
    .Delete
Exit Sub
DelErr:
    MsgBox Err.Description
End With
End Sub
Private Sub btnEdit_Click()
txtDate.Enabled = True
txtIncharge.Enabled = True
    
    btnUpdate.Enabled = True
    btnEdit.Enabled = True
    
    btnDelete.Enabled = False
    btnNew.Enabled = False
    btnSave.Enabled = False
End Sub

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnNew_Click()
    ClearEntries
    btnSave.Enabled = True
    btnNew.Enabled = False
    btnDelete.Enabled = False
    
    txtTour.Text = txtTour + 1
    txtNo.Text = "0"
    txtSpecific.Text = "Individual"
    
End Sub
Private Sub ClearEntries()
txtID.Text = ""
txtID.Enabled = True
txtNo.Text = ""
txtNo.Enabled = True
txtIncharge.Text = ""
txtIncharge.Enabled = True
txtDate.Text = ""
txtDate.Enabled = True
txtDestination.Text = ""
txtDestination.Enabled = True
txtPackage.Text = ""
txtamount.Text = ""
End Sub
Private Sub ChargeAmount()
'On Error Resume Next

Dim li As ListItem
Dim LV As ListView

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    
    .Open "SELECT * FROM dbPackage Where IDNo='" & txtDestination.Text & "'", cConnect
     'adOpenDynamic , adLockOptimistic
        
     Set txtPackage.DataSource = rs
     txtPackage.DataField = "Amount"
     
End With
End Sub

Private Sub btnSave_Click()
If MsgBox("Are you sure you want to save this account?", vbYesNo) = vbYes Then
    SaveTransaction
     OR_Number
End If
End Sub
Private Sub OR_Number()
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    .Open "Select *FROM dbTourNo", cConnect, adOpenStatic, adLockOptimistic
     
    .Fields(1) = txtTour.Text
    .Update
End With
End Sub
Private Sub SaveTransaction()
On Error GoTo SaveErr
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
With rs
    .Open "SELECT * FROM dbTour WHERE ID='" & txtID.Text & "'", cConnect, adOpenStatic, adLockOptimistic
    
    If .RecordCount <> 0 Then
        MsgBox "ID Number already in use!", vbInformation, "Warning"
        btnSave.Enabled = True
    Else
        .AddNew
        .Fields("ID") = txtID.Text
        .Fields("TourNo") = txtTour.Text
        .Fields("Amount") = txtamount.Text
        .Fields("TourDate") = txtDate.Text
        .Fields("Destination") = txtDestination.Text
        .Fields("NoPerson") = txtNo.Text
        .Fields("Specific") = txtSpecific.Text
        .Fields("Incharge") = txtIncharge.Text
        .Update
            
            btnNew.Enabled = True
            btnAdd.Enabled = True
            btnSave.Enabled = False
            
    End If
    Exit Sub
SaveErr:
    MsgBox Err.Description
End With
End Sub

Private Sub btnUpdate_Click()
If MsgBox("Are you sure you want to Update this account?", vbYesNo) = vbYes Then
    UpdateFile
    btnExit.Enabled = True
End If
End Sub
Private Sub UpdateFile()
On Error GoTo UpdateErr
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
With rs

    .Open "SELECT * From dbTour where ID='" & txtID.Text & "'", cConnect, adOpenStatic, adLockOptimistic
         
    rs("TourDate") = IIf(txtDate.Text = "", Null, txtDate.Text)
    rs("Incharge") = IIf(txtIncharge.Text = "", Null, txtIncharge.Text)
    rs("NoPerson") = IIf(txtNo.Text = "", Null, txtNo.Text)
    rs("Specific") = IIf(txtSpecific.Text = "", Null, txtSpecific.Text)
    .Update
            MsgBox "Update successfully!", vbInformation, ""
    Unload Me
    Exit Sub
UpdateErr:
    MsgBox "Save this transaction first"
End With
End Sub
Private Sub cbDestination_Change()
  txtDestination.Text = cbDestination.Text
End Sub

Private Sub cbDestination_Click()
txtDestination.Text = cbDestination.Text
End Sub

Private Sub Form_Load()
    Max_length
    TourPackage
    TourNo
    lblDate.Caption = Date
End Sub
Private Sub Max_length()
    txtID.MaxLength = 6
    txtTour.MaxLength = 6
End Sub
Private Sub TourNo()
Set rs = New ADODB.Recordset
    rs.Open "Select dbTourNo.* From dbTourNo", cConnect
    txtTour.Text = rs.Fields(1)
 End Sub
Private Sub TourPackage()
Dim vntemp As Variant
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

  With rs
    .Open "SELECT * FROM dbPackage", cConnect, _
     adOpenDynamic, adLockOptimistic
  End With
    Do While Not rs.EOF
     vntemp = rs!IDNo
      If IsNull(vntemp) Then vntemp = ""
       cbDestination.AddItem CStr(vntemp)
        rs.MoveNext
    Loop
 End Sub

Private Sub txtDate_LostFocus()
    If lblDate > txtDate Then
        MsgBox "Enter Specific Date", vbInformation, "Date"
    End If
End Sub

Private Sub txtDestination_Change()
    ChargeAmount
End Sub

Private Sub txtDestination_Click()
    cbDestination.Visible = True
    cbDestination.SetFocus
    'txtDestination.Enabled = False
End Sub
Private Sub txtIncharge_Click()
    frmIncharge.Show vbModal
End Sub

Private Sub txtNo_Change()
On Error Resume Next
   Calculate
   
   If txtNo.Text = 1 Then
    txtSpecific.Text = "INDIVIDUAL"
    Else
     If txtNo.Text > 1 Then
        txtSpecific.Text = "GROUP"
     End If
     End If
End Sub
Private Sub Calculate()
    txtamount.Text = Val(txtPackage.Text) * Val(txtNo.Text)
End Sub
Private Sub txtNo_Click()
  '  frmAddClient.Show vbModal
End Sub
