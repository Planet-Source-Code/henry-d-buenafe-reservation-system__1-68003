VERSION 5.00
Begin VB.Form frmPilot 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlane.frx":0000
   ScaleHeight     =   2940
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnSave 
      Enabled         =   0   'False
      Height          =   525
      Left            =   750
      Picture         =   "frmPlane.frx":39F50
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   615
      Width           =   660
   End
   Begin VB.CommandButton btnDelete 
      Enabled         =   0   'False
      Height          =   525
      Left            =   1455
      Picture         =   "frmPlane.frx":3A04B
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   615
      Width           =   660
   End
   Begin VB.CommandButton btnUpdate 
      Enabled         =   0   'False
      Height          =   525
      Left            =   2160
      Picture         =   "frmPlane.frx":3A43C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   615
      Width           =   660
   End
   Begin VB.CommandButton btnEdit 
      Enabled         =   0   'False
      Height          =   525
      Left            =   2865
      Picture         =   "frmPlane.frx":3A65E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   615
      Width           =   660
   End
   Begin VB.CommandButton btnExit 
      Height          =   525
      Left            =   3570
      Picture         =   "frmPlane.frx":3A84E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   615
      Width           =   660
   End
   Begin VB.TextBox txtage 
      Alignment       =   2  'Center
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
      Left            =   3735
      TabIndex        =   11
      Text            =   "0"
      Top             =   2550
      Width           =   570
   End
   Begin VB.TextBox txtcontact 
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
      Left            =   1065
      TabIndex        =   10
      Top             =   2535
      Width           =   1965
   End
   Begin VB.TextBox txtaddress 
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
      Left            =   1065
      TabIndex        =   9
      Top             =   2205
      Width           =   5235
   End
   Begin VB.TextBox txtname 
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
      Left            =   1065
      TabIndex        =   5
      Top             =   1875
      Width           =   5235
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
      Left            =   4665
      TabIndex        =   3
      Top             =   1545
      Width           =   1635
   End
   Begin VB.CommandButton btnNew 
      Height          =   525
      Left            =   75
      Picture         =   "frmPlane.frx":3AC90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   615
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   5565
      Picture         =   "frmPlane.frx":3AEF8
      Top             =   390
      Width           =   720
   End
   Begin VB.Label Label11 
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
      Height          =   165
      Left            =   -30
      TabIndex        =   22
      Top             =   1215
      Width           =   855
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
      Left            =   585
      TabIndex        =   21
      Top             =   1215
      Width           =   855
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
      Left            =   1425
      TabIndex        =   20
      Top             =   1230
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
      Left            =   2130
      TabIndex        =   19
      Top             =   1215
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
      Left            =   2835
      TabIndex        =   18
      Top             =   1230
      Width           =   705
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
      Left            =   3540
      TabIndex        =   17
      Top             =   1230
      Width           =   705
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
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
      Left            =   3270
      TabIndex        =   8
      Top             =   2580
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
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
      Left            =   45
      TabIndex        =   7
      Top             =   2655
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   45
      TabIndex        =   6
      Top             =   2310
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Name"
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
      Left            =   15
      TabIndex        =   4
      Top             =   1980
      Width           =   930
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
      Left            =   3615
      TabIndex        =   2
      Top             =   1620
      Width           =   1035
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Incharge"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmPilot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    .Open "SELECT *FROM dbPilot WHERE UserID='" & txtID.Text & "'", _
     cConnect, adOpenDynamic, adLockPessimistic
    .Delete
Exit Sub
DelErr:
    MsgBox Err.Description
End With
End Sub
Private Sub btnEdit_Click()
    txtName.Enabled = True
    txtAddress.Enabled = True
    txtContact.Enabled = True
    txtage.Enabled = True
    
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
    txtage.Text = "23"
    ClearEntries
    btnSave.Enabled = True
    btnNew.Enabled = False
    btnDelete.Enabled = False

End Sub
Private Sub ClearEntries()
    txtName.Text = ""
    txtAddress.Text = ""
    txtContact.Text = ""
    txtID.Text = ""
    
    txtName.Enabled = True
    txtage.Enabled = True
    txtContact.Enabled = True
    txtID.Enabled = True
    txtAddress.Enabled = True

End Sub
Private Sub Save_User()
On Error GoTo SaveErr
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
With rs
    .Open "SELECT * FROM dbPilot WHERE UserID='" & txtID.Text & "'", cConnect, adOpenStatic, adLockOptimistic
    
    If .RecordCount <> 0 Then
        MsgBox "ID Number already in use!", vbInformation, "Warning"
    Else
        .AddNew
        .Fields("UserID") = txtID.Text
        .Fields("name") = txtName.Text
        .Fields("Address") = txtAddress.Text
        .Fields("ContactNo") = txtContact.Text
        .Fields("Age") = txtage.Text
        .Update
    End If
    Exit Sub
SaveErr:
    MsgBox Err.Description
End With
End Sub

Private Sub btnSave_Click()
EmptyField
OverAge
If MsgBox("Are you sure you want to save this account?", vbYesNo) = vbYes Then
    Save_User
    btnSave.Enabled = False
    btnNew.Enabled = True
    ClearEntries
End If
End Sub

Private Sub btnUpdate_Click()
If MsgBox("Are you sure you want to Update this account?", vbYesNo) = vbYes Then
    UpdateFile
End If
End Sub

Private Sub UpdateFile()
On Error GoTo UpdateErr
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
With rs

    .Open "SELECT * From dbPilot where UserID='" & txtID.Text & "'", _
    cConnect, adOpenStatic, adLockOptimistic
         
    rs("UserID") = IIf(txtID.Text = "", Null, txtID.Text)
    rs("name") = IIf(txtName.Text = "", Null, txtName.Text)
    rs("Address") = IIf(txtAddress.Text = "", Null, txtAddress.Text)
    rs("ContactNo") = IIf(txtContact.Text = "", Null, txtContact.Text)
    rs("Age") = IIf(txtage.Text = "", Null, txtage.Text)
    .Update
    
    MsgBox "Update successfully!", vbInformation, ""
    Unload Me
    Exit Sub
UpdateErr:
    MsgBox Err.Description
End With
End Sub

Private Sub Form_Load()
txtID.MaxLength = 6
txtName.MaxLength = 100
txtAddress.MaxLength = 200
txtContact.MaxLength = 13
txtage.MaxLength = 2
End Sub
Private Sub OverAge()
On Error Resume Next
    If txtage.Text > 45 Then
        MsgBox "Over Age,Please follow company about this matter", vbInformation, "Under Age"
            txtage.SetFocus
            SendKeys "{Home}+{End}"
            Exit Sub
    Else
    If txtage.Text < 18 Then
        MsgBox "Under Age,Please follow company about this matter", vbInformation, "Over Age"
            txtage.SetFocus
            SendKeys "{Home}+{End}"
            Exit Sub
            Else
        End If
    End If
End Sub
Private Sub txtage_Change()
'OverAge
End Sub

Private Sub txtage_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtcontact_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub EmptyField()
If is_empty(txtID) = True Then Exit Sub
If is_empty(txtName) = True Then Exit Sub
End Sub
