VERSION 5.00
Begin VB.Form frmCustomers 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Account"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "frmCustomers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDE6E0&
      Enabled         =   0   'False
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   7095
      TabIndex        =   6
      Top             =   840
      Width           =   7155
      Begin VB.ComboBox cbActive 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   75
         Width           =   1485
      End
      Begin VB.TextBox txtName 
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
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   14
         Top             =   450
         Width           =   5325
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
         Left            =   1650
         MaxLength       =   16
         TabIndex        =   13
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   1650
         MaxLength       =   80
         TabIndex        =   12
         Top             =   780
         Width           =   5325
      End
      Begin VB.TextBox txtContactName 
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
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1110
         Width           =   5325
      End
      Begin VB.TextBox txtContactTitle 
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
         Height          =   360
         Left            =   1650
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1440
         Width           =   5325
      End
      Begin VB.TextBox txtPhoneNo 
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
         Left            =   1650
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1815
         Width           =   2730
      End
      Begin VB.TextBox txtFaxNo 
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
         Left            =   1650
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2145
         Width           =   2730
      End
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00EDE6E0&
         Caption         =   "Active"
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
         Height          =   225
         Left            =   6180
         TabIndex        =   7
         Top             =   1170
         Width           =   1050
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Active/Inactive"
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
         Index           =   7
         Left            =   3945
         TabIndex        =   23
         Top             =   180
         Width           =   1485
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Index           =   0
         Left            =   210
         TabIndex        =   21
         Top             =   195
         Width           =   420
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Index           =   1
         Left            =   210
         TabIndex        =   20
         Top             =   510
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Index           =   2
         Left            =   210
         TabIndex        =   19
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
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
         Left            =   210
         TabIndex        =   18
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Title"
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
         Index           =   4
         Left            =   210
         TabIndex        =   17
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number"
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
         Index           =   5
         Left            =   210
         TabIndex        =   16
         Top             =   1830
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax Number"
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
         Index           =   6
         Left            =   210
         TabIndex        =   15
         Top             =   2160
         Width           =   1005
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      BackColor       =   &H00C4BFB7&
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleMode       =   0  'User
      ScaleWidth      =   7080
      TabIndex        =   0
      Top             =   0
      Width           =   7140
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
         Picture         =   "frmCustomers.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   870
         Picture         =   "frmCustomers.frx":114C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   15
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Delete"
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
         Index           =   4
         Left            =   1725
         Picture         =   "frmCustomers.frx":1456
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
         Left            =   2580
         Picture         =   "frmCustomers.frx":1760
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   15
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Cancel          =   -1  'True
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
         Left            =   3435
         Picture         =   "frmCustomers.frx":1A6A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   15
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cID As String
Public lNewRec As Boolean

Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case 0
            ClearEntries
            cmdButton(0).Enabled = False
            cmdButton(3).Enabled = False
            cmdButton(4).Enabled = False
            cmdButton(6).Enabled = True
            Picture1.Enabled = True
            txtCode.Enabled = True
            lNewRec = True
        Case 3
            cmdButton(0).Enabled = False
            cmdButton(3).Enabled = False
            cmdButton(4).Enabled = False
            cmdButton(6).Enabled = True
            Picture1.Enabled = True
            lNewRec = False
            txtCode.Enabled = False
        Case 4
            If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbYes Then
            Delete_Customer
            cmdButton(0).Enabled = True
            cmdButton(4).Enabled = False
            cmdButton(3).Enabled = False
            End If
        Case 6
            If txtCode.Text = "" Then
                MsgBox "The field is required.Please check it!", vbExclamation, "System version 1.0"
                txtCode.SetFocus
            Exit Sub
            End If
            
            If txtName.Text = "" Then
                MsgBox "The field is required.Please check it!", vbExclamation, "System version 1.0"
                txtName.SetFocus
            Exit Sub
            End If
            
            If cbActive.Text = "" Then
                MsgBox "The field is required.Please check it!", vbExclamation, "System version 1.0"
            Exit Sub
            End If
            
            If MsgBox("Are you sure you want to Save this transaction?", vbYesNo) = vbYes Then
                cmdButton(0).Enabled = True
                cmdButton(3).Enabled = True
                cmdButton(4).Enabled = True
                cmdButton(6).Enabled = False
                Picture1.Enabled = False
                
                SaveRecord
            End If
        Case 7
              Unload Me
              
    End Select
End Sub

Private Sub Delete_Customer()
On Error GoTo DelErr
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim sSQL As String

  sSQL = "SELECT *FROM dbCustomers WHERE CustomerNo='" & txtCode.Text & "'"
  
  Set cn = New ADODB.Connection
  Set rs = New ADODB.Recordset
    
  cn.Open cConnect
  rs.CursorLocation = adUseClient
  rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
  rs.Delete
 
  cn.Close
    
Exit Sub
DelErr:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    txtCode.MaxLength = 6
    txtName.MaxLength = 150
    txtAddress.MaxLength = 250
    txtContactName.MaxLength = 150
    txtContactTitle.MaxLength = 150
    txtPhoneNo.MaxLength = 15
    txtFaxNo.MaxLength = 15
    cbActive.AddItem "Active"
    cbActive.AddItem "Inactive"
End Sub
Public Sub ClearEntries()
    cID = ""
    lNewRec = False
    txtCode.Text = ""
    txtName.Text = ""
    txtAddress.Text = ""
    txtContactName.Text = ""
    txtContactTitle.Text = ""
    txtPhoneNo.Text = ""
    txtFaxNo.Text = ""
End Sub
Private Sub SaveRecord()
On Error GoTo SaveErr
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Dim sSQL As String
    
    sSQL = "SELECT dbCustomers.* " & _
                "From dbCustomers " & _
                "WHERE (((dbCustomers.Customerno)='" & txtCode.Text & "'))"
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
        
    If lNewRec = True Then
        With rs
                .AddNew
                .Fields("Customerno") = txtCode.Text
                .Fields("CustomerName") = txtName.Text
                .Fields("Address") = txtAddress.Text
                .Fields("ContactName") = txtContactName.Text
                .Fields("ContactTitle") = txtContactTitle.Text
                .Fields("PhoneNo") = txtPhoneNo.Text
                .Fields("FaxNo") = txtFaxNo.Text
                .Fields("Active") = cbActive.Text
                .Update
        End With
    Else
        With rs
            .Fields("CustomerName") = txtName.Text
            .Fields("Address") = txtAddress.Text
            .Fields("ContactName") = txtContactName.Text
            .Fields("ContactTitle") = txtContactTitle.Text
            .Fields("PhoneNo") = txtPhoneNo.Text
            .Fields("FaxNo") = txtFaxNo.Text
            .Fields("Active") = cbActive.Text
            .Update
        End With
    End If
Exit Sub
    rs.Close
    cn.Close
SaveErr:
    cn.Close
    cmdButton(0).Enabled = False
    cmdButton(3).Enabled = False
    cmdButton(4).Enabled = False
    cmdButton(6).Enabled = True
    Picture1.Enabled = True
    MsgBox Err.Description
End Sub
Private Sub txtFaxNo_Change()
If Not IsNumeric(txtFaxNo.Text) Then
    MsgBox "field required numeric input", vbExclamation
    txtFaxNo.SetFocus
    Exit Sub
    End If
End Sub
Private Sub txtFaxNo_KeyPress(KeyAscii As Integer)
'you can also this code for not accepting numeric input and press text field if you enter
'a non numeric input.
'If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub txtPhoneNo_Change()
If Not IsNumeric(txtPhoneNo.Text) Then
    MsgBox "field required numeric input", vbExclamation
    txtPhoneNo.SetFocus
    Exit Sub
End If
End Sub
