VERSION 5.00
Begin VB.Form frmUser 
   BackColor       =   &H00EDE6E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6075
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6075
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDE6E0&
      Enabled         =   0   'False
      Height          =   1215
      Left            =   15
      ScaleHeight     =   1155
      ScaleWidth      =   6030
      TabIndex        =   6
      Top             =   825
      Width           =   6090
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
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   450
         Width           =   4365
      End
      Begin VB.TextBox txtIDNo 
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
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   1905
      End
      Begin VB.TextBox txtpassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1470
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   780
         Width           =   4365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   120
         TabIndex        =   12
         Top             =   525
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID No."
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
         Left            =   120
         TabIndex        =   11
         Top             =   165
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Password"
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
         Left            =   105
         TabIndex        =   10
         Top             =   870
         Width           =   1245
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      BackColor       =   &H00C4BFB7&
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleMode       =   0  'User
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6075
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
         Picture         =   "frmUser.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmUser.frx":114C
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
         Picture         =   "frmUser.frx":1456
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmUser.frx":1760
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmUser.frx":1A6A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   15
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
            lNewRec = True
        Case 1
            Unload Me
      '      LoadList
      '      frmCustomerLookup.Show vbModal
        Case 3
            If txtName.Text = "Administrator" Then ' Contant Declaration username Administrator
                MsgBox "you cannot Delete/Edit this account", vbCritical, "WARNING"
            Else
            cmdButton(0).Enabled = False
            cmdButton(3).Enabled = False
            cmdButton(4).Enabled = False
            cmdButton(6).Enabled = True
            Picture1.Enabled = True
            UnlockME
            lNewRec = False
            End If
       Case 4
            If txtName.Text = "Administrator" Then ' Contant Declaration username Administrator
                MsgBox "you cannot Delete/Edit this account", vbCritical, "WARNING"
            Exit Sub     'if txtname is true then execute msg. and exit form user.
            Else
            If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbYes Then
            Delete_User
            cmdButton(0).Enabled = True
            cmdButton(4).Enabled = False
            cmdButton(3).Enabled = False
            End If
            End If
        Case 6
            If MsgBox("Are you sure you want to delete this record(s)?", vbYesNo) = vbYes Then
            SaveRecord
            cmdButton(0).Enabled = True
            cmdButton(3).Enabled = True
            cmdButton(4).Enabled = True
            cmdButton(6).Enabled = False
            Picture1.Enabled = False
                    
            frmUserAccount.LoadList
            frmUserAccount.Refresh
            End If
        Case 7
            Unload Me
            frmUserAccount.LoadList
       End Select
End Sub
Public Sub ClearEntries()
    'cID = ""
    lNewRec = False
    txtIDNo.Text = ""
    txtName.Text = ""
    txtpassword.Text = ""
    cmdButton(btnSave).Enabled = False
    cmdButton(btnDup).Enabled = False
    cmdButton(btnEdit).Enabled = False
    cmdButton(btnDel).Enabled = False
    'SetEntry False
End Sub

Public Sub UnlockME()
    txtName.Locked = False
    txtpassword.Locked = False
End Sub

Private Sub Delete_User()
On Error GoTo DelErr
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim sSQL As String
   
  sSQL = "SELECT *FROM dbLogin WHERE UserID='" & txtIDNo.Text & "'"
  
  Set cn = New ADODB.Connection
  Set rs = New ADODB.Recordset

    cn.ConnectionString = cConnect
    cn.Open
    
With rs
    
    .CursorLocation = adUseClient
    .Open sSQL, cn, adOpenDynamic, adLockOptimistic
    .Delete
     cn.Close
End With
Exit Sub
DelErr:
    MsgBox Err.Description
End Sub
Private Sub SaveRecord()
    'On Error GoTo SaveErr
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Dim sSQL As String
    
    sSQL = "SELECT dbLogin.* " & _
                "From dbLogin " & _
                "WHERE (((dbLogin.UserID)='" & txtIDNo.Text & "'))"
   
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
        
    If lNewRec = True Then
        With rs
            
            If .RecordCount > 0 Then
                MsgBox "ID Number already in use!", vbInformation, "Warning"
                Exit Sub
            Else
                .AddNew
                .Fields("UserID") = txtIDNo.Text
                .Fields("Username") = txtName.Text
                .Fields("Password") = txtpassword.Text
                .Update
            End If
        End With
    Else
        With rs
            .Fields("Username") = txtName.Text
            .Fields("Password") = txtpassword.Text
            .Update
        End With
    End If
    
    rs.Close
    cn.Close
        
End Sub
Private Sub Form_Load()
txtIDNo.MaxLength = 10
txtName.MaxLength = 150
txtpassword.MaxLength = 20
End Sub
Private Sub txtpassword_DblClick()
If MsgBox("Are you sure you want to generate password account?", vbYesNo + vbExclamation) = vbYes Then
txtpassword.PasswordChar = ""
End If
End Sub
