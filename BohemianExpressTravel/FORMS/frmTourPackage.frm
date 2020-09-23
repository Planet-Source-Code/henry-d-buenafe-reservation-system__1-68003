VERSION 5.00
Begin VB.Form frmTourPackage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tour Package"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmTourPackage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      BackColor       =   &H00C4BFB7&
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleMode       =   0  'User
      ScaleWidth      =   7005
      TabIndex        =   11
      Top             =   0
      Width           =   7065
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
         Left            =   3420
         Picture         =   "frmTourPackage.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
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
         Left            =   2565
         Picture         =   "frmTourPackage.frx":134C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
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
         Left            =   1710
         Picture         =   "frmTourPackage.frx":1656
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
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
         Left            =   855
         Picture         =   "frmTourPackage.frx":1960
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
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
         Left            =   0
         Picture         =   "frmTourPackage.frx":1C6A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00EDE6E0&
      Enabled         =   0   'False
      Height          =   2400
      Left            =   0
      ScaleHeight     =   2340
      ScaleWidth      =   7005
      TabIndex        =   0
      Top             =   810
      Width           =   7065
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   1920
         Width           =   2640
      End
      Begin VB.TextBox txtDetails 
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
         Height          =   795
         Left            =   1650
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1110
         Width           =   5325
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
         Left            =   1650
         MaxLength       =   80
         TabIndex        =   3
         Top             =   780
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
         TabIndex        =   2
         Top             =   120
         Width           =   2610
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
         TabIndex        =   1
         Top             =   450
         Width           =   5325
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         TabIndex        =   10
         Top             =   1965
         Width           =   675
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Details"
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
         TabIndex        =   9
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   840
         Width           =   975
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
         TabIndex        =   7
         Top             =   510
         Width           =   480
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
         TabIndex        =   6
         Top             =   165
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmTourPackage"
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
            txtCode.Enabled = True
            cmdButton(0).Enabled = False
            cmdButton(3).Enabled = False
            cmdButton(4).Enabled = False
            cmdButton(6).Enabled = True
            Picture1.Enabled = True
            lNewRec = True
            txtAmount.Text = "0.00"
        Case 1
            Unload Me
            frmCustomerLookup.Show vbModal
        Case 3
            txtCode.Enabled = False
            cmdButton(0).Enabled = False
            cmdButton(3).Enabled = False
            cmdButton(4).Enabled = False
            cmdButton(6).Enabled = True
            Picture1.Enabled = True
            lNewRec = False
        Case 4
            If MsgBox("Are you sure you want to delete this record(s)?", vbYesNo, "WARNING") = vbYes Then
                DeleteRec
                Unload Me
            End If
        Case 6
            If is_empty(txtCode) And is_empty(txtName) Then 'Look declaration is empty into Public declaration mMain Module.
             'If is_empty(txtName) Then
                Exit Sub
            Else
            If MsgBox("Are you sure you want to save this transaction?", vbYesNo) = vbYes Then
            SaveRecord
            cmdButton(0).Enabled = True
            cmdButton(3).Enabled = True
            cmdButton(4).Enabled = True
            cmdButton(6).Enabled = False
            Picture1.Enabled = False
            
            frmCustomerLookup.LoadList
            frmCustomerLookup.Refresh
            End If
            End If
            'End If
        Case 7
            Unload Me
            
    End Select
End Sub
Public Sub ClearEntries()
    cID = ""
    lNewRec = False
    txtCode.Text = ""
    txtName.Text = ""
    txtDestination.Text = ""
    txtDetails.Text = ""
    txtAmount.Text = ""
    cmdButton(btnSave).Enabled = False
    cmdButton(btnDup).Enabled = False
    cmdButton(btnEdit).Enabled = False
    cmdButton(btnDel).Enabled = False
    'SetEntry False
End Sub

Private Sub SaveRecord()
' On Error GoTo SaveErr
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Dim sSQL As String
    
    sSQL = "SELECT dbTourPackage.* " & _
            "From dbTourPackage " & _
            "WHERE (((dbTourPackage.TourPackageID)='" & txtCode.Text & "'))"
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn, adOpenStatic, adLockOptimistic
        
    If lNewRec = True Then
        With rs
            
            If .RecordCount > 0 Then
                MsgBox "ID Number already in use!", vbInformation, "Warning"
                Exit Sub
            Else
                .AddNew
                .Fields("TourpackageID") = txtCode.Text
                .Fields("Description") = txtName.Text
                .Fields("Destination") = txtDestination.Text
                .Fields("Particular") = txtDetails.Text
                .Fields("Amount") = txtAmount.Text
                .Update
            End If
        End With
    Else
        With rs
            .Fields("TourpackageID") = txtCode.Text
            .Fields("Description") = txtName.Text
            .Fields("Destination") = txtDestination.Text
            .Fields("Particular") = txtDetails.Text
            .Fields("Amount") = txtAmount.Text
            .Update
        End With
    End If
    
    rs.Close
    cn.Close
'SaveErr:
'    MsgBox Err.Description
End Sub

Private Sub DeleteRec()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Dim sSQL As String
    
    sSQL = "SELECT dbTourPackage.* " & _
            "From dbTourPackage " & _
            "WHERE (((dbTourPackage.TourPackageID)='" & txtCode.Text & "'))"
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn, adOpenStatic, adLockOptimistic
    
    rs.Delete
        
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
End Sub

Private Sub Form_Load()
txtCode.MaxLength = 6
txtName.MaxLength = 100
txtDestination.MaxLength = 150
txtDetails.MaxLength = 250
txtAmount.MaxLength = 7
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub txtCode_GotFocus()
On Error Resume Next
    txtAmount.Text = CCur(txtAmount)
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
