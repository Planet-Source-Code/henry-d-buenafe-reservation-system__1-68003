VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTransactionPackage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Package"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTransactionPackage.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   6540
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1170
      TabIndex        =   16
      Top             =   3465
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox txtParticular 
      Height          =   1545
      Left            =   1185
      TabIndex        =   15
      Top             =   1875
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   2725
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmTransactionPackage.frx":39F50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1170
      TabIndex        =   13
      Top             =   1530
      Width           =   1815
   End
   Begin VB.CommandButton btnNew 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   135
      Picture         =   "frmTransactionPackage.frx":39FCB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   630
      Width           =   630
   End
   Begin VB.CommandButton btnSave 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   825
      Picture         =   "frmTransactionPackage.frx":3A233
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   630
      Width           =   660
   End
   Begin VB.CommandButton btnDelete 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1545
      Picture         =   "frmTransactionPackage.frx":3A32E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   630
      Width           =   660
   End
   Begin VB.CommandButton btnUpdate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2265
      Picture         =   "frmTransactionPackage.frx":3A71F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   630
      Width           =   660
   End
   Begin VB.CommandButton btnEdit 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2985
      Picture         =   "frmTransactionPackage.frx":3A941
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   630
      Width           =   660
   End
   Begin VB.CommandButton btnExit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3705
      Picture         =   "frmTransactionPackage.frx":3AB31
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   630
      Width           =   660
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   135
      TabIndex        =   18
      Top             =   3495
      Width           =   1350
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Particular"
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   135
      TabIndex        =   17
      Top             =   1890
      Width           =   1350
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Number"
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   135
      TabIndex        =   14
      Top             =   1575
      Width           =   1350
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Packages"
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
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6495
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
      TabIndex        =   11
      Top             =   1215
      Width           =   810
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
      TabIndex        =   10
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
      Left            =   1545
      TabIndex        =   9
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
      TabIndex        =   8
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
      Left            =   2955
      TabIndex        =   7
      Top             =   1215
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
      Left            =   3645
      TabIndex        =   6
      Top             =   1200
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   5730
      Picture         =   "frmTransactionPackage.frx":3AF73
      Top             =   15
      Width           =   720
   End
End
Attribute VB_Name = "frmTransactionPackage"
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
    .Open "SELECT *FROM dbPackage WHERE IDNo='" & txtID.Text & "'", _
     cConnect, adOpenDynamic, adLockPessimistic
    .Delete
Exit Sub
DelErr:
    MsgBox Err.Description
End With
End Sub

Private Sub btnEdit_Click()
    txtamount.Enabled = True
    txtParticular.Enabled = True
    
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
'Count_Rec
    btnSave.Enabled = True
    btnNew.Enabled = False
    btnDelete.Enabled = False
End Sub
Private Sub ClearEntries()
txtID.Text = ""
txtParticular.Text = ""
txtamount.Text = "0.00"

txtID.Enabled = True
txtParticular.Enabled = True
txtamount.Enabled = True

End Sub

Private Sub btnSave_Click()
If MsgBox("Are you sure you want to save this account?", vbYesNo) = vbYes Then
    SaveTransaction
    btnSave.Enabled = False
    btnNew.Enabled = True
    ClearEntries
End If
End Sub
Private Sub SaveTransaction()
On Error GoTo SaveErr
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
With rs
    .Open "SELECT * FROM dbPackage WHERE IDNo='" & txtID.Text & "'", cConnect, adOpenStatic, adLockOptimistic
    
    If .RecordCount <> 0 Then
        MsgBox "ID Number already in use!", vbInformation, "Warning"
    Else
        .AddNew
        .Fields("IDNo") = txtID.Text
        .Fields("Particular") = txtParticular.Text
        .Fields("Amount") = txtamount.Text
        .Update
    End If
    Exit Sub
SaveErr:
    MsgBox Err.Description
End With
End Sub
Private Sub Count_Rec()
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    .Open "SELECT * FROM dbPackage", cConnect, adOpenStatic, adLockOptimistic
    If .RecordCount > 3 Then
    MsgBox "Transaction is full", vbInformation, "Transaction"
    End If
    End With
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

    .Open "SELECT * From dbPackage where IDNo='" & txtID.Text & "'", cConnect, adOpenStatic, adLockOptimistic
         
    rs("IDNo") = IIf(txtID.Text = "", Null, txtID.Text)
    rs("Particular") = IIf(txtParticular.Text = "", Null, txtParticular.Text)
    rs("Amount") = IIf(txtamount.Text = "", Null, txtamount.Text)
    .Update
    
    MsgBox "Update successfully!", vbInformation, ""
    Unload Me
    Exit Sub
UpdateErr:
    MsgBox Err.Description
End With
End Sub
