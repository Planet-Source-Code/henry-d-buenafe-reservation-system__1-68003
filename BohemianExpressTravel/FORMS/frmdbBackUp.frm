VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBBackup 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Database"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox b8Line2 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   11115
      TabIndex        =   5
      Top             =   540
      Width           =   11115
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "&Create Backup"
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
      Height          =   375
      Left            =   2790
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   1800
      Width           =   1395
   End
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   1
      Top             =   0
      Width           =   10305
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   345
         Left            =   570
         TabIndex        =   2
         Top             =   90
         Width           =   2445
      End
   End
   Begin MSComctlLib.ProgressBar progStat 
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   990
      Visible         =   0   'False
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox b8Line1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   11115
      TabIndex        =   7
      Top             =   1650
      Width           =   11115
   End
   Begin VB.Label lblCBK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Creating Backup File..."
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   780
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "frmDBBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim m_AutoBackup As Boolean
Public WithEvents clsBKU As clsHuffman
Attribute clsBKU.VB_VarHelpID = -1

Public Function ShowForm(Optional ByVal bAutoBackup As Boolean = False)
    
    m_AutoBackup = bAutoBackup
    
    
    If bAutoBackup = False Then
        Me.Show vbModal
    Else
        cmdBackup.Enabled = False
        BackUpDB
    End If
    
End Function


Private Sub clsBKU_EncodeFinish()
    
    progStat.Visible = False
    lblCBK.Visible = False
    DoEvents
    
    If m_AutoBackup = False Then
        MsgBox "HCSIM database backup was successfully created.", vbInformation
    End If
    
    'close this form
    Unload Me
    
End Sub

Private Sub clsBKU_Progress(Procent As Integer)

    progStat.Value = Procent

End Sub

Private Sub cmdBackup_Click()
    cmdBackup.Enabled = False
    BackUpDB
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If m_AutoBackup = False Then
        cmdBackup.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsBKU = Nothing
End Sub



Private Sub BackUpDB()

    Dim FSO As New FileSystemObject
    
    Dim sDBFN As String
    Dim sDBTmpFN As String
    
    If FSO.FolderExists(App.Path & "/Backup") = False Then
        FSO.CreateFolder App.Path & "/Backup"
    End If
    
    'set backup file path filename
    sDBFN = App.Path & "/Backup/" & Format$(Date, "yyyymmdd") & ".bak"
    
    'set temporary file
    sDBTmpFN = sDBFN & Now - DateValue(Now) & GetTickCount
    
    If FSO.FileExists(sDBTmpFN) = True Then
        FSO.DeleteFile sDBTmpFN
    End If
    
    'show ctl
    progStat.Visible = True
    lblCBK.Visible = True
    DoEvents
    
    'start backup
    Set frmDBBackup.clsBKU = New clsHuffman
    frmDBBackup.clsBKU.EncodeFile DBPathFileName, sDBTmpFN
    
    'rename file
    If FSO.FileExists(sDBFN) = True Then
        FSO.DeleteFile sDBFN
    End If
    FSO.MoveFile sDBTmpFN, sDBFN
    
    
    Set FSO = Nothing
    progStat.Visible = False
    lblCBK.Visible = False
End Sub
