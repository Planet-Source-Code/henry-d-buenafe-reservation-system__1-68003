VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About "
   ClientHeight    =   6000
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6975
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4141.307
   ScaleMode       =   0  'User
   ScaleWidth      =   6549.886
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6105
      Top             =   6015
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   75
      ScaleHeight     =   4545
      ScaleWidth      =   6765
      TabIndex        =   3
      Top             =   690
      Width           =   6825
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3210
         Left            =   255
         Picture         =   "frmAbout.frx":0742
         ScaleHeight     =   3210
         ScaleWidth      =   2460
         TabIndex        =   6
         Top             =   240
         Width           =   2460
      End
      Begin RichTextLib.RichTextBox txtMessage 
         Height          =   4095
         Left            =   2985
         TabIndex        =   4
         Top             =   210
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   7223
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmAbout.frx":BC3A
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
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5640
      TabIndex        =   0
      Top             =   5610
      Width           =   1260
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Boggyman TM  2006 - 2007  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   5685
      Width           =   3045
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Reservation System Version 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   375
      Width           =   6690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bohemian Express Travel "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   30
      Width           =   6780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   6465.371
      Y1              =   3799.649
      Y2              =   3799.649
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   84.515
      X2              =   6465.371
      Y1              =   3810.002
      Y2              =   3810.002
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String
Dim i

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
txtMessage.FileName = "\BohemianExpressTravel\About.rtf"
str1 = txtMessage.Text
i = 0
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 100
i = i + 1
txtMessage.Text = Left(str1, i)
If i = Len(str1) Then
i = 1
Timer1.Interval = 3000
End If
End Sub
