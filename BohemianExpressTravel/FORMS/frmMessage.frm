VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMessage 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Message"
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7800
      Top             =   -60
   End
   Begin RichTextLib.RichTextBox txtMessage 
      Height          =   6990
      Left            =   135
      TabIndex        =   1
      Top             =   345
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   12330
      _Version        =   393217
      BackColor       =   16119285
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMessage.frx":0000
   End
   Begin VB.Label lblStop 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   7560
      TabIndex        =   2
      Top             =   -30
      Width           =   1815
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Message from the author"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   5520
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String
Dim i As String
Private Sub btnOk_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
txtMessage.FileName = "\BohemianExpressTravel\Read Me.rtf"
str1 = txtMessage.Text
i = 0
End Sub

Private Sub lblStop_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 100
i = i + 1
txtMessage.Text = Left(str1, i)
If i = Len(str1) Then
i = 1
Timer1.Interval = 3000
Unload Me
End If
'Unload Me
End Sub

