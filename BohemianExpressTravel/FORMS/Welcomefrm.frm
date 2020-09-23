VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   90
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   225
      Picture         =   "Welcomefrm.frx":0000
      ScaleHeight     =   1425
      ScaleWidth      =   1545
      TabIndex        =   7
      Top             =   420
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   75
      Top             =   45
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -1680
      Top             =   -75
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Exit Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4005
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   2
      Top             =   1995
      Width           =   1530
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0C000&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   1
      Top             =   2010
      Width           =   1515
   End
   Begin VB.Image image1 
      Height          =   60
      Left            =   0
      Picture         =   "Welcomefrm.frx":1366
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   6975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cashering System Version 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   2295
      TabIndex        =   6
      Tag             =   "App Description"
      Top             =   1035
      Width           =   3375
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "CREATOR/DEVELOPER:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   2400
      TabIndex        =   5
      Tag             =   "App Description"
      Top             =   1530
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer: Xavier "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   150
      Left            =   2415
      TabIndex        =   4
      Tag             =   "App Description"
      Top             =   1710
      Width           =   3165
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2005-2006"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   180
      Left            =   285
      TabIndex        =   3
      Top             =   1875
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME Bohemian Express"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   990
      Left            =   2385
      TabIndex        =   0
      Top             =   75
      Width           =   3030
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
    fMainForm.lblLogo.Caption = "Bohemian Express Travel Agency" 'Logo
    Unload Me 'close splash form
End Sub

Private Sub cmd2_Click()
    End 'Close/End Program.
End Sub

Private Sub Timer2_Timer()
 If image1.Left > 5225 Then
    image1.Left = -3945
    Else
    image1.Left = image1.Left + 100
    End If
End Sub
