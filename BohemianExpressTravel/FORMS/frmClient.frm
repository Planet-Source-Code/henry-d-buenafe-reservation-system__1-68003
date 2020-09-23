VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnOk 
      Caption         =   "Exit![Enter]"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1050
      TabIndex        =   2
      Top             =   480
      Width           =   5820
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5640
      Left            =   75
      TabIndex        =   1
      Top             =   1080
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   9948
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmClient.frx":0000
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   90
      Picture         =   "frmClient.frx":008B
      Top             =   270
      Width           =   720
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
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
      ForeColor       =   &H80000010&
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   7425
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
End
End Sub

