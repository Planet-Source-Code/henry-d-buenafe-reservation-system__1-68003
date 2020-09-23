VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIncharge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tour Incharge"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIncharge.frx":0000
   ScaleHeight     =   4815
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnselect 
      Height          =   585
      Left            =   75
      Picture         =   "frmIncharge.frx":39F50
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Width           =   660
   End
   Begin VB.CommandButton btnClose 
      Height          =   585
      Left            =   750
      Picture         =   "frmIncharge.frx":3A0D6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   660
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3930
      Left            =   60
      TabIndex        =   0
      Top             =   810
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   6932
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID No"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select"
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
      Left            =   45
      TabIndex        =   4
      Top             =   630
      Width           =   705
   End
   Begin VB.Label Label4 
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
      Left            =   720
      TabIndex        =   3
      Top             =   630
      Width           =   705
   End
End
Attribute VB_Name = "frmIncharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    Unload Me
End Sub
Private Sub btnselect_Click()
On Error Resume Next
 With ListView1.ListItems(ListView1.SelectedItem.Index)
    frmTourIN.txtIncharge.Text = .Text
    Unload Me
    End With
End Sub

Private Sub Form_Load()
    Incharge
End Sub
Private Sub Incharge()
'On Error Resume Next
Dim li As ListItem
Dim LV As ListView
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
With rs
    .Open "SELECT * FROM dbPilot", cConnect
    ', adOpenDynamic, adLockOptimistic
     Set LV = ListView1
     LV.ListItems.Clear
     If .RecordCount <> 0 Then
    .MoveFirst
     Do While Not .EOF
    
        Set li = LV.ListItems.Add(, , rs!Name & "")
        li.ListSubItems.Add , , rs!UserID & ""
        li.Tag = rs!UserID
        .MoveNext
                
    Loop
    End If
End With
End Sub
Private Sub ListView1_DblClick()
btnselect_Click
End Sub
