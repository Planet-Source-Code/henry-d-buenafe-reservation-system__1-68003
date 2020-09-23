VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTour 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTour.frx":0000
   ScaleHeight     =   4860
   ScaleWidth      =   7905
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Height          =   585
      Left            =   3090
      Picture         =   "frmTour.frx":39F50
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   570
      Width           =   660
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5460
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.CommandButton btnsearch 
      Height          =   585
      Left            =   2355
      Picture         =   "frmTour.frx":3A392
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   570
      Width           =   660
   End
   Begin VB.CommandButton btnRefresh 
      Height          =   585
      Left            =   1620
      Picture         =   "frmTour.frx":3A6D0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   570
      Width           =   660
   End
   Begin VB.CommandButton btnNew 
      Height          =   585
      Left            =   885
      Picture         =   "frmTour.frx":3A7F5
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   570
      Width           =   660
   End
   Begin VB.CommandButton btnselect 
      Height          =   585
      Left            =   165
      Picture         =   "frmTour.frx":3A96B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   570
      Width           =   660
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3045
      Left            =   135
      TabIndex        =   3
      Top             =   1545
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   5371
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tour Date"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Distination"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tour No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "No. of Client"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Incharge"
         Object.Width           =   2540
      EndProperty
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
      Left            =   3060
      TabIndex        =   15
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label lblsearch 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search Package No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   3825
      TabIndex        =   14
      Top             =   945
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Left            =   2325
      TabIndex        =   13
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
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
      Left            =   1590
      TabIndex        =   12
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label Label1 
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
      Height          =   165
      Left            =   870
      TabIndex        =   11
      Top             =   1200
      Width           =   705
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
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   225
      Left            =   6060
      TabIndex        =   2
      Top             =   4590
      Width           =   1650
   End
   Begin VB.Label lblSpecific 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Specific"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   6090
      TabIndex        =   1
      Top             =   1170
      Width           =   1650
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Express Travel Tour"
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
      Height          =   450
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   7725
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   7020
      Picture         =   "frmTour.frx":3AAF1
      Top             =   405
      Width           =   720
   End
End
Attribute VB_Name = "frmTour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnNew_Click()
    frmTourIN.Show vbModal
End Sub

Private Sub btnRefresh_Click()
    Requery
End Sub
Private Sub Requery()
'On Error Resume Next
Dim li As ListItem
Dim LV As ListView

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs

    .Open "SELECT * FROM dbTour Where Specific='" & lblSpecific & "'", _
     cConnect, adOpenDynamic, adLockOptimistic
        
        Set LV = ListView1
        LV.ListItems.Clear
     
        If .RecordCount <> 0 Then
        .MoveFirst
        Do While Not .EOF
    
        Set li = LV.ListItems.Add(, , rs!TourDate & "")
        li.ListSubItems.Add , , rs!Destination & ""
        li.ListSubItems.Add , , rs!TourNo & ""
        li.ListSubItems.Add , , rs!ID & ""
        li.ListSubItems.Add , , rs!NoPerson & ""
        li.ListSubItems.Add , , rs!Incharge & ""
        li.Tag = rs!ID
        .MoveNext
                
    Loop
    End If
 End With
End Sub
Private Sub btnselect_Click()
With ListView1.ListItems(ListView1.SelectedItem.Index)
    frmTourIN.txtDate.Text = .Text
    frmTourIN.txtDestination.Text = .ListSubItems(1)
    frmTourIN.txtTour.Text = .ListSubItems(2)
    frmTourIN.txtID.Text = .ListSubItems(3)
    frmTourIN.txtNo.Text = .ListSubItems(4)
    frmTourIN.txtIncharge.Text = .ListSubItems(5)
          
    frmTourIN.btnSave.Enabled = False
    frmTourIN.btnNew.Enabled = False
    
    frmTourIN.btnDelete.Enabled = True
    frmTourIN.btnEdit.Enabled = True
    frmTourIN.btnAdd.Enabled = True
    frmTourIN.Show vbModal
    End With
End Sub
Private Sub Form_Load()
    lblDate.Caption = Date
    txtSearch.MaxLength = 6
End Sub

Private Sub ListView1_DblClick()
btnselect_Click
End Sub

Private Sub txtSearch_Change()
    If ListView1.ListItems.Count < 1 Then Exit Sub
    Call search_in_listview(ListView1, txtSearch.Text)
End Sub
