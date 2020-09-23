VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Package"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTransaction.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   7470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnselect 
      Height          =   585
      Left            =   165
      Picture         =   "frmTransaction.frx":39F50
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   570
      Width           =   660
   End
   Begin VB.CommandButton btnNew 
      Height          =   585
      Left            =   900
      Picture         =   "frmTransaction.frx":3A0D6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   570
      Width           =   660
   End
   Begin VB.CommandButton btnRefresh 
      Height          =   585
      Left            =   1635
      Picture         =   "frmTransaction.frx":3A24C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   570
      Width           =   660
   End
   Begin VB.CommandButton btnsearch 
      Height          =   585
      Left            =   2370
      Picture         =   "frmTransaction.frx":3A371
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Left            =   5115
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.CommandButton btnClose 
      Height          =   585
      Left            =   3105
      Picture         =   "frmTransaction.frx":3A6AF
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   570
      Width           =   660
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3525
      Left            =   135
      TabIndex        =   6
      Top             =   1560
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   6218
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Particular Packages "
         Object.Width           =   14993
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "ID Number"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Package"
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
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7380
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
      Left            =   135
      TabIndex        =   12
      Top             =   1215
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
      Left            =   885
      TabIndex        =   11
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
      Left            =   1605
      TabIndex        =   10
      Top             =   1200
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   6615
      Picture         =   "frmTransaction.frx":3AAF1
      Top             =   420
      Width           =   720
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
      Left            =   2340
      TabIndex        =   9
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
      Left            =   3705
      TabIndex        =   8
      Top             =   945
      Visible         =   0   'False
      Width           =   1440
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
      TabIndex        =   7
      Top             =   1215
      Width           =   705
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnNew_Click()
    frmTransactionPackage.Show vbModal
End Sub

Private Sub btnRefresh_Click()
Packages
End Sub

Private Sub btnsearch_Click()
    txtSearch.Visible = True
    lblsearch.Visible = True
End Sub

Private Sub btnselect_Click()
 With ListView1.ListItems(ListView1.SelectedItem.Index)
        frmTransactionPackage.txtParticular.Text = .Text
        frmTransactionPackage.txtamount.Text = .ListSubItems(1)
        frmTransactionPackage.txtID.Text = .ListSubItems(2)
          
        frmTransactionPackage.btnSave.Enabled = False
        frmTransactionPackage.btnDelete.Enabled = True
        frmTransactionPackage.btnEdit.Enabled = True
        frmTransactionPackage.Show vbModal
    End With
End Sub

Private Sub Form_Load()
Packages
End Sub

Private Sub Packages()
On Error Resume Next
Dim li As ListItem
Dim LV As ListView

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    
    .Open "SELECT * FROM dbPackage", cConnect, adOpenDynamic, adLockOptimistic
        
        Set LV = ListView1
        LV.ListItems.Clear
     
    If .RecordCount <> 0 Then
    .MoveFirst
    Do While Not .EOF
    
        Set li = LV.ListItems.Add(, , rs!Particular & "")
        li.ListSubItems.Add , , Format(rs("Amount"), "#,##0.00") & ""
        li.ListSubItems.Add , , rs!IDNo & ""
        li.Tag = rs!IDNo
        .MoveNext
                
    Loop
    End If
End With
End Sub

Private Sub ListView1_DblClick()
    btnselect_Click
End Sub
Private Sub txtSearch_Change()
 If ListView1.ListItems.Count < 1 Then Exit Sub
    Call search_in_listview(ListView1, txtSearch.Text)
End Sub
Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
