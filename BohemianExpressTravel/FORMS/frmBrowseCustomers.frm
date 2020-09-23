VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrowseCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7350
      TabIndex        =   13
      Top             =   765
      Width           =   1215
   End
   Begin VB.CommandButton btnselect 
      Height          =   585
      Left            =   120
      Picture         =   "frmBrowseCustomers.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   165
      Width           =   660
   End
   Begin VB.CommandButton btnNew 
      Height          =   585
      Left            =   825
      Picture         =   "frmBrowseCustomers.frx":0186
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   165
      Width           =   660
   End
   Begin VB.CommandButton btnRefresh 
      Height          =   585
      Left            =   1530
      Picture         =   "frmBrowseCustomers.frx":02FC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   165
      Width           =   660
   End
   Begin VB.CommandButton btnsearch 
      Height          =   585
      Left            =   2235
      Picture         =   "frmBrowseCustomers.frx":0421
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   165
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
      Left            =   6165
      TabIndex        =   1
      Top             =   390
      Width           =   2415
   End
   Begin VB.CommandButton btnClose 
      Height          =   585
      Left            =   2940
      Picture         =   "frmBrowseCustomers.frx":075F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   165
      Width           =   660
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4350
      Left            =   105
      TabIndex        =   6
      Top             =   1245
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   7673
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Address"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Contact No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Relative Name"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contact"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Incharge"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "UserID"
         Object.Width           =   2540
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
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   90
      TabIndex        =   12
      Top             =   825
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
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   765
      TabIndex        =   11
      Top             =   840
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
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   1500
      TabIndex        =   10
      Top             =   840
      Width           =   705
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
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   2205
      TabIndex        =   9
      Top             =   840
      Width           =   705
   End
   Begin VB.Label lblsearch 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search ID No."
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
      Height          =   300
      Left            =   4530
      TabIndex        =   8
      Top             =   420
      Width           =   1485
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
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   2880
      TabIndex        =   7
      Top             =   840
      Width           =   705
   End
End
Attribute VB_Name = "frmBrowseCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBrowse As ADODB.Recordset
Dim dbConnection As ADODB.Connection

Dim frm As New frmCustomers

Public Sub LoadList()
    Dim li As ListItem
    Dim LV As ListView
    
    Dim sSQL As String
    
    sSQL = "SELECT dbCustomers.* " & _
            "From dbCustomers " & _
            "WHERE (((dbCustomers.CustomerName) Like '" & txtSearch.Text & "%'))"

    On Error Resume Next
    
    dbConnection.ConnectionString = cConnect
    dbConnection.Open
    
    rsBrowse.Open sSQL, dbConnection, adOpenDynamic, adLockOptimistic

    Set LV = ListView1
    LV.ListItems.Clear
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "Customer Name", 5000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Code", 1600, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Active", 1000, lvwColumnLeft

    With rsBrowse
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = LV.ListItems.Add(, , rsBrowse("CustomerName") & "")
                li.ListSubItems.Add , , rsBrowse("CustomerNo") & ""
                li.ListSubItems.Add , , rsBrowse("Active") & ""
                li.Tag = rsBrowse("CustomerNo")
                .MoveNext
            Loop
            .Close
            dbConnection.Close

        Else
            .Close
            dbConnection.Close
            Exit Sub
        End If
    End With
   
End Sub

Private Sub Form_Initialize()
    Set dbConnection = New ADODB.Connection
    Set rsBrowse = New ADODB.Recordset
End Sub

Private Sub Form_Terminate()
    Set dbConnection = Nothing
    Set rsBrowse = Nothing
End Sub

Private Sub ListView1_DblClick()
    With ListView1.ListItems(ListView1.SelectedItem.Index)
        frm.SelectRecord (.Tag)
        frm.Show
        
    End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        LoadList
    End If
End Sub
