VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmployeesLookup 
   BackColor       =   &H00EDE6E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employees Lookup"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "frmEmployeesLookup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDE6E0&
      Caption         =   "Search Opotions:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5520
      TabIndex        =   6
      Top             =   0
      Width           =   3225
      Begin VB.ComboBox cbpSearchOptions 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   3060
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4305
      TabIndex        =   5
      Top             =   75
      Width           =   1095
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EDE6E0&
      Caption         =   "&Find for:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4215
      Begin VB.ComboBox txbFilterBy 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   3975
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EDE6E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8805
      TabIndex        =   0
      Top             =   5040
      Width           =   8805
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6315
         TabIndex        =   9
         Top             =   30
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   5055
         TabIndex        =   2
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7575
         TabIndex        =   1
         Top             =   30
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4380
      Left            =   -15
      TabIndex        =   8
      Top             =   705
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7726
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "SmallImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee Name"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID Number"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Activation"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":08CA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":15A4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":23F6
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":3248
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":3B22
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":43FC
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":4CD6
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":56A0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":5F7A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":6294
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":6B6E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":7448
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":7D22
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":803C
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":8916
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":91F0
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":9ACA
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":A3A4
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":AC7E
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":B558
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":BE32
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":C70C
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":CFE6
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":D8C0
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":E19A
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":EA74
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":F34E
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":FC28
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":10502
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":10DDC
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":11692
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":11F6C
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":123BE
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":12810
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":14FC2
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":16244
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeesLookup.frx":1655E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEmployeesLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim cnCustomers As ADODB.Connection
'Dim rsCustomers As ADODB.Recordset

Private Sub cbpSearchOptions_Change()
    If ListView1.ListItems.Count < 1 Then Exit Sub
    Call search_in_listview(ListView1, cbpSearchOptions.Text)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
    frmEmployees.ClearEntries
    frmEmployees.cmdButton(0).Enabled = False
    frmEmployees.cmdButton(3).Enabled = False
    frmEmployees.cmdButton(4).Enabled = False
    frmEmployees.cmdButton(6).Enabled = True
    frmEmployees.Picture1.Enabled = True
    frmEmployees.lNewRec = True
    
    frmEmployees.Show vbModal
End Sub

Private Sub cmdRefresh_Click()
 txbFilterBy.Text = ""
 LoadList
End Sub
Private Sub cmdSelect_Click()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT Employees.* " & _
            "From Employees " & _
            "WHERE (((Employees.EmployeeID)='" & ListView1.SelectedItem.Tag & "'))"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn ', adOpenDynamic, adLockOptimistic
    'rs.Requery
   
    frmEmployees.txtCode = rs("EmployeeID")
    frmEmployees.txtName = IIf(IsNull(rs("Name")), "", rs("Name"))
    frmEmployees.txtAddress = IIf(IsNull(rs("Address")), "", rs("Address"))
    frmEmployees.txtCity = IIf(IsNull(rs("City")), "", rs("City"))
    frmEmployees.txtPhoneNo = IIf(IsNull(rs("HomePhone")), "", rs("HomePhone"))
    frmEmployees.cbActive = IIf(IsNull(rs("Active")), "", rs("Active"))
    
    frmEmployees.cmdButton(3).Enabled = True
    frmEmployees.cmdButton(4).Enabled = True
    
    rs.Close
    cn.Close
    
    frmEmployees.Show vbModal
End Sub

Private Sub Form_Activate()
    LoadList
End Sub


Public Sub LoadList()
    On Error Resume Next
    Dim li As ListItem
    Dim LV As ListView
    Dim vData As String
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
 '   Dim rec As Integer
    Dim sSQL  As String
    
    sSQL = "SELECT Employees.* " & _
            "From Employees " & _
            "WHERE (((Employees.Name) Like '" & txbFilterBy.Text & "%'))"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
  
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
            
    Set LV = ListView1
    LV.ListItems.Clear

    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
   '     rec = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("Name") & "")
            li.SmallIcon = 19
            li.ListSubItems.Add , , rs("EmployeeID") & ""
            li.ListSubItems.Add , , rs("Active") & ""
            li.Tag = rs("EmployeeID")
            rs.MoveNext
        Loop
        
    End If
    
    rs.Close
    cn.Close
    
End Sub
Private Sub ListView1_DblClick()
    cmdSelect_Click
End Sub
Private Sub txbFilterBy_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
   If KeyCode = vbKeyReturn Then
   LoadList
End If
End Sub


