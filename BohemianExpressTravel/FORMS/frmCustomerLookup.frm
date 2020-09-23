VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomerLookup 
   BackColor       =   &H00EDE6E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frmCustomerLookup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EDE6E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8880
      TabIndex        =   6
      Top             =   5100
      Width           =   8880
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
         Height          =   510
         Left            =   6345
         TabIndex        =   9
         Top             =   30
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
         Height          =   510
         Left            =   7605
         TabIndex        =   8
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
         Height          =   510
         Left            =   5085
         TabIndex        =   7
         Top             =   30
         Width           =   1215
      End
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
      Left            =   15
      TabIndex        =   4
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
         TabIndex        =   5
         Top             =   240
         Width           =   3975
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
      Height          =   540
      Left            =   4335
      TabIndex        =   3
      Top             =   105
      Width           =   1095
   End
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
      TabIndex        =   0
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
         TabIndex        =   1
         Top             =   240
         Width           =   3060
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4350
      Left            =   -15
      TabIndex        =   2
      Top             =   720
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7673
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
         Text            =   "Customer Name"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer No"
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
            Picture         =   "frmCustomerLookup.frx":0E42
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":1B1C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":296E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":37C0
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":409A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":4974
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":524E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":5C18
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":64F2
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":680C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":70E6
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":79C0
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":829A
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":85B4
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":8E8E
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":9768
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":A042
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":A91C
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":B1F6
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":BAD0
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":C3AA
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":CC84
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":D55E
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":DE38
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":E712
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":EFEC
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":F8C6
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":101A0
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":10A7A
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":11354
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":11C0A
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":124E4
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":12936
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":12D88
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":1553A
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":167BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerLookup.frx":16AD6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCustomerLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cbpSearchOptions_Change()
    If ListView1.ListItems.Count < 1 Then Exit Sub
    Call search_in_listview(ListView1, cbpSearchOptions.Text)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub Form_Activate() 'using form Activate can initialize or activate
    LoadList                'the loadlist command everytime the form is show
End Sub
Private Sub cmdNew_Click()
    frmCustomers.ClearEntries
    frmCustomers.cmdButton(0).Enabled = False
    frmCustomers.cmdButton(3).Enabled = False
    frmCustomers.cmdButton(4).Enabled = False
    frmCustomers.cmdButton(6).Enabled = True
    frmCustomers.Picture1.Enabled = True
    frmCustomers.lNewRec = True
    
    frmCustomers.Show vbModal
End Sub

Private Sub cmdRefresh_Click()
    txbFilterBy.Text = ""
    LoadList
End Sub

Private Sub cmdSelect_Click()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT dbCustomers.* " & _
            "From dbCustomers " & _
            "WHERE (((dbCustomers.CustomerNo)='" & ListView1.SelectedItem.Tag & "'))"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.Open cConnect
    rs.Open sSQL, cn ', adOpenDynamic, adLockOptimistic
    
    frmCustomers.txtCode = rs("CustomerNo")
    frmCustomers.txtName = IIf(IsNull(rs("CustomerName")), "", rs("CustomerName"))
    frmCustomers.txtAddress = IIf(IsNull(rs("Address")), "", rs("Address"))
    frmCustomers.txtContactName = IIf(IsNull(rs("ContactName")), "", rs("ContactName"))
    frmCustomers.txtContactTitle = IIf(IsNull(rs("ContactTitle")), "", rs("ContactTitle"))
    frmCustomers.txtPhoneNo = IIf(IsNull(rs("PhoneNo")), "", rs("PhoneNo"))
    frmCustomers.txtFaxNo = IIf(IsNull(rs("FaxNo")), "", rs("FaxNo"))
    frmCustomers.cbActive = IIf(IsNull(rs("Active")), "", rs("Active"))
     
    frmCustomers.cmdButton(3).Enabled = True
    frmCustomers.cmdButton(4).Enabled = True
    
    rs.Close
    cn.Close
       
    frmCustomers.Show vbModal
End Sub
Public Sub LoadList()
    Dim li As ListItem
    Dim LV As ListView
    Dim vData As String
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Dim sSQL  As String
    
    sSQL = "SELECT dbCustomers.* " & _
            "From dbCustomers " & _
            "WHERE (((dbCustomers.CustomerName) Like '" & txbFilterBy.Text & "%'))"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
        
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    
    Set LV = ListView1
    LV.ListItems.Clear
          
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("CustomerName") & "")
            li.SmallIcon = 2
            li.ListSubItems.Add , , rs("CustomerNo") & ""
            li.ListSubItems.Add , , rs("Active") & ""
            li.Tag = rs("CustomerNo")
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
