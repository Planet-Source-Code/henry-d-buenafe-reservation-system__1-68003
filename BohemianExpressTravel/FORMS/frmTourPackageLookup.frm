VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTourPackageLookup 
   BackColor       =   &H00EDE6E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tour Package Lookup"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "frmTourPackageLookup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8820
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
      TabIndex        =   7
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
         TabIndex        =   8
         Top             =   255
         Width           =   3060
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4305
      TabIndex        =   6
      Top             =   90
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
      Left            =   45
      TabIndex        =   4
      Top             =   15
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
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EDE6E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8820
      TabIndex        =   0
      Top             =   5085
      Width           =   8820
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
         Height          =   525
         Left            =   5100
         TabIndex        =   3
         Top             =   60
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
         Height          =   525
         Left            =   7590
         TabIndex        =   2
         Top             =   60
         Width           =   1215
      End
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
         Height          =   525
         Left            =   6345
         TabIndex        =   1
         Top             =   60
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4365
      Left            =   0
      TabIndex        =   9
      Top             =   735
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7699
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Particular Package"
         Object.Width           =   5592
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Destination"
         Object.Width           =   5645
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2469
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   5175
      Top             =   45
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
            Picture         =   "frmTourPackageLookup.frx":1CFA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":29D4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":3826
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":4678
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":4F52
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":582C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":6106
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":6AD0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":73AA
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":76C4
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":7F9E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":8878
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":9152
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":946C
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":9D46
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":A620
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":AEFA
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":B7D4
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":C0AE
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":C988
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":D262
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":DB3C
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":E416
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":ECF0
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":F5CA
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":FEA4
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":1077E
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":11058
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":11932
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":1220C
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":12AC2
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":1339C
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":137EE
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":13C40
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":163F2
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":17674
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTourPackageLookup.frx":1798E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTourPackageLookup"
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
Private Sub cbpSearchOptions_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdNew_Click()
    frmTourPackage.ClearEntries
    frmTourPackage.cmdButton(0).Enabled = False
    frmTourPackage.cmdButton(3).Enabled = False
    frmTourPackage.cmdButton(4).Enabled = False
    frmTourPackage.cmdButton(6).Enabled = True
    frmTourPackage.Picture1.Enabled = True
    frmTourPackage.lNewRec = True
    frmTourPackage.txtAmount.Text = "0.00"
    frmTourPackage.Show vbModal
End Sub

Private Sub cmdRefresh_Click()
txbFilterBy.Text = ""
LoadList
End Sub

Private Sub cmdSelect_Click()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT dbTourPackage.* " & _
            "From dbTourPackage " & _
            "WHERE (((dbTourPackage.TourPackageID)='" & ListView1.SelectedItem.Tag & "'))"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
      
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenStatic, adLockOptimistic
    
    frmTourPackage.txtCode = rs("TourPackageID")
    frmTourPackage.txtName = IIf(IsNull(rs("Description")), "", rs("Description"))
    frmTourPackage.txtDestination = IIf(IsNull(rs("Destination")), "", rs("Destination"))
    frmTourPackage.txtDetails = IIf(IsNull(rs("Particular")), "", rs("Particular"))
    frmTourPackage.txtAmount = IIf(IsNull(rs("Amount")), "", rs("Amount"))
    
    frmTourPackage.cmdButton(3).Enabled = True
    frmTourPackage.cmdButton(4).Enabled = True
    
    rs.Close
    cn.Close
   
    frmTourPackage.Show vbModal
      
End Sub
Private Sub Form_Activate()
    LoadList
End Sub
Public Sub LoadList()
    Dim li As ListItem
    Dim LV As ListView
    Dim vData As String
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Dim sSQL  As String
    
    sSQL = "SELECT dbTourPackage.* " & _
            "From dbTourPackage " & _
            "WHERE (((dbTourPackage.Description) Like '" & txbFilterBy.Text & "%'))"

    
'    sSQL = "SELECT dbCustomers.* " & _
'            "From dbCustomers " & _
'            "WHERE (((dbCustomers.CustomerName) Like '" & txbFilterBy.Text & "%'))"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
 '   cn.ConnectionString = cConnect
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenStatic, adLockOptimistic
            
    Set LV = ListView1
    'vData = frm.txbFilterBy.Text
    
    LV.ListItems.Clear
   ' LV.ColumnHeaders.Clear
   ' LV.ColumnHeaders.Add , , "PAckage Name", 5000, lvwColumnLeft
   ' LV.ColumnHeaders.Add , , "Code", 1600, lvwColumnLeft
   ' LV.ColumnHeaders.Add , , "Destination", 1000, lvwColumnLeft
   ' LV.ColumnHeaders.Add , , "Amount", 1000, lvwColumnLeft
            
    If (rs.RecordCount > 0) Then
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("Description") & "")
            li.SmallIcon = 15
            li.ListSubItems.Add , , rs("TourPackageID") & ""
            li.ListSubItems.Add , , rs("Destination") & ""
            li.ListSubItems.Add , , Format(rs("Amount"), "#,##0.00") & ""
            li.Tag = rs("TourPackageID")
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    cn.Close
    
    'cmdRefresh.Enabled = (frmTourPackageLookup.ListView1.ListItems.Count > 0)
    'cmdSelect.Enabled = (frmTourPackageLookup.ListView1.ListItems.Count > 0)
    'ListView1.Enabled = (frmCustomerLookup.ListView1.ListItems.Count > 0)
End Sub


Private Sub Form_Load()
LoadList
End Sub
Private Sub ListView1_DblClick()
    cmdSelect_Click
End Sub

Private Sub txbFilterBy_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        LoadList
'        cmdRefresh.Enabled = (ListView1.ListItems.Count > 0)
'        cmdSelect.Enabled = (ListView1.ListItems.Count > 0)
'        ListView1.Enabled = (ListView1.ListItems.Count > 0)
    End If
End Sub


