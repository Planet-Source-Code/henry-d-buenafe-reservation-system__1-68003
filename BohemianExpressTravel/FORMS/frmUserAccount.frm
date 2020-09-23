VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserAccount 
   BackColor       =   &H00C4BFB7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Account"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1770
      Picture         =   "frmUserAccount.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   90
      Width           =   840
   End
   Begin VB.CommandButton btnSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   945
      Picture         =   "frmUserAccount.frx":3EB4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   840
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      Picture         =   "frmUserAccount.frx":4079
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C4BFB7&
      Caption         =   "Search Opotions:"
      Height          =   675
      Left            =   4140
      TabIndex        =   1
      Top             =   180
      Width           =   3225
      Begin VB.ComboBox cbpSearchOptions 
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Top             =   255
         Width           =   3060
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5595
      Left            =   105
      TabIndex        =   0
      Top             =   930
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   9869
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
         Text            =   "Account Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   585
      Top             =   -135
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
            Picture         =   "frmUserAccount.frx":4466
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":5140
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":5F92
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":6DE4
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":76BE
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":7F98
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":8872
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":923C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":9B16
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":9E30
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":A70A
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":AFE4
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":B8BE
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":BBD8
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":C4B2
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":CD8C
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":D666
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":DF40
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":E81A
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":F0F4
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":F9CE
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":102A8
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":10B82
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":1145C
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":11D36
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":12610
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":12EEA
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":137C4
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":1409E
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":14978
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":1522E
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":15B08
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":15F5A
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":163AC
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":18B5E
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":19DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":1A0FA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnNew_Click()
    frmUser.ClearEntries
    frmUser.UnlockME
    frmUser.txtIDNo.Locked = False
    frmUser.cmdButton(0).Enabled = False
    frmUser.cmdButton(3).Enabled = False
    frmUser.cmdButton(4).Enabled = False
    frmUser.cmdButton(6).Enabled = True
    frmUser.Picture1.Enabled = True
    frmUser.lNewRec = True
    frmUser.Show vbModal
End Sub

Private Sub btnSelect_Click()
    Dim cConn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT dbLogin.* " & _
            "From dbLogin " & _
            "WHERE (((dbLogin.UserID)='" & ListView1.SelectedItem.Tag & "'))"

    Set cConn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cConn.Open cConnect
    rs.Open sSQL, cConn, adOpenDynamic, adLockOptimistic
    
    frmUser.txtIDNo = rs("UserID")
    frmUser.txtName = IIf(IsNull(rs("UserName")), "", rs("UserName"))
    frmUser.txtpassword = IIf(IsNull(rs("Password")), "", rs("Password"))
        
    frmUser.cmdButton(3).Enabled = True
    frmUser.cmdButton(4).Enabled = True
    frmUser.Picture1.Enabled = True
    
    rs.Close
    cConn.Close
    
    frmUser.Show vbModal
End Sub

Private Sub cbpSearchOptions_Change()
  If ListView1.ListItems.Count < 1 Then Exit Sub
    Call search_in_listview(ListView1, cbpSearchOptions.Text)
End Sub

Private Sub Form_Load()
LoadList
End Sub
Public Sub LoadList()
 '   On Error Resume Next
    Dim li As ListItem
    Dim LV As ListView
    Dim vData As String
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
        
    Dim sSQL  As String
    
    sSQL = "SELECT *FROM dbLogin ORDER BY Username"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
        
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    Set LV = ListView1
    LV.ListItems.Clear
            
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
     rec = rs.RecordCount
        rs.MoveFirst
        
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("Username") & "")
            li.SmallIcon = 12
            li.ListSubItems.Add , , rs("UserID") & ""
            'li.ListSubItems.Add , , rs("Password") & ""
            li.ListSubItems.Add , , String(Len(rs.Fields!Password), "X")
            li.Tag = rs("UserID")
            rs.MoveNext
        Loop
     End If
    
    rs.Close
    cn.Close
    
End Sub
Private Sub ListView1_DblClick()
btnSelect_Click
End Sub
