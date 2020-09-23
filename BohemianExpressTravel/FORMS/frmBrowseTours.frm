VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowseTours 
   BackColor       =   &H00EDE6E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lookup"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "frmBrowseTours.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDE6E0&
      Caption         =   "Search: "
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
      Left            =   5595
      TabIndex        =   7
      Top             =   15
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
      Height          =   585
      Left            =   4305
      TabIndex        =   6
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
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
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
         Left            =   75
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
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   8865
      TabIndex        =   0
      Top             =   4995
      Width           =   8865
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
         Height          =   495
         Left            =   6345
         TabIndex        =   3
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7605
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5085
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4260
      Left            =   15
      TabIndex        =   9
      Top             =   705
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7514
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
      NumItems        =   0
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
            Picture         =   "frmBrowseTours.frx":57E2
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":64BC
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":730E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":8160
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":8A3A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":9314
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":9BEE
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":A5B8
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":AE92
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":B1AC
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":BA86
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":C360
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":CC3A
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":CF54
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":D82E
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":E108
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":E9E2
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":F2BC
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":FB96
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":10470
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":10D4A
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":11624
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":11EFE
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":127D8
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":130B2
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":1398C
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":14266
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":14B40
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":1541A
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":15CF4
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":165AA
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":16E84
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":172D6
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":17728
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":19EDA
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":1B15C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseTours.frx":1B476
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowseTours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OnSelect()

Public pTourNo As String
Public pTransdate As Date
Public pTourDate As Date
Public pTourPackageID As String
Public pTourPackageName As String
Public pAmount As Double
'Public pDesctination As String
Public pNoPerson As Integer
Public pDestination As String
Public pTourGuideID As String
Public pDescription As String
Public pTourPackageAmount As Double
Public pName As String


Public Sub LoadList()
    Dim li As ListItem
    Dim LV As ListView
    Dim vData As String
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Dim sSQL  As String
    
    sSQL = "SELECT Tours.*, dbTourPackage.Description, dbTourPackage.Amount AS TourPackageAmount, Employees.Name " & _
            "FROM (Tours LEFT JOIN dbTourPackage ON Tours.TourPackageID = dbTourPackage.TourPackageID) LEFT JOIN Employees ON Tours.TourGuideID = Employees.EmployeeID " & _
            "ORDER BY Tours.TourNo DESC"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
   ' cn.ConnectionString = cConnect
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
       
    Set LV = ListView1
      
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "TourDate", 1500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "TourNo", 1500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Destination", 2500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Amount", 2000, lvwColumnRight
                
    If (rs.RecordCount <> 0) Or (rs.RecordCount = -1) Then
'        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("TourDate") & "")
            li.SmallIcon = 24
            li.ListSubItems.Add , , rs("TourNo") & ""
            li.ListSubItems.Add , , rs("Destination") & ""
            li.ListSubItems.Add , , Format(rs("Amount"), "#,##0.00") & ""
            li.Tag = rs("TourNo")
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    cn.Close
   
End Sub


Private Sub cbpSearchOptions_Change()
    If ListView1.ListItems.Count < 1 Then Exit Sub
    Call search_in_listview(ListView1, cbpSearchOptions.Text)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
LoadList
End Sub

Private Sub cmdSelect_Click()
 Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT Tours.*, dbTourPackage.Description, dbTourPackage.Amount AS TourPackageAmount, Employees.Name " & _
            "FROM (Tours LEFT JOIN dbTourPackage ON Tours.TourPackageID = dbTourPackage.TourPackageID) LEFT JOIN Employees ON Tours.TourGuideID = Employees.EmployeeID " & _
            "Where (((Tours.TourNo) = '" & ListView1.SelectedItem.Tag & "')) " & _
            "ORDER BY Tours.TourNo DESC"
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
  '  cn.ConnectionString = cConnect
    cn.Open cConnect
    
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
        
    pTourNo = rs("TourNo")
    pTransdate = IIf(IsNull(rs("TransDate")), "", rs("TransDate"))
    pTourDate = IIf(IsNull(rs("TourDate")), "", rs("TourDate"))
    pTourPackageID = IIf(IsNull(rs("TourPackageID")), "", rs("TourPackageID"))
    pTourPackageName = IIf(IsNull(rs("Description")), "", rs("Description"))
    pAmount = IIf(IsNull(rs("Amount")), "", rs("Amount"))
    pDestination = IIf(IsNull(rs("Destination")), "", rs("Destination"))
    pNoPerson = IIf(IsNull(rs("NoPerson")), "", rs("NoPerson"))
    pTourGuideID = IIf(IsNull(rs("TourGuideID")), "", rs("TourGuideID"))
    pDescription = IIf(IsNull(rs("Description")), "", rs("Description"))
    pTourPackageAmount = IIf(IsNull(rs("TourPackageAmount")), "", rs("TourPackageAmount"))
    pName = IIf(IsNull(rs("Name")), "", rs("Name"))
                
    rs.Close
    cn.Close
    
    'Set rs = Nothing
    'Set cn = Nothing
    
    RaiseEvent OnSelect
    
    Unload Me

End Sub

Private Sub ListView1_DblClick()
    cmdSelect_Click
End Sub


