VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSchedule 
   BackColor       =   &H00EDE6E0&
   Caption         =   "Tour Schedule"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   9960
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search Tour Schedule"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Tour Account"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComCtl2.DTPicker dt1 
         Height          =   300
         Left            =   2445
         TabIndex        =   3
         Top             =   180
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20250625
         CurrentDate     =   39118
      End
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   60
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":1992
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":3324
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":6648
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":7FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":996C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":B2FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":CC90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":E624
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":F300
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":FBE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":108BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":11598
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":12274
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":12F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":13C2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   45
      TabIndex        =   2
      Top             =   810
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   15
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblDate.Caption = Date
    dt1.Value = Date
End Sub
Private Sub Form_Resize()
    'This will resize the grid when the form is resized
    ListView1.Height = IIf(Me.ScaleHeight = 0, 0, Me.ScaleHeight - 1300)
    ListView1.Width = Me.ScaleWidth
End Sub
Private Sub SearchDATE()
'On Error GoTo SearchErr
Dim sSQL  As String
Dim rpt As New rptDailySales
'Dim d As String

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    'd = InputBox("Enter Date To to search tour date:")
  
     '   sSQL = "SELECT Tours.* " & _
     '       "From Tours " & _
     '       "WHERE (((Tours.TourDate) Between #" & lblDate & "# And #" & d & "#))"
     
     sSQL = "SELECT Tours.* " & _
            "From Tours " & _
            "WHERE (((Tours.TourDate) Between #" & lblDate & "# And #" & dt1.Value & "#))"
    
    cn.ConnectionString = cConnect
    cn.Open
    rs.Open sSQL, cn
        
    Set LV = ListView1
    'vData = frm.txbFilterBy.Text
    
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "Tour Date", 1200, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Tour Destination", 5000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Tour No", 1600, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Tour Package", 2200, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Tour Guide", 1600, lvwColumnLeft
    LV.ColumnHeaders.Add , , "No of Person", 1200, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Amount", 2000, lvwColumnRight
    
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
       
        rec = rs.RecordCount
        rs.MoveFirst
        
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("TourDate") & "")
            li.ListSubItems.Add , , rs("Destination") & ""
            li.ListSubItems.Add , , rs("TourNo") & ""
            li.ListSubItems.Add , , rs("TourPackageID") & ""
            li.ListSubItems.Add , , rs("TourGuideID") & ""
            li.ListSubItems.Add , , rs("NoPerson") & ""
            li.ListSubItems.Add , , Format(rs("Amount"), "#,##0.00") & ""
            li.Tag = rs("TourNo")
            rs.MoveNext
        Loop
        
    End If
    
    ', adOpenStatic, adLockBatchOptimistic
       
    'Set rpt.DataSource = rs
    'rpt.Title = "Bohemian Express Travel Sales Report"
    'rpt.Show vbModal
           
           
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    
'    Exit Sub
'SearchErr:
'    MsgBox "No Transaction found!", vbCritical, "Warning"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
    Case "Search"
        SearchDATE
    Case "Exit"
        Unload Me
    End Select
End Sub
