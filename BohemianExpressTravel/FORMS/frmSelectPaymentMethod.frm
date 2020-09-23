VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectPaymentMethod 
   BackColor       =   &H00EDE6E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Payment Method"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmSelectPaymentMethod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   3
      Top             =   30
      Width           =   5115
      Begin VB.ComboBox txbFilterBy 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   4875
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
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   4800
      Width           =   5145
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
         Left            =   2670
         TabIndex        =   6
         Top             =   15
         Width           =   1185
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
         Left            =   1425
         TabIndex        =   2
         Top             =   15
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
         Height          =   540
         Left            =   3885
         TabIndex        =   1
         Top             =   15
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3960
      Left            =   0
      TabIndex        =   5
      Top             =   765
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   6985
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
End
Attribute VB_Name = "frmSelectPaymentMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OnSelect()

Public pPaymentMethodID As String
Public pDescription As String

Public Sub LoadList()
    Dim li As ListItem
    Dim LV As ListView
    Dim vData As String
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Dim sSQL  As String
    
    sSQL = "SELECT PaymentMethod.* " & _
            "FROM PaymentMethod"
    
'    sSQL = "SELECT Employees.* " & _
'            "From Employees " & _
'            "WHERE (((Employees.Name) Like '" & txbFilterBy.Text & "%'))"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
            
    Set LV = ListView1
    
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "Description", 5000, lvwColumnLeft
                    
    If (rs.RecordCount <> 0) Or (rs.RecordCount = -1) Then
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("Description") & "")
            li.Tag = rs("PaymentMethodID")
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    cn.Close
    
  '  cmdRefresh.Enabled = (frmSelectTourPackage.ListView1.ListItems.Count > 0)
  '  cmdSelect.Enabled = (frmSelectTourPackage.ListView1.ListItems.Count > 0)
End Sub




Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
txbFilterBy.Text = ""
LoadList
End Sub

Private Sub cmdSelect_Click()
 Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT PaymentMethod.* " & _
            "FROM PaymentMethod " & _
            "WHERE (((PaymentMethod.PaymentMethodID)=" & ListView1.SelectedItem.Tag & "))"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn  ', adOpenDynamic, adLockOptimistic
    
    pPaymentMethodID = rs("PaymentMethodID")
    pDescription = IIf(IsNull(rs("Description")), "", rs("Description"))

                
    rs.Close
    cn.Close
    
    RaiseEvent OnSelect
    
    Unload Me
End Sub

Private Sub ListView1_DblClick()
cmdSelect_Click
End Sub


Private Sub txbFilterBy_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
   
    If KeyCode = vbKeyReturn Then
        Dim li As ListItem
        Dim LV As ListView
        Dim vData As String
        Dim rs As ADODB.Recordset
        Dim cn As ADODB.Connection
    
    Dim sSQL  As String
    
    sSQL = "SELECT PaymentMethod.* " & _
            "From PaymentMethod " & _
            "WHERE (((PaymentMethod.Description) Like '" & txbFilterBy.Text & "%'))"

        Set cn = New ADODB.Connection
        Set rs = New ADODB.Recordset
    
        cn.ConnectionString = cConnect
        cn.Open
    
        rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
            Set LV = ListView1
    
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "Description", 2000, lvwColumnLeft
                    
    If (rs.RecordCount <> 0) Or (rs.RecordCount = -1) Then
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("Description") & "")
            li.Tag = rs("PaymentMethodID")
            rs.MoveNext
        Loop
    End If
    
        rs.Close
        cn.Close
    
        End If
End Sub
