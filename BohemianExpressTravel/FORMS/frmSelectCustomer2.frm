VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectCustomer2 
   BackColor       =   &H00EDE6E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "frmSelectCustomer2.frx":0000
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
      TabIndex        =   6
      Top             =   0
      Width           =   3225
      Begin VB.ComboBox cbpSearchOptions 
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
         TabIndex        =   7
         Top             =   255
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
      Height          =   630
      Left            =   4305
      TabIndex        =   5
      Top             =   60
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   4
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
         Left            =   6315
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
         Height          =   540
         Left            =   7575
         TabIndex        =   1
         Top             =   30
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   0
      TabIndex        =   8
      Top             =   735
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
End
Attribute VB_Name = "frmSelectCustomer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OnSelect()

Public pCustomerNo As String
Public pCustomerName As String
Public pAddress As String
Public pPhone As String

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
    
    
'    sSQL = "SELECT Employees.* " & _
'            "From Employees " & _
'            "WHERE (((Employees.LastName) Like '" & txbFilterBy.Text & "%'))"

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
            
    Set LV = ListView1
    
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "Customer Name", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Customer No", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Address", 1500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Phone", 1500, lvwColumnLeft
    
    If (rs.RecordCount <> 0) Or (rs.RecordCount = -1) Then
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("CustomerName") & "")
            li.ListSubItems.Add , , rs("CustomerNo") & ""
            li.ListSubItems.Add , , rs("Address") & ""
            li.ListSubItems.Add , , rs("PhoneNo") & ""
            li.Tag = rs("CustomerNo")
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    cn.Close
    
    cmdRefresh.Enabled = (frmSelectTourPackage.ListView1.ListItems.Count > 0)
    cmdSelect.Enabled = (frmSelectTourPackage.ListView1.ListItems.Count > 0)
End Sub


Private Sub cbpSearchOptions_Change()
If ListView1.ListItems.Count < 1 Then Exit Sub
Call search_in_listview(ListView1, cbpSearchOptions.Text)
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
    
    
    sSQL = "SELECT dbCustomers.* " & _
            "From dbCustomers " & _
            "WHERE (((dbCustomers.CustomerNo) =  '" & ListView1.SelectedItem.Tag & "'))"


'    sSQL = "SELECT Employees.* " & _
'            "From Employees " & _
'            "WHERE (((Employees.EmployeeID)='" & ListView1.SelectedItem.Tag & "'))"
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn  ', adOpenDynamic, adLockOptimistic
    
    pCustomerNo = rs("CustomerNo")
    pCustomerName = IIf(IsNull(rs("CustomerName")), "", rs("CustomerName"))
    pAddress = IIf(IsNull(rs("Address")), "", rs("Address"))
    pPhone = IIf(IsNull(rs("PhoneNo")), "", rs("PhoneNo"))
    
                
    rs.Close
    cn.Close
    
    RaiseEvent OnSelect
    
    Unload Me
End Sub

Private Sub ListView1_DblClick()
cmdSelect_Click

End Sub

Private Sub txbFilterBy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        LoadList
    End If
End Sub
