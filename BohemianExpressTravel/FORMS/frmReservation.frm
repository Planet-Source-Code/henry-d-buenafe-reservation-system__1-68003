VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReservation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservation Accounts"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReservation.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker Dt1 
      Height          =   300
      Left            =   8145
      TabIndex        =   17
      Top             =   840
      Width           =   1365
      _ExtentX        =   2408
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
      Format          =   54460417
      CurrentDate     =   39107
   End
   Begin VB.TextBox txtClient 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5070
      TabIndex        =   15
      Top             =   840
      Width           =   1470
   End
   Begin VB.TextBox txtTourNo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6600
      TabIndex        =   13
      Top             =   840
      Width           =   1470
   End
   Begin VB.CommandButton btnPrint 
      Height          =   585
      Left            =   2265
      Picture         =   "frmReservation.frx":39F50
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   585
      Width           =   660
   End
   Begin VB.CommandButton btnRefresh 
      Height          =   585
      Left            =   795
      Picture         =   "frmReservation.frx":3A13F
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   585
      Width           =   660
   End
   Begin VB.CommandButton btnsearch 
      Height          =   585
      Left            =   1530
      Picture         =   "frmReservation.frx":3A264
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   585
      Width           =   660
   End
   Begin VB.CommandButton btnClose 
      Height          =   585
      Left            =   3000
      Picture         =   "frmReservation.frx":3A5A2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   585
      Width           =   660
   End
   Begin VB.CommandButton btnNew 
      Height          =   585
      Left            =   75
      Picture         =   "frmReservation.frx":3A9E4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   585
      Width           =   660
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5760
      Left            =   30
      TabIndex        =   0
      Top             =   1545
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   10160
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
         Text            =   "Tour Date"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Client Name"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tour No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Destination"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Incharge"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Donwpayment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Client No"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search Client No"
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
      Left            =   5160
      TabIndex        =   16
      Top             =   1230
      Width           =   1275
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search Tour No."
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
      Left            =   6765
      TabIndex        =   14
      Top             =   1230
      Width           =   1110
   End
   Begin VB.Label lblSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search Date"
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
      Left            =   8370
      TabIndex        =   12
      Top             =   1215
      Width           =   945
   End
   Begin VB.Label Label5 
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
      Left            =   2970
      TabIndex        =   11
      Top             =   1215
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
      Left            =   765
      TabIndex        =   9
      Top             =   1215
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
      ForeColor       =   &H00C0FFFF&
      Height          =   165
      Left            =   1500
      TabIndex        =   8
      Top             =   1215
      Width           =   705
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
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
      Left            =   2235
      TabIndex        =   7
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
      Left            =   60
      TabIndex        =   3
      Top             =   1215
      Width           =   705
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reservation System Version 1.0"
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
      Left            =   5685
      TabIndex        =   1
      Top             =   5385
      Width           =   9795
   End
End
Attribute VB_Name = "frmReservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnNew_Click()
frmTourIN.Show vbModal
End Sub

Private Sub btnRefresh_Click()
Reservation
End Sub
Private Sub Form_Load()
Reservation
End Sub

Private Sub Reservation()
Dim li As ListItem
Dim LV As ListView

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    
    .Open "SELECT * FROM dbReserve ORDER BY DATE", cConnect, _
     adOpenDynamic, adLockOptimistic
        
        Set LV = ListView1
        LV.ListItems.Clear
     
        If .RecordCount <> 0 Then
        .MoveFirst
        Do While Not .EOF
    
        Set li = LV.ListItems.Add(, , rs!Date & "")
        li.ListSubItems.Add , , rs!ClientName & ""
        li.ListSubItems.Add , , rs!TourNo & ""
        li.ListSubItems.Add , , rs!Destination & ""
        li.ListSubItems.Add , , rs!Incharge & ""
        li.ListSubItems.Add , , rs!Downpayment & ""
        li.ListSubItems.Add , , rs!Balance & ""
        li.ListSubItems.Add , , rs!ClientID & ""
        .MoveNext
                
    Loop
    End If
End With
End Sub

Private Sub Label6_Click()

End Sub

Private Sub lblSearch_Click()
On Error GoTo DateERR
Dim li As ListItem
Dim LV As ListView

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    
    .Open "SELECT * FROM dbReserve WHERE Date=#" & Dt1.Value & "#", cConnect, _
     adOpenDynamic, adLockOptimistic
        
        Set LV = ListView1
        LV.ListItems.Clear
     
        If .RecordCount <> 0 Then
        .MoveFirst
        Do While Not .EOF
    
        Set li = LV.ListItems.Add(, , rs!Date & "")
        li.ListSubItems.Add , , rs!ClientName & ""
        li.ListSubItems.Add , , rs!TourNo & ""
        li.ListSubItems.Add , , rs!Destination & ""
        li.ListSubItems.Add , , rs!Incharge & ""
        li.ListSubItems.Add , , rs!Downpayment & ""
        li.ListSubItems.Add , , rs!Balance & ""
        li.ListSubItems.Add , , rs!ClientID & ""
        .MoveNext
        
        Loop
        Exit Sub
    End If
End With
DateERR:
    MsgBox "No Found Record!"
End Sub

Private Sub txtClient_Change()
If ListView1.ListItems.Count < 1 Then Exit Sub
    Call search_in_listview(ListView1, txtClient.Text)
End Sub

Private Sub txtTourNo_Change()
Dim li As ListItem
Dim LV As ListView

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

With rs
    
    .Open "SELECT * FROM dbReserve WHERE TourNo='" & txtTourNo.Text & "'", cConnect
    '_  adOpenDynamic, adLockOptimistic
        
        Set LV = ListView1
        LV.ListItems.Clear
     
        If .RecordCount <> 0 Then
        .MoveFirst
        Do While Not .EOF
    
        Set li = LV.ListItems.Add(, , rs!Date & "")
        li.ListSubItems.Add , , rs!ClientName & ""
        li.ListSubItems.Add , , rs!TourNo & ""
        li.ListSubItems.Add , , rs!Destination & ""
        li.ListSubItems.Add , , rs!Incharge & ""
        li.ListSubItems.Add , , rs!Downpayment & ""
        li.ListSubItems.Add , , rs!Balance & ""
        li.ListSubItems.Add , , rs!ClientID & ""
        .MoveNext
                
    Loop
    End If
End With
End Sub
