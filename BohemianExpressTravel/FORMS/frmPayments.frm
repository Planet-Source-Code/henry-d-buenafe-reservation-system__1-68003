VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPayments 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payments"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "frmPayments.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00EDE6E0&
      Height          =   3270
      Left            =   0
      ScaleHeight     =   3210
      ScaleWidth      =   9810
      TabIndex        =   0
      Top             =   0
      Width           =   9870
      Begin VB.TextBox txtPayeeName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4155
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   24
         Top             =   2715
         Width           =   2700
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3780
         Picture         =   "frmPayments.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2715
         Width           =   360
      End
      Begin VB.TextBox txtPayeeID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   22
         Top             =   2715
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   8490
         TabIndex        =   20
         Top             =   1980
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   7185
         TabIndex        =   19
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtTourNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         MaxLength       =   14
         TabIndex        =   17
         Top             =   420
         Width           =   1935
      End
      Begin VB.TextBox txtReceipt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7785
         MaxLength       =   14
         TabIndex        =   8
         Top             =   45
         Width           =   1935
      End
      Begin VB.TextBox txtApprovalNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         MaxLength       =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   3405
      End
      Begin VB.TextBox txtPaymentID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   6
         Top             =   765
         Width           =   1560
      End
      Begin VB.TextBox txtCardNumber 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         MaxLength       =   80
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1455
         Width           =   3405
      End
      Begin VB.CommandButton cmdLookUpTourPackage 
         Height          =   315
         Left            =   3780
         Picture         =   "frmPayments.frx":17F4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   765
         Width           =   360
      End
      Begin VB.TextBox txtCheckNumber 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         MaxLength       =   120
         TabIndex        =   3
         Top             =   2145
         Width           =   3405
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         MaxLength       =   16
         TabIndex        =   2
         Top             =   1110
         Width           =   1950
      End
      Begin VB.TextBox txtPayDescription 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4155
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   1
         Top             =   765
         Width           =   2700
      End
      Begin MSComCtl2.DTPicker dtTransDate 
         Height          =   315
         Left            =   2205
         TabIndex        =   9
         Top             =   60
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54001665
         CurrentDate     =   38353
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paid By"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   21
         Top             =   2820
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tour No."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   18
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   2235
         Width           =   1215
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approval No."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   1860
         Width           =   1065
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card number"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   14
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Method"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt No."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   6675
         TabIndex        =   12
         Top             =   135
         Width           =   945
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans Date"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   105
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   1155
         Width           =   675
      End
   End
   Begin VB.Label lblLogo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reservation System Version 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   30
      TabIndex        =   25
      Tag             =   "App Description"
      Top             =   3360
      Width           =   3150
   End
End
Attribute VB_Name = "frmPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OnPay()
Private WithEvents frmSelPayMethod As frmSelectPaymentMethod
Attribute frmSelPayMethod.VB_VarHelpID = -1
Dim dbConn As ADODB.Connection
Dim dbRec As ADODB.Recordset
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdLookUpTourPackage_Click()
    frmSelPayMethod.LoadList
    frmSelPayMethod.Show vbModal
End Sub
Private Sub cmdSave_Click()
    Dim sSQL  As String
    Dim rpt As New rtpReceipt
            
    sSQL = "SELECT Payments.* " & _
            "From Payments " & _
            "WHERE (((Payments.PaymentID)='" & txtReceipt.Text & "'))"
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
   
    If txtPaymentID.Text = "" Then
'        cn.Close
        MsgBox "Please enter Payment Method.", vbOKOnly, "Empty"
        Exit Sub
    End If
    
    If txtAmount.Text = "" Then
 '       cn.Close
        MsgBox "Please enter Payment Method.", vbOKOnly, "Empty"
        txtAmount.SetFocus
        Exit Sub
    End If
        
    SaveRecord
    IncrementRefNo
         
    RaiseEvent OnPay
                              
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenStatic, adLockBatchOptimistic
           
    If MsgBox("Do you want to print this payments?", vbYesNo) = vbYes Then
           
    Set rpt.DataSource = rs
        
    rpt.Title = "Bohemian Express Travel"
    rpt.Show vbModal
           
   ' End If
    
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Initialize()
    Set frmSelPayMethod = New frmSelectPaymentMethod
End Sub

Private Sub Form_Load()
    
    txtReceipt.MaxLength = 10
    txtCheckNumber.MaxLength = 10
    txtApprovalNo.MaxLength = 10
    txtCardNumber.MaxLength = 15
    dtTransDate.Value = Date
    Picture1.Enabled = True
    
End Sub

Private Sub Form_Terminate()
    Set frmSelPayMethod = Nothing
End Sub

Private Sub frmSelPayMethod_OnSelect()
    txtPaymentID.Text = frmSelPayMethod.pPaymentMethodID
    txtPayDescription.Text = frmSelPayMethod.pDescription
End Sub

Private Sub SaveRecord()
On Error GoTo SaveErr
    Dim sSQL As String
    
    sSQL = "SELECT Payments.* " & _
            "FROM Payments"
        
    Set dbConn = New ADODB.Connection
    Set dbRec = New ADODB.Recordset
    
    dbConn.ConnectionString = cConnect
    dbConn.Open
    
    dbRec.Open sSQL, dbConn, adOpenDynamic, adLockOptimistic
        
    With dbRec
        .AddNew
        .Fields("PaymentID") = txtReceipt.Text
        .Fields("PaymentDate") = dtTransDate.Value
        .Fields("TourNo") = txtTourNo.Text
        .Fields("PaymentMethod") = txtPaymentID.Text
        .Fields("Amount") = txtAmount.Text
        .Fields("CreditCardNumber") = txtCardNumber.Text
        .Fields("Approval") = txtApprovalNo.Text
        .Fields("CheckNumber") = txtCheckNumber.Text
        .Fields("PaidBy") = txtPayeeID.Text
        .Update
    End With
    
    dbRec.Close
    dbConn.Close
        
    Set dbConn = Nothing
    Set dbRec = Nothing
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Public Function GetLastRefNo() As Long
    Dim sSQL As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
                
    sSQL = "SELECT sysLastNo.LastNo From sysLastNo " & _
           "WHERE (((sysLastNo.Key)='RECEIPTNO')) "
        
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = cConnect
    cn.Open
    
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        GetLastRefNo = rs("LastNo").Value + 1
    Else
        GetLastRefNo = 1
    End If
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
End Function

Private Sub IncrementRefNo()
On Error GoTo HandleError
    Dim sSQL As String
    
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
'    If lNewRec = False Then
'        Exit Sub
'    End If
    
    sSQL = "SELECT sysLastNo.* From sysLastNo " & _
           "WHERE sysLastNo.Key = 'RECEIPTNO' "
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = cConnect
    cn.Open
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    
     If Not (rs.BOF And rs.EOF) Then
            rs("LastNo") = rs("LastNo") + 1
        Else
            rs.AddNew
            rs("Key") = "RECEIPTNO"
            rs("LastNo") = 1
        End If
        
        rs.Update
        
        rs.Close
        cn.Close
        Set rs = Nothing
        Set cn = Nothing

    Exit Sub
HandleError:
    MsgBox Error
End Sub
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub txtApprovalNo_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub txtCardNumber_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub txtCheckNumber_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
