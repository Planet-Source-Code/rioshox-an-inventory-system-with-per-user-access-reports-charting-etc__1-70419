VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPendingSO 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   675
      Width           =   960
   End
   Begin VB.CommandButton cmdPrintInvoice 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print Invoice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton cmdPrintOrderSlip 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print SO Slip"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4320
      Width           =   1590
   End
   Begin VB.CommandButton cmdShip 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Approve"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1890
      Width           =   960
   End
   Begin VB.ComboBox cboFilterStatus 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmPendingSO.frx":0000
      Left            =   1905
      List            =   "frmPendingSO.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   675
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.CommandButton cmdDecline 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Decline"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2415
      Width           =   960
   End
   Begin VB.CommandButton cmdViewCustomer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7365
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4335
      Width           =   1590
   End
   Begin VB.CommandButton cmdApproved 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Endorse"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1530
      Width           =   960
   End
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9345
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8640
      Width           =   960
   End
   Begin MSFlexGridLib.MSFlexGrid msfPendingSOHeader 
      Height          =   2505
      Left            =   375
      TabIndex        =   6
      Top             =   1470
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   4419
      _Version        =   393216
      Rows            =   7
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14663855
      GridColor       =   0
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid msfSODetails 
      Height          =   2400
      Left            =   585
      TabIndex        =   12
      Top             =   5520
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   4233
      _Version        =   393216
      Rows            =   7
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14663855
      GridColor       =   0
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblRecordCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5745
      TabIndex        =   23
      Top             =   4020
      Width           =   3225
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   405
      TabIndex        =   19
      Top             =   735
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Total: 0.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6030
      TabIndex        =   16
      Top             =   8235
      Width           =   3825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   435
      TabIndex        =   15
      Top             =   4890
      Width           =   2250
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   5
      Left            =   585
      TabIndex        =   11
      Top             =   5220
      Width           =   1635
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   6
      Left            =   2220
      TabIndex        =   10
      Top             =   5220
      Width           =   3480
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unit Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   7
      Left            =   5700
      TabIndex        =   9
      Top             =   5220
      Width           =   1200
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   8
      Left            =   6885
      TabIndex        =   8
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Extended Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   9
      Left            =   7965
      TabIndex        =   7
      Top             =   5220
      Width           =   1890
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Order Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   3945
      TabIndex        =   5
      Top             =   1140
      Width           =   1590
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   7335
      TabIndex        =   4
      Top             =   1140
      Width           =   1620
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Required Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   5535
      TabIndex        =   3
      Top             =   1140
      Width           =   1800
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   2025
      TabIndex        =   2
      Top             =   1140
      Width           =   1920
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transaction No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   390
      TabIndex        =   1
      Top             =   1140
      Width           =   1635
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Pending Orders"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   0
      Top             =   165
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   4185
      Left            =   195
      Top             =   540
      Width           =   10140
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3795
      Left            =   180
      Top             =   4785
      Width           =   10140
   End
End
Attribute VB_Name = "frmPendingSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboFilterStatus_Click()
    'Select Case cboFilterStatus.ListIndex
        'Case 0:
        '    cmdApproved.Enabled = True
        '    cmdShip.Enabled = False
        '    cmdDecline = True
        '    lblTitle = "Pending Order(s)"
        '    If g_CurrentUser.ApproveSO Then Me.cmdApproved.Enabled = True
        'Case 1:
        '    cmdApproved.Enabled = False
        '    cmdShip.Enabled = True
        '    cmdDecline = True
        '    lblTitle = "Pending Shipments(s)"
        '    If g_CurrentUser.ApproveShip Then cmdShip.Enabled = True
        'Case 2:
        '    cmdApproved.Enabled = False
        '    cmdShip.Enabled = False
        '    cmdDecline = False
        '    lblTitle = "Declined Order(s)"
    'End Select
    
    
    
    RefreshHeader
End Sub

Private Sub cmdApproved_Click()
    Dim rsApproved As ADODB.Recordset
    If msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0) = "" Then
        MsgBox "nothing to approve!", vbCritical
        Exit Sub
    End If
    
    
    Set rsApproved = New ADODB.Recordset
    rsApproved.Open _
        "SELECT tblTransactions.TransactionID, tblTransactions.EmployeeID, [LastName]+', '+[FirstName]  AS EmployeeName " & _
        "FROM tblTransactions LEFT JOIN tblEmployees ON tblTransactions.EmployeeID = tblEmployees.EmployeeID " & _
        "WHERE tblTransactions.EmployeeID = 0 AND TransactionID = " & msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0), connRFM, adOpenStatic
    
    On Error GoTo cmdApproved_Click_ERR
    If Not NoRecord(rsApproved) Then
        connRFM.Execute "UPDATE tblTransactions SET EmployeeID = " & g_CurrentUser.EmployeeID & " WHERE TransactionID = " & msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0)
        MsgBox "Transaction number " & msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0) & " ready for shipping!", vbInformation
    Else
        MsgBox "Cannot approve, transaction already approved by " & rsApproved("EmployeeName") & "." & vbCrLf & "Refresh list and/or check status of Transaction.", vbCritical
    End If
    
    ClearRS rsApproved
    
    RefreshHeader
    
cmdApproved_Click_ERR:
    If Err.Number <> 0 Then
        MsgBox "Cannot perform current operation because an error occurred." & vbCrLf & Err.Description, vbCritical
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDecline_Click()
    Dim ReasonText$
    
    If msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0) = "" Then
        MsgBox "Nothing to decline!", vbCritical
        Exit Sub
    End If
    
    If MsgBox("Sure to decline?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    ReasonText$ = InputBox("You can input a short note why you choose to decline the transaction", "Reason", "Type something here")
    
    connRFM.Execute "INSERT INTO tblDeclinedTransactions (ReasonText,TransactionID,DeclinedDate,EmployeeID) VALUES  (" & _
        Enquote(ReasonText, True) & _
        EnNone(msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0), True) & _
        EnNone(Now(), True) & _
        EnNone(CStr(g_CurrentUser.EmployeeID)) & ")"
        
    RefreshHeader
    MsgBox "Transaction declined!", vbInformation
End Sub

Private Sub cmdPrintInvoice_Click()
    If msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0) = "" Then
        MsgBox "Nothing to print!", vbCritical
        Exit Sub
    End If
    
    ShowReport "rptInvoice", "TransactionID = " & Enquote(msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0))
End Sub

Private Sub cmdPrintOrderSlip_Click()
    If msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0) = "" Then
        MsgBox "Nothing to print!", vbCritical
        Exit Sub
    End If
    
    ShowReport "rptOrderSlip", "TransactionID = " & Enquote(msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0))
End Sub

Private Sub cmdRefresh_Click()
    RefreshHeader
End Sub

Private Sub cmdShip_Click()
    Dim rsShipped As ADODB.Recordset, bHasBegun As Boolean
    Dim rsTransactionDetails As ADODB.Recordset
    
    If msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0) = "" Then
        MsgBox "Nothing to mark as shipped!", vbCritical
        Exit Sub
    End If
    
    Set rsShipped = New ADODB.Recordset
    
    On Error GoTo cmdShip_Click_ERR
    rsShipped.Open _
        "SELECT tblPosted.TransactionID, tblPosted.EmployeeID, [LastName]+', '+[FirstName]+' '+Left([MiddleName],1)+'.' AS EmployeeName " & _
        "FROM tblEmployees INNER JOIN tblPosted ON tblEmployees.EmployeeID = tblPosted.EmployeeID " & _
        "WHERE TransactionID = " & msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0), connRFM, adOpenStatic
    
    
    If NoRecord(rsShipped) Then
        connRFM.BeginTrans
            bHasBegun = True
            
            connRFM.Execute "UPDATE tblTransactions SET ShippedDate = " & Enquote(Date) & " WHERE TransactionID = " & msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0)
            connRFM.Execute _
                "INSERT INTO tblPosted (TransactionID,EmployeeID) VALUES (" & _
                EnNone(msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0), True) & _
                EnNone(CStr(g_CurrentUser.EmployeeID)) & ")"
                
            Set rsTransactionDetails = New ADODB.Recordset
            rsTransactionDetails.Open "SELECT TransactionID,ProductID,Quantity FROM tblTransactionDetails WHERE TransactionID = " & msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0), connRFM, adOpenStatic
            If Not NoRecord(rsTransactionDetails) Then
                connRFM.Execute "UPDATE tblProducts SET UnitsOnOrder = UnitsOnOrder - " & rsTransactionDetails("Quantity") & " WHERE ProductID = " & rsTransactionDetails("ProductID")
                connRFM.Execute "UPDATE tblProducts SET UnitsInStock = UnitsInStock - " & rsTransactionDetails("Quantity") & " WHERE ProductID = " & rsTransactionDetails("ProductID")
                rsTransactionDetails.MoveNext
            End If
        
        connRFM.CommitTrans
        
         
        '=================================== Uncomment to use cell phone notification
        'Dim dNumber$, msgResult$
        'dNumber$ = GetNumber(msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 1))
        'If msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 4) = "SMS" Then msgResult = Mobile1.SendSMSMessage(dNumber, "Your order was approved and will be delivered on the date you specified", 167, 0, 0, "")
        '=================================== Uncomment to use cell phone notification
        
        MsgBox "Transaction number " & msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0) & " marked as shipped!", vbInformation
    Else
        MsgBox "Cannot mark this transaction as shipped, transaction already marked as shipped by " & rsShipped("EmployeeName") & "." & vbCrLf & "Refresh list and/or check status of Transaction.", vbCritical
    End If
    
    ClearRS rsShipped
    
    RefreshHeader
    
cmdShip_Click_ERR:
    If Err.Number <> 0 Then
        If bHasBegun Then connRFM.RollbackTrans
        MsgBox "Cannot perform current operation because an error occurred." & vbCrLf & Err.Description, vbCritical
    End If
End Sub

Private Sub cmdViewCustomer_Click()
    Dim rsActiveTable As ADODB.Recordset
    
    If msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 1) = "" Then Exit Sub
    
    Set rsActiveTable = New ADODB.Recordset
    rsActiveTable.Open "SELECT * FROM tblCustomers WHERE CustomerID = " & Enquote(msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 1)), connRFM
    
    Hide
    With frmCustomersDE
        .Tag = TO_EDIT
        
        .lblTitle = "Customer Details"
        
        .txtCustomerID.Locked = True
        
        .txtCustomerID = rsActiveTable("CustomerID")
        .txtCompanyName = rsActiveTable("Companyname") & ""
        .txtContactName = rsActiveTable("ContactName") & ""
        .txtAddress = rsActiveTable("Address") & ""
        .txtCity = rsActiveTable("City") & ""
        .txtRegion = rsActiveTable("Region") & ""
        .txtPhone = rsActiveTable("Phone") & ""
        .txtZIP = rsActiveTable("PostalCode") & ""
        .txtEmail = rsActiveTable("EmailAddress") & ""
        .txtIPin = rsActiveTable("SMSIPin") & ""
        .txtCreditLimit = Format(rsActiveTable("CreditLimit"), "STANDARD")
        .txtMinimumPurchase = Format(rsActiveTable("MinimumPurchase"), "STANDARD")
        
        .cmdSave.Visible = False
        ClearRS rsActiveTable
        
        .Show 1
    End With
    Show
End Sub


Private Sub Form_Load()
    Dim ictr%
    
    CenterFrm Me
    
    For ictr = 0 To 4
        msfPendingSOHeader.ColWidth(ictr) = Me.lblHeader(ictr).Width
    Next
    
    For ictr = 5 To 9
        msfSODetails.ColWidth(ictr - 5) = Me.lblHeader(ictr).Width
    Next
    
    cboFilterStatus.ListIndex = 0
    RefreshHeader
        
End Sub

Private Sub RefreshHeader()
    Dim rsTransactions As New ADODB.Recordset
    Dim strSQL$
    
    Select Case cboFilterStatus.ListIndex
        Case 0:
            strSQL = "EXEC spPendingSO"
        Case 1:
            strSQL = "EXEC spForShipping"
        Case Else
            strSQL = "SELECT * FROM qryTransactions WHERE ShippedDate Is Null AND EmployeeID = 0 ORDER BY OrderDate"
    End Select
    
    
    rsTransactions.Open strSQL, connRFM, adOpenStatic, adLockReadOnly
    
    msfPendingSOHeader.Rows = 0
    
    If Not NoRecord(rsTransactions) Then
        While Not rsTransactions.EOF
            If Not IsDeclined(rsTransactions("TransactionID")) Then msfPendingSOHeader.AddItem rsTransactions("TransactionID") & vbTab & rsTransactions("CustomerID") & vbTab & Format(rsTransactions("OrderDate"), "medium date") & vbTab & Format(rsTransactions("RequiredDate"), "medium date") & vbTab & rsTransactions("OrderSource")
            rsTransactions.MoveNext
        Wend
        
    End If
    
    msfPendingSOHeader.AddItem ""
    
    msfSODetails.Clear
    msfSODetails.Rows = 1
    
    lblRecordCount = msfPendingSOHeader.Rows - 1 & " returned"
    ResizeGrid msfPendingSOHeader, 8595, rsTransactions.RecordCount
    
End Sub

Private Sub RefreshDetails()
    Dim rsTransactionDetails As New ADODB.Recordset
    Dim dTotal As Double
    
    
    rsTransactionDetails.Open "EXEC spTransactionDetails " & msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0), connRFM, adOpenStatic
    
    
    msfSODetails.Rows = 0
    
    If Not NoRecord(rsTransactionDetails) Then
        While Not rsTransactionDetails.EOF
            msfSODetails.AddItem rsTransactionDetails("ProductID") & vbTab & rsTransactionDetails("ProductName") & vbTab & Format(rsTransactionDetails("UnitPrice"), "STANDARD") & vbTab & rsTransactionDetails("Quantity") & vbTab & Format(rsTransactionDetails("Quantity") * rsTransactionDetails("UnitPrice"), "STANDARD")
            dTotal = dTotal + rsTransactionDetails("Quantity") * rsTransactionDetails("UnitPrice")
            rsTransactionDetails.MoveNext
        Wend
        
        lblTotal = "Sales Order Total: " & Format(dTotal, "STANDARD")
        msfSODetails.AddItem ""
    End If
    
    lblRecordCount = msfPendingSOHeader.Row + 1 & " of " & msfPendingSOHeader.Rows - 1 & " records"
    ResizeGrid msfSODetails, 9300, rsTransactionDetails.RecordCount
    
End Sub


Private Sub msfPendingSOHeader_EnterCell()
    If msfPendingSOHeader.TextMatrix(msfPendingSOHeader.Row, 0) = "" Then
        msfSODetails.Rows = 0
        msfSODetails.AddItem ""
        lblTotal = "Sales Order Total: 0.00"
        Exit Sub
    End If
    
    RefreshDetails
End Sub


Private Function IsDeclined(TransactionID As Long) As Boolean
    Dim rsDeclined As New ADODB.Recordset
    
    rsDeclined.Open "SELECT TransactionID FROM tblDeclinedTransactions WHERE TransactionID =" & TransactionID, connRFM
    
    If Not NoRecord(rsDeclined) Then
        IsDeclined = True
    Else
        IsDeclined = False
    End If
    
    ClearRS rsDeclined
End Function

Private Function GetNumber(CustomerID As String) As String
    Dim rsCustomers As New ADODB.Recordset
    
    rsCustomers.Open "SELECT SMSNumber,CustomerID FROM tblCustomers WHERE CustomerID = " & Enquote(CustomerID), connRFM, adOpenStatic
    GetNumber = rsCustomers("SMSNumber") & ""
    
    ClearRS rsCustomers
    
End Function
