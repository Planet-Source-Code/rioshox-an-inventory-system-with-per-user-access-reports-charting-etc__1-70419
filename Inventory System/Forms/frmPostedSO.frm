VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPostedSO 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   Caption         =   "a"
   ClientHeight    =   9330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrintInvoice 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print Invoice"
      Height          =   345
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4455
      Width           =   1590
   End
   Begin VB.TextBox txtStartDate 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   690
      Width           =   1830
   End
   Begin VB.TextBox txtEndDate 
      Height          =   285
      Left            =   2115
      TabIndex        =   4
      Top             =   690
      Width           =   1830
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show"
      Height          =   330
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   660
      Width           =   960
   End
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   330
      Left            =   9330
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8880
      Width           =   960
   End
   Begin VB.CommandButton cmdViewCustomer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Customer"
      Height          =   345
      Left            =   3675
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4455
      Width           =   1590
   End
   Begin VB.CommandButton cmdPrintOrderSlip 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print SO Slip"
      Height          =   345
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid msfPendingSOHeader 
      Height          =   2715
      Left            =   360
      TabIndex        =   9
      Top             =   1635
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   4789
      _Version        =   393216
      Rows            =   7
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14663855
      GridColor       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
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
      Left            =   570
      TabIndex        =   10
      Top             =   5715
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
      Height          =   270
      Left            =   5700
      TabIndex        =   23
      Top             =   4380
      Width           =   3225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   240
      Index           =   2
      Left            =   255
      TabIndex        =   1
      Top             =   465
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   240
      Index           =   1
      Left            =   2130
      TabIndex        =   3
      Top             =   465
      Width           =   780
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Posted Orders"
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
      Left            =   285
      TabIndex        =   0
      Top             =   165
      Width           =   3735
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
      Left            =   345
      TabIndex        =   22
      Top             =   1305
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
      Left            =   2145
      TabIndex        =   21
      Top             =   1305
      Width           =   1920
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
      Left            =   6000
      TabIndex        =   20
      Top             =   1305
      Width           =   2040
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
      Left            =   8040
      TabIndex        =   19
      Top             =   1305
      Width           =   1860
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
      Left            =   4065
      TabIndex        =   18
      Top             =   1305
      Width           =   1935
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
      Left            =   7950
      TabIndex        =   17
      Top             =   5415
      Width           =   1890
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
      Left            =   6870
      TabIndex        =   16
      Top             =   5415
      Width           =   1080
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
      Left            =   5685
      TabIndex        =   15
      Top             =   5415
      Width           =   1200
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
      Left            =   2205
      TabIndex        =   14
      Top             =   5415
      Width           =   3480
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
      Left            =   570
      TabIndex        =   13
      Top             =   5415
      Width           =   1635
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
      Index           =   0
      Left            =   420
      TabIndex        =   12
      Top             =   5085
      Width           =   2250
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
      Left            =   6015
      TabIndex        =   11
      Top             =   8430
      Width           =   3825
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3795
      Left            =   165
      Top             =   4980
      Width           =   10140
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3765
      Left            =   165
      Top             =   1125
      Width           =   10155
   End
End
Attribute VB_Name = "frmPostedSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()

End Sub

Private Sub cmdClose_Click()
    Unload Me
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
    Dim tmpDate$
    
    If Not IsDate(txtStartDate) Then
        txtStartDate.SetFocus
        MsgBox "Invalid start date!", vbCritical
        Exit Sub
    End If
    
    
    If Not IsDate(txtEndDate) Then
        txtEndDate.SetFocus
        MsgBox "Invalid end date!", vbCritical
        Exit Sub
    End If
    
    If CDate(txtEndDate) < CDate(txtStartDate) Then
        tmpDate$ = txtEndDate
        txtEndDate = txtStartDate
        txtStartDate = tmpDate$
    End If
    
    hGlass True
    RefreshHeader
    hGlass False
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


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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
    
    txtStartDate = Date - 7
    txtEndDate = Date
    
    RefreshHeader
    
    
        
End Sub

Private Sub RefreshHeader()
    Dim rsTransactions As New ADODB.Recordset
    Dim strSQL$
    
    rsTransactions.Open "EXEC spPostedSO " & Enquote(txtStartDate) & "," & Enquote(txtEndDate), connRFM, adOpenStatic
    
    msfPendingSOHeader.Rows = 0
    
    If Not NoRecord(rsTransactions) Then
        While Not rsTransactions.EOF
            msfPendingSOHeader.AddItem rsTransactions("TransactionID") & vbTab & rsTransactions("CustomerID") & vbTab & Format(rsTransactions("OrderDate"), "medium date") & vbTab & Format(rsTransactions("RequiredDate"), "medium date") & vbTab & rsTransactions("OrderSource")
            rsTransactions.MoveNext
        Wend
        
    End If
    
    msfPendingSOHeader.AddItem ""
    
    lblRecordCount = msfPendingSOHeader.Rows - 1 & " returned"
    
    
    
    msfSODetails.Clear
    msfSODetails.Rows = 1
    
     
    
    ResizeGrid msfPendingSOHeader, 9555, rsTransactions.RecordCount
    
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


Private Sub txtEndDate_GotFocus()
    HighlightMe
End Sub

Private Sub txtEndDate_LostFocus()
    If IsDate(txtEndDate) Then txtEndDate = Format(txtEndDate, "mm/dd/yyyy")
End Sub

Private Sub txtStartDate_GotFocus()
    HighlightMe
End Sub

Private Sub txtStartDate_LostFocus()
    If IsDate(txtStartDate) Then txtStartDate = Format(txtStartDate, "mm/dd/yyyy")
End Sub
