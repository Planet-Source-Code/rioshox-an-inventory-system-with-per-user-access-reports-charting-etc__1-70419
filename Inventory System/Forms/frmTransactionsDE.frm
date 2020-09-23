VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTransactionsDE 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optPickupDel 
      BackColor       =   &H00F5EADB&
      Caption         =   "Delivery"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4125
      TabIndex        =   35
      Top             =   2790
      Width           =   1275
   End
   Begin VB.OptionButton optPickupDel 
      BackColor       =   &H00F5EADB&
      Caption         =   "Pickup"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3165
      TabIndex        =   34
      Top             =   2790
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox txtProductID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   195
      TabIndex        =   13
      Top             =   3660
      Width           =   1320
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   6360
      TabIndex        =   19
      Top             =   3660
      Width           =   885
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   7320
      TabIndex        =   21
      Top             =   3660
      Width           =   930
   End
   Begin VB.TextBox txtProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCF8F4&
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3660
      Width           =   3480
   End
   Begin VB.CommandButton cmdAddLine 
      BackColor       =   &H00FFFEFC&
      Caption         =   "Add Line"
      Height          =   345
      Left            =   8415
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3615
      Width           =   1065
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FCF8F4&
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3660
      Width           =   1275
   End
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
      Height          =   345
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Click this to refresh products and customers table. This prevents error in SO data entry."
      Top             =   690
      Width           =   1365
   End
   Begin VB.CommandButton cmdFindProduct 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   -480
      Picture         =   "frmTransactionsDE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7710
      Width           =   870
   End
   Begin VB.CheckBox chkPostNow 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5EADB&
      Caption         =   "Post This Transaction Now"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6495
      TabIndex        =   29
      Top             =   705
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CommandButton cmdDeleteLine 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete This Line"
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7005
      Width           =   1605
   End
   Begin VB.CommandButton cmdNewTransaction 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Transaction"
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
      Left            =   1965
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7005
      Width           =   1590
   End
   Begin VB.CommandButton cmdAddtransaction 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add &Transaction"
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
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7005
      Width           =   1515
   End
   Begin VB.TextBox txtRequiredDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FCF8F4&
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   1605
      TabIndex        =   11
      Top             =   2790
      Width           =   1245
   End
   Begin VB.TextBox txtOrderDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FCF8F4&
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   180
      TabIndex        =   9
      Top             =   2790
      Width           =   1245
   End
   Begin VB.CommandButton cmdCheckBalance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check Balance"
      Height          =   345
      Left            =   4785
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1485
      Width           =   1365
   End
   Begin VB.ComboBox cboCustomers 
      ForeColor       =   &H007D584F&
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   1500
      Width           =   1635
   End
   Begin VB.TextBox txtContactNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCF8F4&
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   4635
      TabIndex        =   7
      Top             =   2145
      Width           =   1245
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCF8F4&
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   2145
      Width           =   4290
   End
   Begin VB.TextBox txtContactName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCF8F4&
      ForeColor       =   &H007D584F&
      Height          =   285
      Left            =   1980
      TabIndex        =   3
      Top             =   1500
      Width           =   2670
   End
   Begin MSFlexGridLib.MSFlexGrid msfTransactionDetails 
      Height          =   2655
      Left            =   180
      TabIndex        =   23
      Top             =   4005
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   8214607
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Data Entry"
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
      Left            =   90
      TabIndex        =   33
      Top             =   165
      Width           =   4110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Product ID"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   210
      TabIndex        =   12
      Top             =   3420
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Quantity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   18
      Top             =   3420
      Width           =   810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   7320
      TabIndex        =   20
      Top             =   3420
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   1515
      TabIndex        =   14
      Top             =   3420
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pri&ce"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   16
      Top             =   3435
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Required Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   1605
      TabIndex        =   10
      Top             =   2550
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   180
      TabIndex        =   8
      Top             =   2550
      Width           =   1335
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   7935
      MouseIcon       =   "frmTransactionsDE.frx":0102
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   6735
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   4635
      TabIndex        =   6
      Top             =   1905
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   1905
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1980
      TabIndex        =   2
      Top             =   1230
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   1230
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   2670
      Left            =   60
      Top             =   570
      Width           =   9645
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   4200
      Left            =   60
      Top             =   3360
      Width           =   9630
   End
End
Attribute VB_Name = "frmTransactionsDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomers As New ADODB.Recordset
Dim rsProducts As New ADODB.Recordset

Dim intRows%

Const NORMAL_GRID_WIDTH As Integer = 9315


Private Sub cboCustomers_Change()
    ComboChange cboCustomers
End Sub

Private Sub cboCustomers_KeyDown(KeyCode As Integer, Shift As Integer)
    ComboKeyDown KeyCode
End Sub

Private Sub cboCustomers_LostFocus()
    rsCustomers.MoveFirst
    rsCustomers.Find "CustomerID = " & Enquote(cboCustomers)
    If Not rsCustomers.EOF Then
        txtContactName = rsCustomers("ContactName") & ""
        'If cboCustomers <> GENERAL_CUSTOMER Then
        '    txtAddress = rsCustomers("CustAddress") & ""
        '    txtContactNo = rsCustomers("Phone") & ""
        '    chkPostNow.Value = Unchecked
        'Else
            txtAddress = ""
            txtContactNo = ""
            'chkPostNow.Value = Checked
        'End If
        txtOrderDate.SetFocus
        
    Else
        MsgBox "Invalid customer ID...", vbExclamation
        cboCustomers.SetFocus
    End If
    
    
    
    
    
    
End Sub

Private Sub cmdAddLine_Click()
    
    
    If Not IsNumeric(txtProductID) Then
        MsgBox "Invalid product ID.", vbExclamation
        txtProductID.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtQty) Then
        MsgBox "Invalid quantity.", vbExclamation
        txtQty.Text = 0
        txtQty.SetFocus
        Exit Sub
    End If
    
    
    
    If CDbl(txtQty) <= 0 Then
        MsgBox "Invalid quantity.", vbExclamation
        txtQty.Text = 0
        txtQty.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtDiscount) Then
        MsgBox "Invalid discount.", vbExclamation
        txtDiscount.Text = 0
        txtDiscount.SetFocus
        Exit Sub
    End If
    
    
    
    If CDbl(txtDiscount) > CDbl(txtPrice) Then
        MsgBox "Cannot apply discount!", vbExclamation
        txtDiscount = 0
        Exit Sub
    End If
    
    If (CDbl(txtDiscount) > 1) Or (CDbl(txtDiscount) > 1) Then
        MsgBox "Invalid discount, must between 0 to 1 only.", vbExclamation
        txtDiscount.Text = 0
        txtDiscount.SetFocus
        Exit Sub
    End If
    
    If isProductInList(txtProductID, txtQty) Then
        GoTo New_Entry
    End If
    
    Dim Discount$
    
    msfTransactionDetails.TextMatrix(intRows%, 0) = txtProductID
    msfTransactionDetails.TextMatrix(intRows%, 1) = txtProductName
    msfTransactionDetails.TextMatrix(intRows%, 2) = TwoDecimals(txtPrice)
    
    msfTransactionDetails.TextMatrix(intRows%, 3) = txtQty
    msfTransactionDetails.TextMatrix(intRows%, 4) = txtDiscount
    msfTransactionDetails.TextMatrix(intRows%, 5) = TwoDecimals((Val(txtPrice) * CDbl(txtQty) * (1 - txtDiscount) / 100) * 100)
     
    msfTransactionDetails.Rows = msfTransactionDetails.Rows + 1
    lblSubTotal.Caption = Format(lblSubTotal.Caption + CDbl(msfTransactionDetails.TextMatrix(intRows%, 5)), "STANDARD")
    
    
    
    intRows% = intRows% + 1
    
    ResizeGrid msfTransactionDetails, NORMAL_GRID_WIDTH, intRows%
    
New_Entry:
    txtProductID = ""
    txtProductName = ""
    txtPrice = ""
    
    txtQty = "0"
    txtDiscount = "0"
    
    txtProductID.SetFocus
End Sub


Private Sub cmdAddtransaction_Click()
    If Not IsDate(txtOrderDate) Then
        MsgBox "Invalid order date...", vbExclamation
        txtOrderDate = Date
        txtOrderDate.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtRequiredDate) Then
        MsgBox "Invalid required date...", vbExclamation
        txtRequiredDate = Date
        txtRequiredDate.SetFocus
        Exit Sub
    End If
    
    
    If msfTransactionDetails.Rows = 1 Then
        MsgBox "Cannot add transactio without details!", vbExclamation
        txtProductID.SetFocus
        Exit Sub
    End If
    
    
    AddTransaction
End Sub

Private Sub cmdAddTransactionShort_Click()
    cmdAddtransaction_Click
End Sub

Private Sub cmdCheckBalance_Click()
    'If cboCustomers <> GENERAL_CUSTOMER Then
        'hGlass True
        'MsgBox cboCustomers & "'s balance is " & Format(GetCustomerBalance(cboCustomers), "STANDARD"), vbInformation
        'hGlass False
    'End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteLine_Click()
    If msfTransactionDetails.Rows <> 1 Then
        If msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 0) <> "" Then
            lblSubTotal = CDbl(lblSubTotal) - CDbl(msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 5))
            msfTransactionDetails.RemoveItem msfTransactionDetails.Row
            intRows = intRows - 1
        Else
            MsgBox "Cannot remove selected row....", vbExclamation
        End If
    End If
End Sub

Private Sub cmdFindProduct_Click()
    'PrevValue = txtProductID
    'frmFindProduct.Show 1
    'txtProductID = PrevValue
    'txtProductID.SetFocus
End Sub

Private Sub cmdNewTransaction_Click()
    ResetValues
End Sub



Private Sub cmdRefresh_Click()
    If rsCustomers.State = adStateOpen Then rsCustomers.Close
    
    rsCustomers.Open "SELECT CustomerID,CreditLimit,ContactName, Address + ', ' + City + PostalCode as CustAddress, Phone FROM tblCustomers ORDER BY CustomerID", connRFM, adOpenStatic, adLockReadOnly
    PopulateCboBox rsCustomers, "CustomerID", cboCustomers
    
    If rsProducts.State = adStateOpen Then rsProducts.Close
    rsProducts.Open "SELECT ProductID,ProductName ,  PackagingType ,UnitPrice,UnitsOnOrder FROM qrySearchProduct", connRFM, adOpenStatic, adLockReadOnly
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    hGlass True
    
    rsCustomers.CursorLocation = adUseClient
    rsProducts.CursorLocation = adUseClient
    
    cmdRefresh_Click
    
    CenterFrm Me
    
    msfTransactionDetails.ColWidth(0) = txtProductID.Width + 10
    msfTransactionDetails.ColWidth(1) = txtProductName.Width + 10
    msfTransactionDetails.ColWidth(2) = txtPrice.Width + 10
    msfTransactionDetails.ColWidth(3) = txtQty.Width + 10
    msfTransactionDetails.ColWidth(4) = txtDiscount.Width + 10
    msfTransactionDetails.ColWidth(5) = cmdAddLine.Width + 35
    
    msfTransactionDetails.ColAlignment(0) = 7
    msfTransactionDetails.ColAlignment(1) = 1
    msfTransactionDetails.ColAlignment(2) = 7
    msfTransactionDetails.ColAlignment(3) = 7
    msfTransactionDetails.ColAlignment(4) = 7
    msfTransactionDetails.ColAlignment(5) = 7
    
    txtOrderDate = Date
    txtRequiredDate = Date
    intRows% = 0
    
    
    
    
    hGlass False
End Sub


Private Sub txtDateDelivered_GotFocus()
    HighlightMe
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ClearRS rsCustomers
    ClearRS rsProducts
End Sub








Private Sub lblSubTotal_Click()
    Dim OldTotal#, Disc#, ictr%
    
    If hasDiscount Then
        MsgBox "Cannot use transaction discount wizard if an item already has a discount.", vbExclamation
        Exit Sub
    End If
    'g_currentuser.AuthorizeDiscount = False
    
    frmAuthorizeDiscountBatch.Show 1
    
    If g_CurrentUser.AuthorizeDiscount Then
        Disc# = lblSubTotal.Tag
        OldTotal# = CDbl(lblSubTotal.Caption)
        lblSubTotal.Caption = 0
    
        For ictr% = 0 To intRows - 1
            msfTransactionDetails.TextMatrix(ictr%, 4) = (msfTransactionDetails.TextMatrix(ictr%, 5) / (OldTotal#)) * Disc# / msfTransactionDetails.TextMatrix(ictr%, 5)
            msfTransactionDetails.TextMatrix(ictr%, 5) = (Val(msfTransactionDetails.TextMatrix(ictr%, 2)) * msfTransactionDetails.TextMatrix(ictr%, 3) * (1 - msfTransactionDetails.TextMatrix(ictr%, 4)) / 100) * 100
            lblSubTotal = CDbl(lblSubTotal) + CDbl(msfTransactionDetails.TextMatrix(ictr%, 5))
            msfTransactionDetails.TextMatrix(ictr%, 4) = TwoDecimals(msfTransactionDetails.TextMatrix(ictr%, 4))
            msfTransactionDetails.TextMatrix(ictr%, 5) = TwoDecimals(msfTransactionDetails.TextMatrix(ictr%, 5))
        Next
        
        lblSubTotal = Format(lblSubTotal, "STANDARD")
    
    End If
    
    lblSubTotal.Tag = 0
End Sub


Private Sub txtDiscount_GotFocus()
    HighlightMe
End Sub

Private Sub txtDiscount_DblClick()
    LineDiscount = 0
    With frmAuthorizeDiscount
        .txtDiscount.Tag = txtQty
        .txtPassword.Tag = txtPrice
        .Show 1
    End With
    txtDiscount = LineDiscount
End Sub


Private Sub txtDiscount_LostFocus()
    txtDiscount.Locked = True
End Sub



Private Sub txtOrderDate_GotFocus()
    HighlightMe
End Sub

Private Sub txtProductID_GotFocus()
    HighlightMe
End Sub

Private Sub txtProductID_LostFocus()
    On Error GoTo txtProductID_LostFocus_ERR
    If Not IsNumeric(txtProductID.Text) Then Exit Sub
    
    rsProducts.MoveFirst
    rsProducts.Find "ProductID = " & txtProductID
    If Not rsProducts.EOF Then
        txtProductName = rsProducts("ProductName")
        txtPrice = rsProducts("UnitPrice")

        txtQty.Text = 1
        txtDiscount = 0
        
        txtQty.SetFocus
        
    Else
        MsgBox "Invalid product id...", vbExclamation
        txtProductID.SetFocus
    End If
    
txtProductID_LostFocus_ERR:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        txtProductID.SetFocus
    End If
End Sub

Private Sub txtProductName_GotFocus()
    HighlightMe
End Sub


Private Sub txtQty_GotFocus()
    HighlightMe
End Sub

Private Sub txtRequiredDate_GotFocus()
    HighlightMe
End Sub

Private Sub AddTransaction()
    Dim rsTransactions As ADODB.Recordset
    Dim rsTransactionDetails As ADODB.Recordset
    Dim ictr%, CurrentTransaction As Long
    Dim PayMode%, BankName$, CheckNo$, CheckDate$
    
    'If chkPostNow.Value = Checked Then
    '    For ictr% = 0 To intRows - 1
    '        If Not isStockEnough(msfTransactionDetails.TextMatrix(ictr%, 0), msfTransactionDetails.TextMatrix(ictr%, 3)) Then
    '            MsgBox "Cannot post, stock for " & msfTransactionDetails.TextMatrix(ictr%, 1) & " is not enough!", vbExclamation
    '            Exit Sub
    '        End If
    '    Next
    'End If
    
    On Error GoTo AddTransaction_ERR
    
    
     
    connRFM.BeginTrans
        Set rsTransactions = New ADODB.Recordset
        Set rsTransactionDetails = New ADODB.Recordset
        
        rsTransactions.Open "tblTransactions", connRFM, adOpenKeyset, adLockOptimistic
        rsTransactionDetails.Open "tblTransactionDetails", connRFM, adOpenKeyset, adLockOptimistic
    
        rsTransactions.AddNew
            rsTransactions("CustomerID") = cboCustomers
            rsTransactions("OrderDate") = txtOrderDate
            rsTransactions("RequiredDate") = txtRequiredDate
            rsTransactions("EmployeeID") = g_CurrentUser.EmployeeID
            If chkPostNow.Value = Checked Then
                rsTransactions("DeliveredDate") = Date
            End If
            rsTransactions("IsPickup") = optPickupDel(0).Value
        rsTransactions.Update
        
        CurrentTransaction = rsTransactions("TransactionID")
        
        For ictr% = 0 To intRows - 1
            rsTransactionDetails.AddNew
                rsTransactionDetails("TransactionID") = CurrentTransaction
                rsTransactionDetails("ProductID") = msfTransactionDetails.TextMatrix(ictr%, 0)
                rsTransactionDetails("UnitPrice") = msfTransactionDetails.TextMatrix(ictr%, 2)
                rsTransactionDetails("Quantity") = msfTransactionDetails.TextMatrix(ictr%, 3)
                rsTransactionDetails("Discount") = msfTransactionDetails.TextMatrix(ictr%, 4)
                
                'Update stocks/units on order
                If chkPostNow.Value = Checked Then
                    connRFM.Execute _
                        "UPDATE tblProducts SET UnitsInStock = UnitsInStock - " & CDbl(msfTransactionDetails.TextMatrix(ictr%, 3)) & " " & _
                        "WHERE ProductID = " & msfTransactionDetails.TextMatrix(ictr%, 0)
                Else
                    'AS of 4/29/2006
                    'connRFM.Execute _
                        "UPDATE tblProducts SET UnitsOnOrder = UnitsOnOrder - " & CDbl(msfTransactionDetails.TextMatrix(ictr%, 3)) & " " & _
                        "WHERE ProductID = " & msfTransactionDetails.TextMatrix(ictr%, 0)
                    connRFM.Execute _
                        "UPDATE tblProducts SET UnitsOnOrder = UnitsOnOrder + " & CDbl(msfTransactionDetails.TextMatrix(ictr%, 3)) & " " & _
                        "WHERE ProductID = " & msfTransactionDetails.TextMatrix(ictr%, 0)
                    
                End If
                'rsProducts.Update
                
                
            rsTransactionDetails.Update
        Next
        
        connRFM.CommitTrans
        
        MsgBox "Transaction added with transaction No.: " & CurrentTransaction, vbInformation
        
        If chkPostNow.Value = Checked Then
            If MsgBox("Would you like to print the sales invoice?", vbQuestion + vbYesNo) = vbYes Then
                ShowReport "rptInvoice", "TransactionID = " & CurrentTransaction
            End If
        End If
        
        If chkPostNow.Value = Checked Then
            If MsgBox("Update payment details now?", vbYesNo + vbQuestion) = vbNo Then
                ResetValues
                Exit Sub
            End If
            
            Dim AmountTendered As String
            
            Do While Not IsNumeric(AmountTendered)
                AmountTendered = InputBox("Enter initial payment...", "Payment Entry", lblSubTotal.Caption)
                If Trim(AmountTendered) = "" Then
                    ResetValues
                    Exit Sub
                End If
                
                If IsNumeric(AmountTendered) Then
                    If CDbl(AmountTendered) > CDbl(lblSubTotal) Then
                        MsgBox "Amount tendered cannot be greater than the total of the transaction!", vbExclamation
                        AmountTendered = "xxx"
                    End If
                End If
            Loop
            
            PayMode% = MsgBox("Select payment mode: " & vbCrLf & "Yes - Cash" & vbCrLf & "No - Check", vbYesNoCancel + vbQuestion)
    
            If PayMode% = vbCancel Then Exit Sub
    
            If PayMode% = vbNo Then
                BankName$ = InputBox("Enter bank name...", "Payment Entry for Current Transaction", "Enter Bank name here")
                If BankName$ = "" Then
                    ResetValues
                    Exit Sub
                End If
                
                CheckNo$ = InputBox("Enter check no...", "Payment Entry for Current Transaction ", "Enter check no here")
                If CheckNo$ = "" Then
                    ResetValues
                    Exit Sub
                End If
                
                Do While Not IsDate(CheckDate$)
                    CheckDate$ = InputBox("Enter check date ...", "Payment Entry for Current Transaction ", Date)
                    If CheckDate$ = "" Then
                        ResetValues
                        Exit Sub
                    End If
                Loop
            End If
            

        End If
    
    ResetValues
    
AddTransaction_ERR:
    If Err.Number <> 0 Then
        connRFM.RollbackTrans
        MsgBox "Cannot add transaction, errors occurred." & vbCrLf & Err.Description, vbCritical
    End If
End Sub


Private Sub ResetValues()
    msfTransactionDetails.Clear
    msfTransactionDetails.Rows = 1
    
    intRows = 0
    cboCustomers.SetFocus
    cboCustomers.ListIndex = 0
    
    txtOrderDate = Date
    txtRequiredDate = Date
    
    txtContactName = ""
    txtAddress = ""
    txtContactNo = ""

    cboCustomers.SetFocus
    
    chkPostNow.Value = Unchecked
    lblSubTotal.Caption = "0.00"
End Sub

Private Function isProductInList(ProductID As Integer, Qty As Double)
    Dim ictr%
    
    On Error GoTo isProductInList_ERR
    If intRows% = 0 Then
        isProductInList = False
        Exit Function
    End If
    
    For ictr% = 0 To intRows% - 1
        If ProductID = msfTransactionDetails.TextMatrix(ictr, 0) Then
            isProductInList = True
            If MsgBox("Product is already in the list." & vbCrLf & "Clicking YES will add the qty to the existing qty in the list. (Note that the discount of the latest entry for the product will be used)", vbQuestion + vbYesNo) = vbYes Then
                msfTransactionDetails.TextMatrix(ictr, 3) = CDbl(msfTransactionDetails.TextMatrix(ictr, 3)) + Qty
                lblSubTotal.Caption = TwoDecimals(CDbl(lblSubTotal.Caption) - msfTransactionDetails.TextMatrix(ictr, 5))
                msfTransactionDetails.TextMatrix(ictr, 5) = TwoDecimals((Val(txtPrice) * msfTransactionDetails.TextMatrix(ictr, 3) * (1 - txtDiscount) / 100) * 100)
                lblSubTotal.Caption = TwoDecimals(CDbl(lblSubTotal.Caption) + msfTransactionDetails.TextMatrix(ictr, 5))
                Exit Function
            End If
        End If
        
    Next

isProductInList_ERR:
    If Err.Number <> 0 Then MsgBox "Errors occurred." & vbCrLf & Err.Description, vbCritical
End Function

Private Function hasDiscount() As Boolean
    Dim ictr%
    
    For ictr% = 0 To intRows - 1
        If CDbl(msfTransactionDetails.TextMatrix(ictr%, 4)) > 0 Then
            hasDiscount = True
            Exit Function
        End If
    Next
    
    hasDiscount = False
End Function
