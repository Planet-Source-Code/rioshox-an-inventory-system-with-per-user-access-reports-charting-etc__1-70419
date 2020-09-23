VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReturnItems 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtOrderDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
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
      Height          =   315
      Left            =   5565
      TabIndex        =   15
      Top             =   8490
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdEditLine 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Line"
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
      Left            =   9855
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2115
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid msfTransactionDetails 
      Height          =   2145
      Left            =   165
      TabIndex        =   4
      Top             =   2040
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   3784
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   8214607
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
   Begin VB.CommandButton cmdReturnItem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return Item"
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
      Left            =   9855
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2595
      Width           =   1275
   End
   Begin VB.CommandButton cmdReplaceItem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Replace Item"
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
      Left            =   9870
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2985
      Width           =   1275
   End
   Begin VB.CommandButton cmdClose 
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
      Height          =   345
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   900
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   885
      Width           =   1755
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Search"
      Default         =   -1  'True
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
      Left            =   2010
      MaskColor       =   &H006F453A&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   885
      Width           =   930
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Items"
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
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   2250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Items On The Sales Order"
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
      Height          =   285
      Index           =   0
      Left            =   195
      TabIndex        =   16
      Top             =   1410
      Width           =   3480
   End
   Begin VB.Label lblCaption 
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   5
      Left            =   8430
      TabIndex        =   13
      Top             =   1755
      Width           =   1335
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   1080
      TabIndex        =   12
      Top             =   1755
      Width           =   4335
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   5415
      TabIndex        =   11
      Top             =   1755
      Width           =   1020
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quanity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   6435
      TabIndex        =   10
      Top             =   1755
      Width           =   945
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   4
      Left            =   7380
      TabIndex        =   9
      Top             =   1755
      Width           =   1050
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   165
      TabIndex        =   8
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Height          =   270
      Left            =   8490
      TabIndex        =   7
      Top             =   165
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Transaction No.:"
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
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   630
      Width           =   2295
   End
   Begin VB.Shape f 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3825
      Left            =   60
      Top             =   510
      Width           =   11145
   End
End
Attribute VB_Name = "frmReturnItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NORMAL_GRID_WIDTH As Integer = 9600

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdReplaceItem_Click()
    Dim Qty$
    
    If Not IsNumeric(msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 0)) Then Exit Sub
    
    If MsgBox("If you have checked the item and a replacement is really needed, click yes to continue...", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    On Error GoTo cmdReplaceItem_Click_ERR
    
    Do While Not IsNumeric(Qty$)
        Qty$ = InputBox("Enter qty to be replaced...", "Replace Item", CInt(msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 3)))
        If Qty$ = "" Then Exit Sub
        If IsNumeric(Qty$) Then
            If CDbl(Qty$) > CDbl(msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 3)) Then
                MsgBox "Return quantity cannot be greater than the original quantity!", vbExclamation
                Qty$ = "xxx"
            End If
        End If
    Loop
    
    connRFM.Execute _
        "UPDATE tblProducts SET UnitsInStock = UnitsInStock - " & Qty & _
        " WHERE ProductID = " & CDbl(msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 0)), adExecuteNoRecords
    
    MsgBox "Inventory updated!", vbInformation
    
cmdReplaceItem_Click_ERR:
        If Err.Number <> 0 Then
            MsgBox "Errors occurred, cannot continue current operation..." & vbCrLf & Err.Description, vbCritical
        End If
End Sub

Private Sub cmdReturnItem_Click()
    Dim Qty$
    
    If msfTransactionDetails.Rows <> 1 Then
        If msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 0) <> "" Then
            
            On Error GoTo cmdDeleteLine_Click_ERR
            
            Do While Not IsNumeric(Qty$)
                Qty$ = InputBox("Enter qty to return...", "Return Item", CInt(msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 3)))
                If Qty$ = "" Then Exit Sub
                If IsNumeric(Qty$) Then
                    If CDbl(Qty$) > CDbl(msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 3)) Then
                        MsgBox "Return quantity cannot be greater than the original quantity!", vbExclamation
                        Qty$ = "xxx"
                    End If
                End If
            Loop
            
            connRFM.Execute _
                "UPDATE tblProducts SET UnitsInStock = UnitsInStock + " & Qty & _
                " WHERE ProductID = " & CDbl(msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 0)), adExecuteNoRecords
            
            connRFM.Execute _
                "UPDATE tblTransactionDetails SET Quantity = Quantity - " & Qty & _
                " WHERE TransactionID = " & txtSearch & " AND ProductID = " & CDbl(msfTransactionDetails.TextMatrix(msfTransactionDetails.Row, 0)), adExecuteNoRecords
            
                    
            ResizeGrid msfTransactionDetails, NORMAL_GRID_WIDTH, msfTransactionDetails.Rows
                   
            cmdSearch_Click
        Else
            MsgBox "Cannot remove selected row....", vbExclamation
        End If
    End If
    
    
    
cmdDeleteLine_Click_ERR:
    If Err.Number <> 0 Then
        MsgBox "Errors occurred, cannot delete line." & vbCrLf & Err.Description, vbCritical
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim rsTransaction As ADODB.Recordset
    
    ResetForm
    
    If Not IsNumeric(txtSearch) Then
        txtSearch.Text = ""
        txtSearch.SetFocus
        Exit Sub
    End If
    
    msfTransactionDetails.Clear
    msfTransactionDetails.Rows = 0
    

    
    Set rsTransaction = New ADODB.Recordset
    rsTransaction.CursorLocation = adUseClient
    rsTransaction.Open _
         "SELECT tblTransactions.CustomerID, tblTransactions.TransactionID, tblTransactions.OrderDate, tblTransactionDetails.ProductID, tblProducts.ProductName, tblTransactionDetails.UnitPrice, tblTransactionDetails.Quantity, tblTransactionDetails.Discount, ([tblTransactionDetails].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100 AS ExtendedPrice " & _
        "FROM tblTransactions INNER JOIN (tblProducts INNER JOIN tblTransactionDetails ON tblProducts.ProductID = tblTransactionDetails.ProductID) ON tblTransactions.TransactionID = tblTransactionDetails.TransactionID " & _
        "WHERE (((tblTransactions.TransactionID)=" & txtSearch & ")) " & _
        "ORDER BY ProductName", connRFM
    
    If NoRecord(rsTransaction) Then
        MsgBox "No transaction with that Transaction number found.", vbExclamation
        txtSearch = ""
        txtSearch.SetFocus
    Else
        
        txtOrderDate = rsTransaction("OrderDate")
        
        If Not NoRecord(rsTransaction) Then
            rsTransaction.MoveFirst
            txtSearch.Tag = rsTransaction("CustomerID")
            While Not rsTransaction.EOF
                msfTransactionDetails.AddItem rsTransaction("ProductID") & vbTab & rsTransaction("ProductName") & vbTab & Format(rsTransaction("UnitPrice"), "STANDARD") & vbTab & rsTransaction("Quantity") & vbTab & rsTransaction("Discount") & vbTab & Format(rsTransaction("ExtendedPrice"), "STANDARD")
                lblSubTotal = Format(CDbl(lblSubTotal) + rsTransaction("ExtendedPrice"), "STANDARD")
                rsTransaction.MoveNext
            Wend
            
        End If
    
    msfTransactionDetails.AddItem ""
    
    ResizeGrid msfTransactionDetails, NORMAL_GRID_WIDTH, rsTransaction.RecordCount
        
    End If
    
    ClearRS rsTransaction
End Sub


Private Sub cmdUpdate_Click()

End Sub

Private Sub Form_Activate()
    txtSearch.SetFocus
End Sub

Private Sub Form_Load()
    Dim ictr%
    
    For ictr% = 0 To 5
        msfTransactionDetails.ColWidth(ictr) = lblCaption(ictr).Width '+ 10
        
    Next
    
    msfTransactionDetails.ColAlignment(0) = 7
    msfTransactionDetails.ColAlignment(1) = 1
    msfTransactionDetails.ColAlignment(2) = 7
    msfTransactionDetails.ColAlignment(3) = 7
    msfTransactionDetails.ColAlignment(4) = 7
    msfTransactionDetails.ColAlignment(5) = 7
    
    
    
    If Not g_CurrentUser.ReturnItem Then
        cmdReplaceItem.Enabled = False
        cmdReturnItem.Enabled = False
    End If
    
    CenterFrm Me
End Sub

Private Sub ResetForm()
    lblSubTotal = 0#
    
    
    txtOrderDate = ""
    
    
    msfTransactionDetails.Clear
    msfTransactionDetails.Width = NORMAL_GRID_WIDTH
    msfTransactionDetails.Rows = 1
    
    

End Sub


Private Sub txtSearch_Change()
    ResetForm
End Sub
