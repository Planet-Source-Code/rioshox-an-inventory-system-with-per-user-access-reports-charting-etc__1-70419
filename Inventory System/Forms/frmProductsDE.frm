VERSION 5.00
Begin VB.Form frmProductsDE 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
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
   LockControls    =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPackagingType 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2220
      Width           =   2280
   End
   Begin VB.ComboBox cboPackagingTypeID 
      Height          =   315
      Left            =   3930
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   6405
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.ComboBox cboproductTypeID 
      Height          =   315
      Left            =   4050
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5895
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.ComboBox cboProductTypes 
      Height          =   315
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2205
      Width           =   2280
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print"
      Height          =   330
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4305
      Width           =   960
   End
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   330
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4305
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   330
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4305
      Width           =   960
   End
   Begin VB.TextBox txtReorderLevel 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   3225
      TabIndex        =   17
      Top             =   3870
      Width           =   1260
   End
   Begin VB.TextBox txtUnitsOnOrder 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1710
      TabIndex        =   15
      Top             =   3870
      Width           =   1260
   End
   Begin VB.TextBox txtUnitsInStock 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   240
      TabIndex        =   13
      Top             =   3870
      Width           =   1260
   End
   Begin VB.TextBox txtUnitPrice 
      Height          =   330
      Left            =   2430
      TabIndex        =   11
      Top             =   2970
      Width           =   2025
   End
   Begin VB.TextBox txtQuantityPerUnit 
      Height          =   330
      Left            =   285
      TabIndex        =   9
      Top             =   2970
      Width           =   2025
   End
   Begin VB.TextBox txtProductName 
      Height          =   330
      Left            =   285
      TabIndex        =   3
      Top             =   1380
      Width           =   4740
   End
   Begin VB.TextBox txtProductID 
      Height          =   330
      Left            =   285
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   690
      Width           =   2460
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packaging Type"
      Height          =   300
      Index           =   8
      Left            =   2880
      TabIndex        =   6
      Top             =   1980
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reorder Level"
      Height          =   300
      Index           =   7
      Left            =   3225
      TabIndex        =   16
      Top             =   3615
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Units On Order"
      Height          =   300
      Index           =   6
      Left            =   1710
      TabIndex        =   14
      Top             =   3615
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Units In Stock"
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   3615
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   300
      Index           =   5
      Left            =   285
      TabIndex        =   0
      Top             =   420
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
      Height          =   300
      Index           =   4
      Left            =   2430
      TabIndex        =   10
      Top             =   2730
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Per unit"
      Height          =   300
      Index           =   3
      Left            =   285
      TabIndex        =   8
      Top             =   2730
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Type"
      Height          =   300
      Index           =   2
      Left            =   285
      TabIndex        =   4
      Top             =   1980
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   300
      Index           =   1
      Left            =   285
      TabIndex        =   2
      Top             =   1140
      Width           =   2190
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Product"
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
      Left            =   195
      TabIndex        =   20
      Top             =   105
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3015
      Left            =   105
      Top             =   390
      Width           =   5310
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1230
      Left            =   105
      Top             =   3480
      Width           =   5310
   End
End
Attribute VB_Name = "frmProductsDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboPackagingType_Click()
    cboPackagingTypeID.ListIndex = cboPackagingType.ListIndex
End Sub

Private Sub cboProductTypes_Click()
    cboproductTypeID.ListIndex = cboProductTypes.ListIndex
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strSQL$, msg$, IsNewProduct As Boolean
    
    
    If Trim(txtProductName.Text) = "" Then
        MsgBox "Product name cannot be empty..", vbExclamation
        txtProductName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtQuantityPerUnit.Text) = "" Then
        MsgBox "Qty/unit cannot be empty..", vbExclamation
        txtQuantityPerUnit.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtUnitPrice.Text) Then
        MsgBox "Invalid unit price...", vbExclamation
        txtUnitPrice.SetFocus
        Exit Sub
    End If
    
    
    If Not IsNumeric(txtUnitsInStock.Text) Then
        MsgBox "Invalid unit in stock...", vbExclamation
        txtUnitsInStock.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtUnitsOnOrder.Text) Then
        MsgBox "Invalid unit on order...", vbExclamation
        txtUnitsOnOrder.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtReorderLevel.Text) Then
        MsgBox "Invalid re-order level...", vbExclamation
        txtReorderLevel.SetFocus
        Exit Sub
    End If
     
    
    
    On Error GoTo cmdSave_Click_ERR
    
    Select Case Tag
        Case TO_ADD:
            strSQL$ = _
                "INSERT INTO tblProducts (PackagingTypeID,ProductName,QuantityPerUnit,ProductTypeID,UnitPrice,UnitsInStock,UnitsOnOrder,ReorderLevel) VALUES (" & _
                EnNone(cboPackagingTypeID, True) & _
                Enquote(txtProductName, True) & _
                Enquote(txtQuantityPerUnit, True) & _
                EnNone(cboproductTypeID, True) & _
                EnNone(txtUnitPrice, True) & _
                EnNone(txtUnitsInStock, True) & _
                EnNone(txtUnitsOnOrder, True) & _
                EnNone(txtReorderLevel) & ")"
            msg$ = "New product added."
            IsNewProduct = -1
        
        Case TO_EDIT:
            strSQL$ = _
                "UPDATE tblProducts SET " & _
                "PackagingTypeID= " & EnNone(cboPackagingTypeID, True) & _
                "ProductName = " & Enquote(txtProductName, True) & _
                "QuantityPerUnit = " & Enquote(txtQuantityPerUnit, True) & _
                "ProductTypeID = " & EnNone(cboproductTypeID, True) & _
                "UnitsInStock = " & EnNone(txtUnitsInStock, True) & _
                "UnitPrice = " & EnNone(txtUnitPrice, True) & _
                "UnitsOnOrder = " & EnNone(txtUnitsOnOrder, True) & _
                "ReOrderLevel = " & EnNone(txtReorderLevel) & _
                " WHERE ProductID = " & txtProductID
            msg$ = "Changes saved."
            IsNewProduct = 0
    End Select
    
    
    connRFM.Execute strSQL, adExecuteNoRecords
    MsgBox msg$, vbInformation
    
    If IsNewProduct Then
    End If
    
    If Tag = TO_EDIT Then
            With frmProducts
                .msfSearchResult.TextMatrix(.msfSearchResult.Row, 0) = Format(txtProductID, "000")
                .msfSearchResult.TextMatrix(.msfSearchResult.Row, 1) = txtProductName
                .msfSearchResult.TextMatrix(.msfSearchResult.Row, 2) = cboPackagingType
                .msfSearchResult.TextMatrix(.msfSearchResult.Row, 3) = txtUnitsInStock
                .msfSearchResult.TextMatrix(.msfSearchResult.Row, 4) = txtUnitsOnOrder
            End With
        End If
    
    
    Unload Me

cmdSave_Click_ERR:
    If Err.Number <> 0 Then
        MsgBox "Errors occurred, cannot perform current operation." & vbCrLf & Err.Description, vbCritical
    End If
    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Dim rsProductTypes As New ADODB.Recordset
    Dim rsProductPackaging As New ADODB.Recordset
    
    rsProductTypes.Open "SELECT * FROM tblProductTypes ORDER BY ProductType", connRFM, adOpenStatic, adLockReadOnly
    rsProductPackaging.Open "SELECT * FROM tblPackagingTypes ORDER BY PackagingType", connRFM, adOpenStatic, adLockReadOnly
    
    PopulateCboBox rsProductPackaging, "PackagingTypeID", cboPackagingTypeID
    PopulateCboBox rsProductPackaging, "PackagingType", cboPackagingType
    
    PopulateCboBox rsProductTypes, "ProductTypeID", cboproductTypeID
    PopulateCboBox rsProductTypes, "ProductType", cboProductTypes
    
    
    ClearRS rsProductTypes
    ClearRS rsProductPackaging
    
    If Not g_CurrentUser.Products Then cmdSave.Enabled = False
CenterFrm Me
End Sub

Private Sub txtProductID_GotFocus()
    HighlightMe
End Sub

Private Sub txtProductName_GotFocus()
    HighlightMe
End Sub

Private Sub txtQuantityPerUnit_GotFocus()
    HighlightMe
End Sub

Private Sub txtReorderLevel_GotFocus()
    HighlightMe
End Sub

Private Sub txtUnitPrice_GotFocus()
    HighlightMe
End Sub

Private Sub txtUnitPrice_LostFocus()
    If IsNumeric(txtUnitPrice) Then txtUnitPrice = Format(txtUnitPrice, "STANDARD")
End Sub

Private Sub txtUnitsInStock_GotFocus()
    HighlightMe
End Sub

Private Sub txtUnitsOnOrder_GotFocus()
    HighlightMe
End Sub
