VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProducts 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdShowDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Product Details"
      Height          =   360
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   2430
   End
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   330
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5325
      Width           =   960
   End
   Begin MSFlexGridLib.MSFlexGrid msfSearchResult 
      Height          =   2670
      Left            =   300
      TabIndex        =   11
      Top             =   2055
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   4710
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14663855
      GridColor       =   0
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.OptionButton optSeacrh 
      BackColor       =   &H00C18B59&
      Caption         =   "Contains"
      Height          =   210
      Index           =   1
      Left            =   1605
      TabIndex        =   5
      Top             =   1185
      Width           =   1500
   End
   Begin VB.OptionButton optSearch 
      BackColor       =   &H00C18B59&
      Caption         =   "Begins with"
      Height          =   210
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   1185
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   330
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   735
      Width           =   960
   End
   Begin VB.TextBox txtSearch 
      Height          =   330
      Left            =   210
      TabIndex        =   2
      Top             =   705
      Width           =   2700
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Units On Order"
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
      Left            =   7770
      TabIndex        =   10
      Top             =   1755
      Width           =   1635
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Units In Stock"
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
      Left            =   6135
      TabIndex        =   9
      Top             =   1755
      Width           =   1635
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Packaging"
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
      Left            =   4500
      TabIndex        =   8
      Top             =   1755
      Width           =   1635
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product Name"
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
      Left            =   1935
      TabIndex        =   7
      Top             =   1755
      Width           =   2565
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
      Index           =   0
      Left            =   300
      TabIndex        =   6
      Top             =   1755
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   480
      Width           =   2250
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Products"
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
      TabIndex        =   0
      Top             =   210
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3705
      Left            =   210
      Top             =   1560
      Width           =   9315
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NORMAL_GRID_WIDTH As Long = 9105
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    If Trim(txtSearch.Text) = "" Then
        txtSearch.SetFocus
        Exit Sub
    End If
     
    msfSearchResult.Clear
    msfSearchResult.Rows = 0
    
    If optSearch(0).Value = True Then
        RefreshGrid Enquote(txtSearch & "%")
        
    Else
        RefreshGrid Enquote("%" & txtSearch & "%")
    End If
End Sub

Private Sub cmdShowDetails_Click()
    Dim rsActiveTable As ADODB.Recordset
    
    If msfSearchResult.TextMatrix(msfSearchResult.Row, 0) = "" Then Exit Sub
    Set rsActiveTable = New ADODB.Recordset
    rsActiveTable.Open "SELECT * FROM qrySearchProduct WHERE ProductID = " & EnNone(msfSearchResult.TextMatrix(msfSearchResult.Row, 0)), connRFM, adOpenStatic, adLockReadOnly
    
    
    Hide
    With frmProductsDE
        .Tag = TO_EDIT
        
        .lblTitle = "Product Details"
        
        .txtProductID = rsActiveTable("ProductID")
        .txtProductName = rsActiveTable("Productname")
        .txtQuantityPerUnit = rsActiveTable("QuantityPerUnit")
        .cboProductTypes = rsActiveTable("ProductType")
        .cboPackagingType = rsActiveTable("PackagingType")
        
        .txtUnitPrice = Format(rsActiveTable("UnitPrice"), "STANDARD")
        .txtUnitsInStock = rsActiveTable("UnitsInStock")
        .txtUnitsOnOrder = rsActiveTable("UnitsOnOrder")
        .txtReorderLevel = rsActiveTable("ReorderLevel")
        
        
        ClearRS rsActiveTable
        
        .Show 1
    End With
    Show


End Sub

Private Sub RefreshGrid(strsearch As String)
    Dim rsActiveTable As New ADODB.Recordset
    
    rsActiveTable.Open _
        "SELECT * FROM qrySearchproduct WHERE ProductName LIKE " & strsearch, connRFM, adOpenStatic, adLockReadOnly
    
    msfSearchResult.Rows = 0
    If Not NoRecord(rsActiveTable) Then
        rsActiveTable.MoveFirst
        While Not rsActiveTable.EOF
            msfSearchResult.AddItem Format(rsActiveTable("ProductID"), "00000") & vbTab & rsActiveTable("ProductName") & vbTab & rsActiveTable("PackagingType") & "" & vbTab & rsActiveTable("UnitsInStock") & vbTab & rsActiveTable("UnitsOnOrder")
            rsActiveTable.MoveNext
        Wend
    End If
    
    msfSearchResult.AddItem ""
    
    ResizeGrid msfSearchResult, NORMAL_GRID_WIDTH, rsActiveTable.RecordCount
    ClearRS rsActiveTable
End Sub

    

Private Sub Form_Load()
    Dim ictr%
    CenterFrm Me
    For ictr = 0 To 4
        msfSearchResult.ColWidth(ictr) = Me.lblHeader(ictr).Width
    Next
    
    
End Sub


