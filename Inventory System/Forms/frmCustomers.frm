VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCustomers 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optSearch 
      BackColor       =   &H00C18B59&
      Caption         =   "Contains"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   1065
      Width           =   1500
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   285
      TabIndex        =   6
      Top             =   705
      Width           =   2700
   End
   Begin VB.CommandButton cmdSeacrh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   735
      Width           =   960
   End
   Begin VB.OptionButton optSearch 
      BackColor       =   &H00C18B59&
      Caption         =   "Begins with"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   345
      TabIndex        =   4
      Top             =   1065
      Value           =   -1  'True
      Width           =   1500
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   960
   End
   Begin VB.CommandButton cmdShowDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Customer Details"
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
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4660
      Width           =   2430
   End
   Begin MSFlexGridLib.MSFlexGrid msfSearchResult 
      Height          =   2670
      Left            =   375
      TabIndex        =   2
      Top             =   1920
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   4710
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14663855
      GridColor       =   0
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Customers"
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
      Left            =   270
      TabIndex        =   12
      Top             =   195
      Width           =   2250
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   285
      TabIndex        =   11
      Top             =   480
      Width           =   2250
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customer Code"
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
      Left            =   375
      TabIndex        =   10
      Top             =   1635
      Width           =   1635
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Company Name"
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
      Left            =   2010
      TabIndex        =   9
      Top             =   1635
      Width           =   3480
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contact Name"
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
      Left            =   5490
      TabIndex        =   8
      Top             =   1635
      Width           =   2160
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contact Number"
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
      Left            =   7650
      TabIndex        =   7
      Top             =   1635
      Width           =   1830
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3690
      Left            =   285
      Top             =   1440
      Width           =   9345
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NORMAL_GRID_WIDTH = 9120
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSeacrh_Click()
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
    rsActiveTable.Open "SELECT * FROM tblCustomers WHERE CustomerID = " & Enquote(msfSearchResult.TextMatrix(msfSearchResult.Row, 0)), connRFM
    
    
    
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
        .txtSMSNumber = rsActiveTable("SMSNumber") & ""
        .txtCreditLimit = Format(rsActiveTable("CreditLimit"), "STANDARD")
        .txtMinimumPurchase = Format(rsActiveTable("MinimumPurchase"), "STANDARD")
        
        ClearRS rsActiveTable
        
        .Show 1
    End With
    Show
End Sub

Private Sub Form_Load()
    Dim ictr%
    
    For ictr = 0 To 3
        msfSearchResult.ColWidth(ictr) = Me.lblHeader(ictr).Width
    Next
    CenterFrm Me
End Sub

Private Sub RefreshGrid(strsearch As String)
    Dim rsActiveTable As New ADODB.Recordset
    
    rsActiveTable.Open _
        "SELECT * FROM qrySearchCustomer WHERE CustomerID LIKE " & strsearch, connRFM, adOpenStatic, adLockReadOnly
    
    msfSearchResult.Rows = 0
    If Not NoRecord(rsActiveTable) Then
        rsActiveTable.MoveFirst
        While Not rsActiveTable.EOF
            msfSearchResult.AddItem rsActiveTable("CustomerID") & vbTab & rsActiveTable("CompanyName") & vbTab & rsActiveTable("ContactName") & "" & vbTab & rsActiveTable("Phone")
            rsActiveTable.MoveNext
        Wend
    End If
    
    msfSearchResult.AddItem ""
    
    ResizeGrid msfSearchResult, NORMAL_GRID_WIDTH, rsActiveTable.RecordCount
    
    ClearRS rsActiveTable
End Sub

    



 


