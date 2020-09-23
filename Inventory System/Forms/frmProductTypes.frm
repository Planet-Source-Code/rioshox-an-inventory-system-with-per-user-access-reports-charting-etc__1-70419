VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProductTypes 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
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
      Left            =   1530
      TabIndex        =   3
      Top             =   1050
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
      Left            =   135
      TabIndex        =   6
      Top             =   570
      Width           =   2700
   End
   Begin VB.CommandButton cmdSearch 
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
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
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
      Left            =   195
      TabIndex        =   4
      Top             =   1050
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
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5055
      Width           =   960
   End
   Begin VB.CommandButton cmdShowDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Product Type Details"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4485
      Width           =   2430
   End
   Begin MSFlexGridLib.MSFlexGrid msfSearchResult 
      Height          =   2505
      Left            =   225
      TabIndex        =   2
      Top             =   1920
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4419
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14663855
      GridColor       =   0
      ScrollBars      =   0
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
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Types "
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
      TabIndex        =   10
      Top             =   75
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
      Left            =   135
      TabIndex        =   9
      Top             =   345
      Width           =   2250
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product Type Code"
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
      Left            =   235
      TabIndex        =   8
      Top             =   1620
      Width           =   1950
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product Type Name"
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
      Left            =   2190
      TabIndex        =   7
      Top             =   1620
      Width           =   2565
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3525
      Left            =   135
      Top             =   1425
      Width           =   4710
   End
End
Attribute VB_Name = "frmProductTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    rsActiveTable.Open "SELECT * FROM qrySearchProductTypes WHERE ProductTypeID = " & EnNone(msfSearchResult.TextMatrix(msfSearchResult.Row, 0)), connRFM, adOpenStatic, adLockReadOnly
    
    
    Hide
    With frmProductTypesDE
        .Tag = TO_EDIT
        
        .lblTitle = "Product Type Details"
        
        .txtProductTypeID = rsActiveTable("ProductTypeID")
        .txtProductType = rsActiveTable("ProductType")
        
        ClearRS rsActiveTable
        .Show 1
    End With
    Show


End Sub

Private Sub RefreshGrid(strsearch As String)
    Dim rsActiveTable As New ADODB.Recordset
    
    rsActiveTable.Open _
        "SELECT * FROM qrySearchProductTypes WHERE ProductType LIKE " & strsearch, connRFM, adOpenStatic, adLockReadOnly
    
    msfSearchResult.Rows = 0
    If Not NoRecord(rsActiveTable) Then
        rsActiveTable.MoveFirst
        While Not rsActiveTable.EOF
            msfSearchResult.AddItem Format(rsActiveTable("ProductTypeID"), "000") & vbTab & rsActiveTable("ProductType")
            rsActiveTable.MoveNext
        Wend
    End If
    
    msfSearchResult.AddItem ""
    
        
    ClearRS rsActiveTable
End Sub

    










Private Sub Form_Load()
    Dim ictr%
    
    For ictr = 0 To 1
        msfSearchResult.ColWidth(ictr) = Me.lblHeader(ictr).Width
    Next
    
    CenterFrm Me
End Sub




