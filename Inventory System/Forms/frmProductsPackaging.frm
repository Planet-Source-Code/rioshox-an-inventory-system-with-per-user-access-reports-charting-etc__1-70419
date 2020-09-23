VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProductsPackaging 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   4920
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
      Left            =   1500
      TabIndex        =   0
      Top             =   1020
      Width           =   1500
   End
   Begin VB.CommandButton cmdShowDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Packaging Type Details"
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
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4490
      Width           =   2670
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
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5025
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
      Left            =   165
      TabIndex        =   3
      Top             =   1020
      Value           =   -1  'True
      Width           =   1500
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
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   570
      Width           =   960
   End
   Begin VB.TextBox txtSearch 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   540
      Width           =   2700
   End
   Begin MSFlexGridLib.MSFlexGrid msfSearchResult 
      Height          =   2505
      Left            =   195
      TabIndex        =   6
      Top             =   1890
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
      Index           =   1
      Left            =   2160
      TabIndex        =   10
      Top             =   1590
      Width           =   2565
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Packaging Code"
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
      Left            =   210
      TabIndex        =   9
      Top             =   1590
      Width           =   1950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   8
      Top             =   315
      Width           =   2250
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Packaging"
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
      Left            =   75
      TabIndex        =   7
      Top             =   45
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3525
      Left            =   105
      Top             =   1395
      Width           =   4710
   End
End
Attribute VB_Name = "frmProductsPackaging"
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
    rsActiveTable.Open "SELECT * FROM qrySearchPackaging WHERE PackagingTypeID = " & EnNone(msfSearchResult.TextMatrix(msfSearchResult.Row, 0)), connRFM, adOpenStatic, adLockReadOnly
    
    
    Hide
    With frmProductsPackagingDE
        .Tag = TO_EDIT
        
        .lblTitle = "Packaging Type Details"
        
        .txtPackagingTypeID = rsActiveTable("PackagingTypeID")
        .txtPackagingType = rsActiveTable("PackagingType")
        .txtPackagingTypeID.Locked = True
        
        ClearRS rsActiveTable
        
        .Show 1
    End With
    Show


End Sub

Private Sub RefreshGrid(strsearch As String)
    Dim rsActiveTable As New ADODB.Recordset
    
    rsActiveTable.Open _
        "SELECT * FROM qrySearchPackaging WHERE PackagingType LIKE " & strsearch, connRFM, adOpenStatic, adLockReadOnly
    
    msfSearchResult.Rows = 0
    If Not NoRecord(rsActiveTable) Then
        rsActiveTable.MoveFirst
        While Not rsActiveTable.EOF
            msfSearchResult.AddItem Format(rsActiveTable("PackagingTypeID"), "000") & vbTab & rsActiveTable("PackagingType")
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






