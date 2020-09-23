VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEmployees 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
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
   LockControls    =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSetSystemAccess 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Set System Access"
      Height          =   360
      Left            =   5115
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4350
      Width           =   2430
   End
   Begin VB.OptionButton optSearch 
      BackColor       =   &H00C18B59&
      Caption         =   "Contains"
      Height          =   210
      Index           =   1
      Left            =   1575
      TabIndex        =   0
      Top             =   990
      Width           =   1500
   End
   Begin VB.CommandButton cmdShowDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Employee Details"
      Height          =   360
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4365
      Width           =   2430
   End
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   330
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4860
      Width           =   960
   End
   Begin VB.OptionButton optSearch 
      BackColor       =   &H00C18B59&
      Caption         =   "Begins with"
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   990
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdSeacrh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      Height          =   330
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   660
      Width           =   960
   End
   Begin VB.TextBox txtSearch 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   630
      Width           =   2700
   End
   Begin MSFlexGridLib.MSFlexGrid msfSearchResult 
      Height          =   2490
      Left            =   255
      TabIndex        =   6
      Top             =   1740
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   4392
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14663855
      GridColor       =   0
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Position"
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
      Left            =   5390
      TabIndex        =   11
      Top             =   1440
      Width           =   2160
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Employee Name"
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
      Left            =   1910
      TabIndex        =   10
      Top             =   1440
      Width           =   3480
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C18B59&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Employee ID"
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
      Left            =   270
      TabIndex        =   9
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   405
      Width           =   2250
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Employees"
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
      Left            =   165
      TabIndex        =   7
      Top             =   120
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3525
      Left            =   180
      Top             =   1245
      Width           =   7560
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmdSetSystemAccess_Click()
    Dim rsSystemAccess As New ADODB.Recordset
    Dim ictr%
    If msfSearchResult.TextMatrix(msfSearchResult.Row, 0) = "" Then Exit Sub
    
    
    If Not g_CurrentUser.UserAccess Then
        If msfSearchResult.TextMatrix(msfSearchResult.Row, 0) <> g_CurrentUser.EmployeeID Then
            MsgBox "Your not allowed to modify passwords of ther users...", vbCritical
            Exit Sub
        End If
        
    End If
    
    
     
    rsSystemAccess.Open "SELECT * FROM tblSystemUsers WHERE EmployeeID = " & msfSearchResult.TextMatrix(msfSearchResult.Row, 0), connRFM, adOpenStatic
    
    Hide
    With frmSystemAccess
        .txtUserName.Tag = msfSearchResult.TextMatrix(msfSearchResult.Row, 0)
        
        If NoRecord(rsSystemAccess) Then
            .Tag = TO_ADD
            .Show 1
        Else
            .Tag = TO_EDIT
            .txtUserName = Decrypt(rsSystemAccess("UserName"), -1)
            .txtPassword = Decrypt(rsSystemAccess("Password"), -1)
            .txtCondirmPassword = Decrypt(rsSystemAccess("Password"), -1)
            
            For ictr% = 0 To TO_ACCESS_COUNT - 1
                If rsSystemAccess(.chkFormBasedAccess(ictr).Tag) Then
                    .chkFormBasedAccess(ictr).Value = Checked
                Else
                    .chkFormBasedAccess(ictr).Value = Unchecked
                End If
                
                If Not g_CurrentUser.UserAccess Then .chkFormBasedAccess(ictr).Enabled = False
            Next
            
            .Show 1
        End If
    End With
    Show
End Sub

Private Sub cmdShowDetails_Click()
    Dim rsActiveTable As ADODB.Recordset
    
    If msfSearchResult.TextMatrix(msfSearchResult.Row, 0) = "" Then Exit Sub
    Set rsActiveTable = New ADODB.Recordset
    rsActiveTable.Open "SELECT * FROM tblEmployees WHERE EmployeeID = " & msfSearchResult.TextMatrix(msfSearchResult.Row, 0), connRFM
    
    
    
    Hide
    With frmEmployeesDE
        .lblTitle = "Employee Details"
        .txtEmployeeID = Format(rsActiveTable("EmployeeID"), "0000")
        .txtLastName = rsActiveTable("LastName") & ""
        .txtFirstName = rsActiveTable("FirstName") & ""
        .txtMiddleName = rsActiveTable("MiddleName") & ""
        .txtAddress = rsActiveTable("Address") & ""
        .txtCity = rsActiveTable("City") & ""
        .txtZIP = rsActiveTable("ZIP") & ""
        .txtBirthdate = rsActiveTable("Birthdate") & ""
        .txtHireDate = rsActiveTable("HireDate") & ""
        .txtDepartment = rsActiveTable("Department") & ""
        .txtPosition = rsActiveTable("Position") & ""
        .txtContactNo = rsActiveTable("ContactNo") & ""
        .txtGender = rsActiveTable("Gender") & ""
        .txtEmailAddress = rsActiveTable("EmailAddress") & ""
        
        If Trim(rsActiveTable("PicturePath") & "") <> "" Then _
        .imgEmployee.Picture = LoadPicture(App.Path & "\Employee Images\" & rsActiveTable("PicturePath"))
        
        ClearRS rsActiveTable
        .Show 1
    End With
    Show
End Sub


Private Sub Form_Load()
    Dim ictr%
    
    CenterFrm Me
    
    For ictr = 0 To 2
        msfSearchResult.ColWidth(ictr) = Me.lblHeader(ictr).Width
    Next
    
    If Not g_CurrentUser.UserAccess Then Me.cmdSetSystemAccess.Caption = "Change Password"
End Sub

Private Sub RefreshGrid(strsearch As String)
    Dim rsActiveTable As New ADODB.Recordset
    
    rsActiveTable.Open _
        "SELECT * FROM qrySearchEmployee WHERE EmployeeName  LIKE " & strsearch, connRFM, adOpenStatic, adLockReadOnly
    
    msfSearchResult.Rows = 0
    If Not NoRecord(rsActiveTable) Then
        rsActiveTable.MoveFirst
        While Not rsActiveTable.EOF
            msfSearchResult.AddItem Format(rsActiveTable("EmployeeID"), "00000") & vbTab & rsActiveTable("EmployeeName") & vbTab & rsActiveTable("Position")
            rsActiveTable.MoveNext
        Wend
    End If
    
    msfSearchResult.AddItem ""
    
    'ResizeGrid msfSearchResult, NORMAL_GRID_WIDTH, rsActiveTable.RecordCount
    
    ClearRS rsActiveTable
End Sub

    



 
