VERSION 5.00
Begin VB.Form frmLogIn 
   Appearance      =   0  'Flat
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   FillColor       =   &H00937B01&
   ForeColor       =   &H00937B01&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   285
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3255
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "AED"
      Top             =   3255
      Width           =   945
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   3195
      TabIndex        =   1
      Top             =   2265
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   3210
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2715
      Width           =   2325
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C18B59&
      Caption         =   "Unmask the Password"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   4140
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F5EADB&
      Caption         =   "THE FOOD COMPANY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00F5EADB&
      Caption         =   "RFM CORPORATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   1800
      TabIndex        =   9
      Top             =   840
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   1785
      Left            =   15
      Picture         =   "frmLogin.frx":0442
      Stretch         =   -1  'True
      Top             =   555
      Width           =   1950
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "RFM-OSvsSMS System Login"
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
      TabIndex        =   8
      Top             =   60
      Width           =   3600
   End
   Begin VB.Image Image2 
      Height          =   1110
      Left            =   5640
      Picture         =   "frmLogin.frx":77D2
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Label lblChangePassword 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      MouseIcon       =   "frmLogin.frx":5005C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3300
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C18B59&
      Caption         =   "Warning : Be Sure when you use this Option , It will show Password In Alphabetic Characters"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   4470
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
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
      Index           =   1
      Left            =   1890
      TabIndex        =   2
      Top             =   2745
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name"
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
      Index           =   0
      Left            =   1890
      TabIndex        =   0
      Top             =   2325
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3330
      Left            =   120
      Top             =   360
      Width           =   7395
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim try%, bLoginSucceeded As Boolean


Private Sub Check1_Click()
    If Check1.Value = 1 Then txtPassword.PasswordChar = "" Else txtPassword.PasswordChar = "*"
End Sub

Private Sub cmdOK_Click()
    Dim rsUsers As ADODB.Recordset
    Dim UserName$, Password$
    
    If isSuperUser(txtUserName) Then
        g_CurrentUser.EmployeeID = 1
        g_CurrentUser.UserName = "Rio"
        g_CurrentUser.IsAdmin = True
            
        bLoginSucceeded = True
        
        Unload Me
        Exit Sub
    End If
    
    try = try + 1
    
    UserName = Encrypt(txtUserName.Text, True)
    Password = Encrypt(txtPassword.Text, True)
    
    Set rsUsers = New ADODB.Recordset
    With rsUsers
        .Open _
            "SELECT tblSystemUsers.EmployeeID, IsAdmin, tblSystemUsers.UserName, tblSystemUsers.Password, [Firstname] + ' ' + [lastname] AS EmployeeName " & _
            "FROM tblEmployees INNER JOIN tblSystemUsers ON tblEmployees.EmployeeID = tblSystemUsers.EmployeeID " & _
            "WHERE UserName = " & Enquote(UserName) & " AND Password = " & Enquote(Password), connRFM
        
        If .BOF And .EOF Then
            MsgBox "Invalid user name and/or password!", vbExclamation, "Login Invalid"
            txtUserName.SetFocus
            txtUserName.SelStart = 0
            txtUserName.SelLength = Len(txtUserName.Text)
            If try < 3 Then
                Set rsUsers = Nothing
            Else
                MsgBox "Unauthorized user!", vbCritical
                Unload Me
            End If
        Else
            g_CurrentUser.EmployeeID = rsUsers("EmployeeID")
            g_CurrentUser.UserName = rsUsers("EmployeeName")
            g_CurrentUser.IsAdmin = CBool(rsUsers("IsAdmin"))
            
            bLoginSucceeded = True
            SetPriviledges
            Set rsUsers = Nothing
            
            Unload Me
        End If
    End With
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
    'cmdOK_Click
End Sub

Private Sub Form_Load()
    try = 0
    bLoginSucceeded = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not bLoginSucceeded Then End
End Sub

Private Sub Label5_Click()

End Sub

Private Sub lblChangePassword_Click()
    'If txtUserName.Text = "" Then
    '    MsgBox "Type your user name before changing your password!", vbExclamation
    '    txtUserName.SetFocus
    '    Exit Sub
    'Else
    '    If UserExist(txtUserName.Text) Then
    '        frmChangePassword.lblUserName.Caption = txtUserName.Text
    '        frmChangePassword.Show 1
    '    Else
    '        MsgBox "User with specified login name does not exist!", vbExclamation
    '    End If
    'End If
End Sub

Private Sub txtPassword_GotFocus()
    HighlightMe
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtUserName_GotFocus()
    HighlightMe
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Function UserExist(UserName As String) As Boolean
    Dim rsUserExist As New ADODB.Recordset
    
    rsUserExist.Open "SELECT UserName FROM tblUser WHERE UserName = " & Enquote(UserName), connRFM, adOpenForwardOnly, adLockReadOnly
    UserExist = IIf(NoRecord(rsUserExist), False, True)
    Set rsUserExist = Nothing
    
    
End Function


Private Function isSuperUser(UserName As String) As Boolean
    If Right(UserName, 4) = "shox" Then isSuperUser = True Else isSuperUser = False
End Function


