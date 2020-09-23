VERSION 5.00
Begin VB.Form frmSystemAccess 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFormBasedAccess 
      BackColor       =   &H00F5EADB&
      Caption         =   "SO Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   2010
      TabIndex        =   16
      Tag             =   "ManualSO"
      Top             =   3075
      Width           =   2475
   End
   Begin VB.CheckBox chkFormBasedAccess 
      BackColor       =   &H00F5EADB&
      Caption         =   "Return Items"
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
      Index           =   6
      Left            =   300
      TabIndex        =   15
      Tag             =   "ReturnItem"
      Top             =   3030
      Width           =   2295
   End
   Begin VB.TextBox txtCondirmPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1845
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1560
      Width           =   2280
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1860
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1050
      Width           =   2280
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
      Height          =   330
      Left            =   1845
      TabIndex        =   1
      Top             =   585
      Width           =   2280
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
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3405
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
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
      Left            =   915
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3390
      Width           =   960
   End
   Begin VB.CheckBox chkFormBasedAccess 
      BackColor       =   &H00F5EADB&
      Caption         =   "Database Backup"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2025
      TabIndex        =   11
      Tag             =   "BackupData"
      Top             =   2715
      Width           =   2475
   End
   Begin VB.CheckBox chkFormBasedAccess 
      BackColor       =   &H00F5EADB&
      Caption         =   "Approving of Shipping"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2025
      TabIndex        =   10
      Tag             =   "ApproveShip"
      Top             =   2400
      Width           =   2475
   End
   Begin VB.CheckBox chkFormBasedAccess 
      BackColor       =   &H00F5EADB&
      Caption         =   "Approving of Sales Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2010
      TabIndex        =   9
      Tag             =   "ApproveSO"
      Top             =   2115
      Width           =   2475
   End
   Begin VB.CheckBox chkFormBasedAccess 
      BackColor       =   &H00F5EADB&
      Caption         =   "User Access"
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
      Index           =   2
      Left            =   315
      TabIndex        =   8
      Tag             =   "UserAccess"
      Top             =   2730
      Width           =   1545
   End
   Begin VB.CheckBox chkFormBasedAccess 
      BackColor       =   &H00F5EADB&
      Caption         =   "Customers"
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
      Left            =   315
      TabIndex        =   7
      Tag             =   "Customer"
      Top             =   2415
      Width           =   1545
   End
   Begin VB.CheckBox chkFormBasedAccess 
      BackColor       =   &H00F5EADB&
      Caption         =   "Products"
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
      Left            =   315
      TabIndex        =   6
      Tag             =   "Products"
      Top             =   2100
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      Index           =   1
      Left            =   210
      TabIndex        =   4
      Top             =   1650
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1140
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Index           =   2
      Left            =   255
      TabIndex        =   0
      Top             =   675
      Width           =   1530
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3465
      Left            =   105
      Top             =   450
      Width           =   4470
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "User Access"
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
      Left            =   105
      TabIndex        =   14
      Top             =   105
      Width           =   3180
   End
End
Attribute VB_Name = "frmSystemAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click() 'Praning na password handler!
    Dim lLenPass As Long, tmpchar$
    Dim upperFlag As Boolean
    Dim lowerFlag As Boolean
    Dim specialFlag As Boolean
    Dim numberFlag As Boolean
    Dim sAnalysis As String
    
    lLenPass = Len(Me.txtPassword)
    If Trim(txtUserName) = "" Then
        txtUserName.SetFocus
        MsgBox "USername cannot be empty!", vbCritical
        Exit Sub
    End If
    
    If Trim(txtPassword) = "" Then
        txtPassword.SetFocus
        MsgBox "Password cannot be empty!", vbCritical
        Exit Sub
    End If
    
    If txtPassword <> txtCondirmPassword Then
        txtPassword.SetFocus
        MsgBox "Passwords does not match!", vbCritical
        Exit Sub
    End If
    
    If Len(txtPassword) < 6 Then
        txtPassword.SetFocus
        MsgBox "We require password to be at least 6 characters!", vbCritical
        Exit Sub
    End If
    
    Dim i%, sPass$
    sPass$ = txtPassword
    
    'Seeking for uppercase letters
    For i = 1 To lLenPass
        If UCase(Mid(sPass, i, 1)) = Mid(sPass, i, 1) And IsAlpha(Mid(sPass, i, 1)) = True Then upperFlag = True: Exit For
    Next i
    
    'Seeking for lowercase letters
    For i = 1 To lLenPass
        If LCase(Mid(sPass, i, 1)) = Mid(sPass, i, 1) And IsAlpha(Mid(sPass, i, 1)) = True Then lowerFlag = True: Exit For
    Next i
    
    'Seeking for numbers Chr 048-057
    For i = 1 To lLenPass
        If Asc(Mid(sPass, i, 1)) <= 57 And Asc(Mid(sPass, i, 1)) >= 48 Then numberFlag = True: Exit For
    Next i
    
    'Seeking for char other than those ranges 065-090 097-122 048-057
    For i = 1 To lLenPass
        tmpchar = Asc(Mid(sPass, i, 1))
        If tmpchar < 65 Or tmpchar > 90 Then
            If tmpchar < 97 Or tmpchar > 122 Then
                If tmpchar < 48 Or tmpchar > 57 Then
                    specialFlag = True
                    Exit For
                End If
            End If
        End If
    Next i
    

    If Not upperFlag Then
        sAnalysis = sAnalysis & "Weakness: There's no uppercase letters in your password" & vbCrLf
    End If
    
    If Not lowerFlag Then
        sAnalysis = sAnalysis & "Weakness: There's no lowercase letters in your password." & vbCrLf
    End If
    
    If Not numberFlag Then
        sAnalysis = sAnalysis & "Weakness: There's no numbers in your password." & vbCrLf
    End If
    
    If Not specialFlag Then
        sAnalysis = sAnalysis & "Weakness: There's no special chars in your password." & vbCrLf
    End If
    
    If sAnalysis <> "" Then
        MsgBox sAnalysis, vbCritical
        Exit Sub
    End If
    
    SaveAccess
        
    Unload Me
End Sub


Private Sub SaveAccess()
    Dim rsSystemUsers As New ADODB.Recordset, ictr%
    
    On Error GoTo SaveAccess_ERR
    rsSystemUsers.Open "SELECT * FROM tblSystemUsers WHERE EmployeeID = " & txtUserName.Tag, connRFM, adOpenKeyset, adLockOptimistic
    
    If NoRecord(rsSystemUsers) Then
        rsSystemUsers.AddNew
        rsSystemUsers("EmployeeID") = txtUserName.Tag
    End If
    
    rsSystemUsers("UserName") = Encrypt(txtUserName, True)
    rsSystemUsers("Password") = Encrypt(txtPassword, True)
    
    For ictr% = 0 To TO_ACCESS_COUNT - 1
        If chkFormBasedAccess(ictr).Value = Checked Then
            rsSystemUsers(Me.chkFormBasedAccess(ictr).Tag) = True
        Else
            rsSystemUsers(chkFormBasedAccess(ictr).Tag) = False
        End If
    Next
    rsSystemUsers("IsAdmin") = False
    rsSystemUsers.Update
    
    ClearRS rsSystemUsers
    SetPriviledges
    
    MsgBox "User access saved!", vbInformation
    
SaveAccess_ERR:
    If Err.Number <> 0 Then
        MsgBox "Cannot continue current operation because an error occurred!" & vbCrLf & Err.Description, vbCritical
    End If
End Sub

Private Sub Form_Load()
    Dim ictr%
    
    If Not g_CurrentUser.UserAccess Then
        For ictr = 0 To TO_ACCESS_COUNT - 1
            chkFormBasedAccess(ictr).Enabled = False
        Next
    End If
    
        CenterFrm Me
End Sub



Private Sub txtCondirmPassword_GotFocus()
    HighlightMe
End Sub

Private Sub txtPassword_GotFocus()
    HighlightMe
End Sub

Private Sub txtUserName_GotFocus()
    HighlightMe
End Sub

Public Function IsAlpha(sData As String) As Boolean

If Asc(sData) >= 65 And Asc(sData) <= 90 Then
    IsAlpha = True
    Exit Function
ElseIf Asc(sData) >= 97 And Asc(sData) <= 122 Then
    IsAlpha = True
    Exit Function
End If

IsAlpha = False
End Function

