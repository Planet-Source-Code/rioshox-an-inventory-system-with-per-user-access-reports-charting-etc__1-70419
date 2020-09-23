VERSION 5.00
Begin VB.Form frmEmployeesDE 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
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
   ScaleHeight     =   7995
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   330
      Left            =   6255
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7560
      Width           =   960
   End
   Begin VB.TextBox txtContactNo 
      Height          =   330
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   6195
      Width           =   2145
   End
   Begin VB.TextBox txtEmailAddress 
      Height          =   330
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   6900
      Width           =   2145
   End
   Begin VB.TextBox txtZIP 
      Height          =   330
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5490
      Width           =   1005
   End
   Begin VB.TextBox txtCity 
      Height          =   330
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   5490
      Width           =   2310
   End
   Begin VB.TextBox txtAddress 
      Height          =   330
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4830
      Width           =   3435
   End
   Begin VB.TextBox txtHireDate 
      Height          =   330
      Left            =   2625
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3855
      Width           =   2145
   End
   Begin VB.TextBox txtPosition 
      Height          =   330
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3855
      Width           =   2145
   End
   Begin VB.TextBox txtDepartment 
      Height          =   330
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   2145
   End
   Begin VB.TextBox txtBirthdate 
      Height          =   330
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2370
      Width           =   1545
   End
   Begin VB.TextBox txtGender 
      Height          =   330
      Left            =   345
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2250
      Width           =   1545
   End
   Begin VB.TextBox txtMiddleName 
      Height          =   330
      Left            =   4785
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1545
      Width           =   2145
   End
   Begin VB.TextBox txtFirstName 
      Height          =   330
      Left            =   2565
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1545
      Width           =   2145
   End
   Begin VB.TextBox txtLastName 
      Height          =   330
      Left            =   345
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1545
      Width           =   2145
   End
   Begin VB.TextBox txtEmployeeID 
      Height          =   330
      Left            =   345
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Details"
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
      Left            =   240
      TabIndex        =   27
      Top             =   180
      Width           =   2250
   End
   Begin VB.Image imgEmployee 
      BorderStyle     =   1  'Fixed Single
      Height          =   2505
      Left            =   4335
      Top             =   4785
      Width           =   2685
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hire Date"
      Height          =   255
      Index           =   11
      Left            =   2625
      TabIndex        =   14
      Top             =   3585
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   24
      Top             =   6900
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Numbers"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   22
      Top             =   5970
      Width           =   2385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   18
      Top             =   5265
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   16
      Top             =   4575
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      Height          =   255
      Index           =   0
      Left            =   345
      TabIndex        =   0
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   255
      Index           =   5
      Left            =   390
      TabIndex        =   10
      Top             =   2880
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      Height          =   255
      Index           =   4
      Left            =   390
      TabIndex        =   12
      Top             =   3585
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      Height          =   255
      Index           =   2
      Left            =   345
      TabIndex        =   6
      Top             =   1980
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name (Last, First Middle)"
      Height          =   255
      Index           =   1
      Left            =   345
      TabIndex        =   2
      Top             =   1290
      Width           =   3945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ZIP"
      Height          =   255
      Index           =   8
      Left            =   2775
      TabIndex        =   20
      Top             =   5235
      Width           =   810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthdate"
      Height          =   255
      Index           =   3
      Left            =   1980
      TabIndex        =   8
      Top             =   1980
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1560
      Left            =   180
      Top             =   2805
      Width           =   6975
   End
   Begin VB.Shape f 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   2235
      Left            =   165
      Top             =   510
      Width           =   6975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3000
      Left            =   180
      Top             =   4485
      Width           =   7005
   End
End
Attribute VB_Name = "frmEmployeesDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Form_Load()
    CenterFrm Me
End Sub

Private Sub txtAddress_GotFocus()
    HighlightMe
End Sub

Private Sub txtBirthdate_GotFocus()
    HighlightMe
End Sub

Private Sub txtCity_GotFocus()
    HighlightMe
End Sub

Private Sub txtContactNo_GotFocus()
    HighlightMe
End Sub

Private Sub txtDepartment_GotFocus()
    HighlightMe
End Sub

Private Sub txtEmailAddress_GotFocus()
    HighlightMe
End Sub

Private Sub txtEmployeeID_GotFocus()
    HighlightMe
End Sub

Private Sub txtFirstName_GotFocus()
    HighlightMe
End Sub

Private Sub txtGender_GotFocus()
    HighlightMe
End Sub

Private Sub txtHireDate_GotFocus()
    HighlightMe
End Sub

Private Sub txtLastName_GotFocus()
    HighlightMe
End Sub


Private Sub txtMiddleName_GotFocus()
    HighlightMe
End Sub

Private Sub txtPosition_GotFocus()
    HighlightMe
End Sub


Private Sub txtZIP_GotFocus()
    HighlightMe
End Sub
