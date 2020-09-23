VERSION 5.00
Begin VB.Form frmAuthorizeDiscount 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "6"
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optMode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "For All"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   1650
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "For Each"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   705
      TabIndex        =   5
      Top             =   1650
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtDiscount 
      Appearance      =   0  'Flat
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
      Left            =   1575
      TabIndex        =   4
      Top             =   1155
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   2385
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "AED"
      Top             =   1755
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
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
      Height          =   285
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "AED"
      Top             =   1755
      Width           =   945
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
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
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   660
      Width           =   2325
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   90
      Top             =   15
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Left            =   255
      TabIndex        =   3
      Top             =   1185
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Type the password to unlock the discount field"
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
      Left            =   195
      TabIndex        =   0
      Top             =   90
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   1
      Top             =   690
      Width           =   1215
   End
End
Attribute VB_Name = "frmAuthorizeDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dMode As Integer
Private Sub cmdOK_Click()
    If Not IsNumeric(txtDiscount) Then
        MsgBox "Invalid discount!", vbCritical
        txtDiscount.SetFocus
        Exit Sub
    End If
    
    If txtPassword = GetSetting("RFM", "System Settings", "DiscountPassword", "RFM") Then
        If dMode = 0 Then
                If CDbl(txtDiscount) > 0 Then
                    LineDiscount = Format(CDbl(txtDiscount) / CDbl(txtPassword.Tag), "#.000000")
                Else
                    LineDiscount = 0
                End If
        End If
    Else
        LineDiscount = 0
        MsgBox "Only authorized user can edit the discount!", vbExclamation
    End If
        
    
    Unload Me
End Sub

Private Sub Command1_Click()
    
    Unload Me
End Sub

Private Sub Form_Load()
    g_CurrentUser.AuthorizeDiscount = False
End Sub

Private Sub optMode_Click(Index As Integer)
    dMode = Index
End Sub
