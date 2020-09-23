VERSION 5.00
Begin VB.Form frmAuthorizeDiscountBatch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1485
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   735
      Width           =   2325
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
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "AED"
      Top             =   1770
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   2145
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "AED"
      Top             =   1770
      Width           =   945
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
      Left            =   1485
      TabIndex        =   3
      Top             =   1230
      Width           =   2325
   End
   Begin VB.Shape Shape1 
      Height          =   2115
      Left            =   90
      Top             =   30
      Width           =   3990
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   165
      TabIndex        =   0
      Top             =   765
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type the password to apply discount to the Transaction"
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
      Height          =   420
      Left            =   180
      TabIndex        =   6
      Top             =   135
      Width           =   4110
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   165
      TabIndex        =   2
      Top             =   1260
      Width           =   1215
   End
End
Attribute VB_Name = "frmAuthorizeDiscountBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not IsNumeric(txtDiscount) Then
        txtDiscount.SetFocus
        Exit Sub
    End If
    
    If txtPassword = GetSetting("RFM", "System Settings", "DiscountPassword", "RFM") Then
        If CDbl(txtDiscount) > CDbl(frmTransactionsDE.lblSubTotal) Then
            MsgBox "Cannot apply discount if discount is greater than the transaction total!", vbExclamation
            Exit Sub
        End If
        
        frmTransactionsDE.lblSubTotal.Tag = txtDiscount
        g_CurrentUser.AuthorizeDiscount = True
    Else
        g_CurrentUser.AuthorizeDiscount = False
        MsgBox "Only authorized user can edit the discount!", vbExclamation
        
    End If
    
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

