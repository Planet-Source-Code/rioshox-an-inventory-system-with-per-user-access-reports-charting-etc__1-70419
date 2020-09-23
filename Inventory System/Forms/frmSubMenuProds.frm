VERSION 5.00
Begin VB.Form frmWhoIsActive 
   BackColor       =   &H00A97332&
   BorderStyle     =   0  'None
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   75
      Top             =   2505
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   3060
      Left            =   0
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblToShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   285
   End
End
Attribute VB_Name = "frmWhoIsActive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Width = 1
    Timer1.Interval = 1
    Timer1.Enabled = True
End Sub


Private Sub lblToShow_Click()
    Select Case dMenu
        Case 0:
            frmSubMenuProducts.Show
            frmSubMenuProducts.SetFocus
        Case 1:
            frmSubMenuCustomers.Show
            frmSubMenuCustomers.SetFocus
    End Select
End Sub

Private Sub Timer1_Timer()
    Width = Width + 50
    If Width >= 480 Then

    Timer1.Enabled = False
    End If

End Sub
