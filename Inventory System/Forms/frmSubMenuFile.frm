VERSION 5.00
Begin VB.Form frmSubMenuFile 
   BackColor       =   &H00DBB493&
   BorderStyle     =   0  'None
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   885
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblSubMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Types"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   405
      TabIndex        =   3
      Top             =   345
      Width           =   1965
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000CCFF&
      Index           =   5
      X1              =   45
      X2              =   1680
      Y1              =   -15
      Y2              =   -15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DBB493&
      Index           =   4
      X1              =   90
      X2              =   1725
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DBB493&
      Index           =   3
      X1              =   120
      X2              =   1755
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00A88062&
      Index           =   2
      X1              =   135
      X2              =   1770
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00A88062&
      Index           =   1
      X1              =   105
      X2              =   1740
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00A88062&
      Index           =   0
      X1              =   165
      X2              =   1800
      Y1              =   285
      Y2              =   285
   End
   Begin VB.Label lblSubMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Employees"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   405
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label lblSubMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   405
      TabIndex        =   1
      Top             =   615
      Width           =   1965
   End
   Begin VB.Label lblSubMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Products"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   405
      TabIndex        =   0
      Top             =   75
      Width           =   1965
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F8F8F8&
      BorderColor     =   &H00A88062&
      FillColor       =   &H00A88062&
      FillStyle       =   0  'Solid
      Height          =   1440
      Index           =   0
      Left            =   -30
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "frmSubMenuFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const intFileSubMenuCount       As Integer = 4
Const MENU_SELECTED             As Long = 16737304
Const MENU_NOT_SELECTED         As Long = vbWhite


Private Sub Form_Load()
    Dim ictr%
    
    For ictr = 0 To intFileSubMenuCount - 1
        'lblSubMenu(ictr).ForeColor = MENU_NOT_SELECTED
    Next
    
    Left = frmTopBar.lblMenuItem(0).Left
    Top = frmTopBar.Shape1.Top + frmTopBar.Shape1.Height
    
End Sub


Private Sub lblSubMenu_Click(Index As Integer)
    Hide
    DoEvents
    Select Case Index
        Case 0:
            
        Case 1:
            
        Case 2:
            
        Case 3:
            
            
        
            
        
    End Select
End Sub

Private Sub lblSubMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ictr%
     
    If lblSubMenu(Index).ForeColor = MENU_SELECTED Then Exit Sub
    
    For ictr = 0 To intFileSubMenuCount - 1
        If ictr = Index Then lblSubMenu(ictr).ForeColor = MENU_SELECTED Else lblSubMenu(ictr).ForeColor = MENU_NOT_SELECTED
    Next
End Sub



