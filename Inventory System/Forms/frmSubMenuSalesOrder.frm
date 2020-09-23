VERSION 5.00
Begin VB.Form frmSubMenuSalesOrder 
   BackColor       =   &H00F5EADB&
   BorderStyle     =   0  'None
   Caption         =   "r"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2355
      Top             =   4530
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
      Height          =   3120
      Left            =   2085
      TabIndex        =   0
      Top             =   75
      Width           =   285
   End
   Begin VB.Shape shpTitleBorder 
      BackColor       =   &H00A97332&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3675
      Left            =   2055
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H00A8612D&
      Caption         =   "  Return Items"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   1785
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00A8612D&
      Caption         =   "  Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   15
      Width           =   2415
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  Pending SO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   375
      Width           =   1770
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  SO Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1365
      Width           =   2175
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  Return Item"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   2175
      Width           =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   2235
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   0
      X2              =   2235
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   2235
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   420
      Index           =   2
      Left            =   0
      Top             =   2070
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   555
      Index           =   1
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   420
      Index           =   0
      Left            =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label3 
      BackColor       =   &H00A8612D&
      Caption         =   "  Report(s)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   2505
      Width           =   2415
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  SO Summary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   2895
      Width           =   2160
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   420
      Index           =   3
      Left            =   0
      Top             =   2790
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   2235
      Y1              =   3225
      Y2              =   3225
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  SO Chart"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   3345
      Width           =   2160
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   420
      Index           =   4
      Left            =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   0
      X2              =   2235
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  Posted SO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   0
      TabIndex        =   9
      Top             =   855
      Width           =   2160
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   420
      Index           =   5
      Left            =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   2100
   End
End
Attribute VB_Name = "frmSubMenuSalesOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SUB_MENU_COUNT As Integer = 6

Dim dWidth%
Private Sub Form_Load()
    Top = (frmLeftFrame.Top + frmLeftFrame.Height / 2) - Height
    Left = frmLeftFrame.Left + frmLeftFrame.Width
    
    
    If Not isFullForm Then
        dWidth% = shpTitleBorder.Width
        shpTitleBorder.Left = 0
        lblToShow.Left = shpTitleBorder.Left + 30
        lblToShow.Enabled = True
    Else
        dWidth = 2400
        lblToShow.Enabled = False
    End If
    
    Width = 1
    Timer1.Interval = 1
    Timer1.Enabled = True
    
    isFullForm = True
    
    lblToShow = "S" & vbCrLf & "A" & vbCrLf & "L" & vbCrLf & "E" & vbCrLf & "S" & vbCrLf & " " & vbCrLf & "O" & vbCrLf & "R" & vbCrLf & "D" & vbCrLf & "E" & vbCrLf & "R"
    
    Set ActiveForm = Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isClicked = True
End Sub

Private Sub lblSubMenuProducts_Click(Index As Integer)

    frmLeftFrame.imgSubMenuFile_Click dMenu
    
    Select Case Index
        Case 0:
            If Not g_CurrentUser.ApproveSO Then
                DeauthorizeNotify
                Exit Sub
            End If
            
            frmPendingSO.lblTitle = "Pending Order(s)"
            frmPendingSO.cmdShip.Visible = False
            frmPendingSO.cboFilterStatus.ListIndex = 0
            frmPendingSO.Show 1
        Case 1:
            If Not g_CurrentUser.ManualSO Then
                DeauthorizeNotify
                Exit Sub
            End If
            frmTransactionsDE.Show 1
        Case 2:
            If Not g_CurrentUser.ReturnItem Then
                DeauthorizeNotify
                Exit Sub
            End If
            frmReturnItems.Show 1
        Case 3:
            frmPrintCustomerTransactions.Show 1
        Case 4:
            frmCharter.Show 1
        Case 5:
            frmPostedSO.Show 1
    End Select
End Sub

Private Sub lblSubMenuProducts_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ictr%
    
    For ictr = 0 To SUB_MENU_COUNT - 1
        If Index = ictr Then Me.shpMenu(ictr).Visible = True Else Me.shpMenu(ictr).Visible = False
    Next
End Sub

Private Sub lblToShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
    Show

End Sub

Private Sub Timer1_Timer()
  
        Width = Width + dWidth% / 20
        If Width >= dWidth% Then
        Timer1.Enabled = False
        End If
  

End Sub



