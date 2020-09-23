VERSION 5.00
Begin VB.Form frmSubMenuShipments 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5EADB&
   BorderStyle     =   0  'None
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   2475
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
      Height          =   2970
      Left            =   2085
      TabIndex        =   0
      Top             =   75
      Width           =   285
   End
   Begin VB.Shape shpTitleBorder 
      BackColor       =   &H00A97332&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   2820
      Left            =   2055
      Top             =   0
      Width           =   360
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
      TabIndex        =   4
      Top             =   750
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   0
      X2              =   2235
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   2235
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  Shipment Reports"
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
      TabIndex        =   3
      Top             =   1215
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  Pending Shipments"
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
      TabIndex        =   2
      Top             =   345
      Width           =   1770
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
      TabIndex        =   1
      Top             =   15
      Width           =   2415
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   555
      Index           =   1
      Left            =   0
      Top             =   1050
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   -15
      X2              =   2220
      Y1              =   1035
      Y2              =   1035
   End
End
Attribute VB_Name = "frmSubMenuShipments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const SUB_MENU_COUNT As Integer = 2
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
        dWidth = 2430
        lblToShow.Enabled = False
    End If
    
    Width = 1
    Timer1.Interval = 1
    Timer1.Enabled = True
    
    isFullForm = True
    
     lblToShow = vbCrLf & "S" & vbCrLf & "H" & vbCrLf & "I" & vbCrLf & "P" & vbCrLf & "M" & vbCrLf & "E" & vbCrLf & "N" & vbCrLf & "T" & vbCrLf & "S"
    
    Set ActiveForm = Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isClicked = True
End Sub

Private Sub lblSubMenuProducts_Click(Index As Integer)

    frmLeftFrame.imgSubMenuFile_Click dMenu
    
    Select Case Index
        Case 0:
            If Not g_CurrentUser.ApproveShip Then
                DeauthorizeNotify
                Exit Sub
            End If
            frmPendingSO.lblTitle = "Pending Shipment(s)"
            frmPendingSO.cmdApproved.Visible = False
            frmPendingSO.cboFilterStatus.ListIndex = 1
            frmPendingSO.Show 1
        Case 1:
            
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


