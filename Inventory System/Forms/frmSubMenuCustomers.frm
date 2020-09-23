VERSION 5.00
Begin VB.Form frmSubMenuCustomers 
   BackColor       =   &H00F5EADB&
   BorderStyle     =   0  'None
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   720
      Top             =   480
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
      Height          =   1935
      Left            =   2145
      TabIndex        =   4
      Top             =   90
      Width           =   285
   End
   Begin VB.Shape shpTitleBorder 
      BackColor       =   &H00A97332&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   2070
      Left            =   2115
      Top             =   0
      Width           =   360
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
      TabIndex        =   3
      Top             =   0
      Width           =   3090
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  Add New Customer"
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
      Top             =   360
      Width           =   1770
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  Edit Existing Customer"
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
      TabIndex        =   1
      Top             =   900
      Width           =   2175
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  Customer Transactions"
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
      TabIndex        =   0
      Top             =   1710
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   2235
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   420
      Index           =   0
      Left            =   0
      Top             =   285
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
      Top             =   735
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   420
      Index           =   2
      Left            =   0
      Top             =   1605
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   0
      X2              =   2235
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   2235
      Y1              =   1305
      Y2              =   1305
   End
   Begin VB.Label Label2 
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
      Left            =   15
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmSubMenuCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const SUB_MENU_COUNT As Integer = 3
Dim dWidth%
Private Sub Form_Load()
    frmSubMenuCustomers.Top = (frmLeftFrame.Top + frmLeftFrame.Height / 2) - Height
    frmSubMenuCustomers.Left = frmLeftFrame.Left + frmLeftFrame.Width
    
    
    If Not isFullForm Then
        dWidth% = shpTitleBorder.Width
        shpTitleBorder.Left = 0
        lblToShow.Left = shpTitleBorder.Left + 30
        lblToShow.Enabled = True
    Else
        dWidth = 2475
        lblToShow.Enabled = False
    End If
    
    Width = 1
    Timer1.Interval = 1
    Timer1.Enabled = True
    
    isFullForm = True
    
    lblToShow = "C" & vbCrLf & "U" & vbCrLf & "S" & vbCrLf & "T" & vbCrLf & "O" & vbCrLf & "M" & vbCrLf & "E" & vbCrLf & "R" & vbCrLf & "S"
    
    Set ActiveForm = Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isClicked = True
End Sub

Private Sub lblSubMenuProducts_Click(Index As Integer)

    frmLeftFrame.imgSubMenuFile_Click dMenu
    
    Select Case Index
        Case 0:
            frmCustomersDE.Tag = TO_ADD
            frmCustomersDE.Show 1
            
        Case 1:
            frmCustomers.Show 1
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
