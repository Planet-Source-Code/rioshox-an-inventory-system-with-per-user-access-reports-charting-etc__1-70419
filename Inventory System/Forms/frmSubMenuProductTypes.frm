VERSION 5.00
Begin VB.Form frmSubMenuProductTypes 
   BackColor       =   &H00F5EADB&
   BorderStyle     =   0  'None
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3240
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
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
      Height          =   3480
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   285
   End
   Begin VB.Shape shpTitleBorder 
      BackColor       =   &H00A97332&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3240
      Left            =   2115
      Top             =   -15
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   0
      X2              =   2235
      Y1              =   1305
      Y2              =   1305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   2235
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H00A8612D&
      Caption         =   "  Product Types"
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
      TabIndex        =   0
      Top             =   0
      Width           =   2370
   End
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  New Product Type"
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
      Top             =   375
      Width           =   1770
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
   Begin VB.Label lblSubMenuProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "  Edit ProductType"
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
      Top             =   915
      Width           =   2175
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00C3BDB7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C3BDB7&
      Height          =   555
      Index           =   1
      Left            =   -15
      Top             =   765
      Visible         =   0   'False
      Width           =   2100
   End
End
Attribute VB_Name = "frmSubMenuProductTypes"
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
    
    lblToShow = "S" & vbCrLf & "H" & vbCrLf & "I" & vbCrLf & "P" & vbCrLf & "M" & vbCrLf & "E" & vbCrLf & "N" & vbCrLf & "T"
    lblToShow = "P" & vbCrLf & "R" & vbCrLf & "O" & vbCrLf & "D" & vbCrLf & "U" & vbCrLf & "C" & vbCrLf & "T" & vbCrLf & " " & vbCrLf & "T" & vbCrLf & "Y" & vbCrLf & "P" & vbCrLf & "E" & vbCrLf & "S"
    Set ActiveForm = Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isClicked = True
End Sub

Private Sub lblSubMenuProducts_Click(Index As Integer)

    frmLeftFrame.imgSubMenuFile_Click dMenu
    
    Select Case Index
        Case 0:
            frmProductTypesDE.txtProductTypeID = "AUTONUMEBR"
            frmProductTypesDE.txtProductTypeID.Locked = True
            frmProductTypesDE.Tag = TO_ADD
            frmProductTypesDE.Show 1
        Case 1:
            frmProductTypes.Show 1
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



