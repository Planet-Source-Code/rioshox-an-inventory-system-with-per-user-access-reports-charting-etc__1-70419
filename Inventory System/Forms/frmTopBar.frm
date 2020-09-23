VERSION 5.00
Begin VB.Form frmTopBar 
   BackColor       =   &H00FBF8F6&
   BorderStyle     =   0  'None
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6915
      Top             =   390
   End
   Begin VB.Label lblSmallNo 
      BackStyle       =   0  'Transparent
      Caption         =   "You are not authorized to access this part of the system"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   810
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Label lblBigNo 
      BackStyle       =   0  'Transparent
      Caption         =   "UNAUTHORIZED ACCESS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   8085
   End
End
Attribute VB_Name = "frmTopBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_Load()
Dim msgResult
    Top = 0
    Left = frmLeftFrame.Left + frmLeftFrame.Width
    'Width = Screen.Width - frmLeftFrame.Width
    
    
    
    

  


    
End Sub

Private Sub Timer1_Timer()
    lblBigNo.Visible = False
    lblSmallNo.Visible = False
    Unload Me
    frmTopBar.Timer1.Enabled = False
    
End Sub
