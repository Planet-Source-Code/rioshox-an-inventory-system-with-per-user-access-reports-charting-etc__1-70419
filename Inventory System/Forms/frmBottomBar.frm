VERSION 5.00
Begin VB.Form frmBottomBar 
   BackColor       =   &H00DBB493&
   BorderStyle     =   0  'None
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10485
   FillColor       =   &H00F8F8F8&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   270
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   -120
      Width           =   11910
   End
End
Attribute VB_Name = "frmBottomBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Top = frmLeftFrame.Top + frmLeftFrame.Height + 20
    Left = frmLeftFrame.Left
    Width = Screen.Width
    'Shape1.Width = frmLeftFrame.Width
    
End Sub
