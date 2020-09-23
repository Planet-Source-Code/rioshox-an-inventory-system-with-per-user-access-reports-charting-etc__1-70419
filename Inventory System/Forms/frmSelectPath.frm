VERSION 5.00
Begin VB.Form frmSelectPath 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DBB493&
   BorderStyle     =   0  'None
   Caption         =   "Select Path ..."
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   5835
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton LaVolpeButton1 
      BackColor       =   &H00FFFEFC&
      Caption         =   "Close"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2415
      Width           =   885
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   2115
      Left            =   255
      TabIndex        =   1
      Top             =   825
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   255
      TabIndex        =   0
      Top             =   480
      Width           =   2430
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Backup Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   285
      TabIndex        =   4
      Top             =   210
      Width           =   2820
   End
   Begin VB.Label lblBackupPath 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2910
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Selected Path ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2895
      TabIndex        =   2
      Top             =   615
      Width           =   2895
   End
End
Attribute VB_Name = "frmSelectPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    lblBackupPath.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo A1:
    Dir1.Path = Drive1.Drive
    Exit Sub
A1:
    MsgBox "Drive Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Drive1.Drive = "C:\"
    
    Me.Dir1.Path = App.Path & "\Backup"
    
    
    If Err.Number <> 0 Then
        MsgBox "Errors occurred." & vbCrLf & Err.num & ": " & Err.Description & vbCrLf & "You may have logged as system administrator on a client computer.", vbCritical
        Me.Dir1.Path = App.Path
    End If
    
    lblBackupPath.Caption = Dir1.Path
End Sub

Private Sub LaVolpeButton1_Click()
    frmBackUpDatabase.txtBackUpPath.Text = IIf(Right(lblBackupPath.Caption, 1) = "\", lblBackupPath.Caption, lblBackupPath.Caption & "\")
    If frmBackUpDatabase.txtBackUpPath = App.Path & "\" Then
        MsgBox "Please choose a different path, cannot create backup in the same folder where the system is located.", vbCritical, "Change backup path"
        Exit Sub
    End If
    Unload Me
End Sub

