VERSION 5.00
Begin VB.Form frmLeftFrame 
   BackColor       =   &H00AA7E5A&
   BorderStyle     =   0  'None
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10815
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraLog 
      BackColor       =   &H00AA7E5A&
      BorderStyle     =   0  'None
      Height          =   2025
      Left            =   45
      TabIndex        =   14
      Top             =   8700
      Width           =   1845
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   19
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lblCurrentUser 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Current User:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00AA7E5A&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00AA7E5A&
         Height          =   1950
         Left            =   120
         Top             =   0
         Width           =   1770
      End
   End
   Begin VB.Frame frmMenu 
      BackColor       =   &H00AA7E5A&
      BorderStyle     =   0  'None
      Height          =   5955
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Label lblSubmenuFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Product Types"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   180
         TabIndex        =   12
         Top             =   2145
         Width           =   1320
      End
      Begin VB.Image imgSubMenuFile 
         Height          =   720
         Index           =   6
         Left            =   465
         Picture         =   "frmLeftFrame.frx":0000
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label lblSubmenuFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Packaging Types"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   11
         Top             =   3285
         Width           =   1530
      End
      Begin VB.Image imgSubMenuFile 
         Height          =   720
         Index           =   5
         Left            =   465
         Picture         =   "frmLeftFrame.frx":077B
         Top             =   2565
         Width           =   720
      End
      Begin VB.Image imgSubMenuFile 
         Height          =   720
         Index           =   2
         Left            =   465
         Picture         =   "frmLeftFrame.frx":0F4E
         Top             =   4815
         Width           =   720
      End
      Begin VB.Label lblSubmenuFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Users"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   375
         TabIndex        =   8
         Top             =   5565
         Width           =   900
      End
      Begin VB.Image imgSubMenuFile 
         Height          =   600
         Index           =   1
         Left            =   465
         Picture         =   "frmLeftFrame.frx":497D8
         Top             =   3660
         Width           =   600
      End
      Begin VB.Label lblSubmenuFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Customers"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   345
         TabIndex        =   7
         Top             =   4320
         Width           =   900
      End
      Begin VB.Label lblSubmenuFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Products"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   345
         TabIndex        =   5
         Top             =   1035
         Width           =   900
      End
      Begin VB.Image imgSubMenuFile 
         Height          =   720
         Index           =   0
         Left            =   465
         Picture         =   "frmLeftFrame.frx":49F6B
         Top             =   255
         Width           =   720
      End
   End
   Begin VB.Frame frmMenu 
      BackColor       =   &H00AA7E5A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0FF&
      Height          =   1395
      Index           =   2
      Left            =   75
      TabIndex        =   6
      Top             =   4635
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Image imgSubMenuFile 
         Height          =   570
         Index           =   4
         Left            =   570
         Picture         =   "frmLeftFrame.frx":4A7A8
         Stretch         =   -1  'True
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblSubmenuFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Database"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   105
         TabIndex        =   10
         Top             =   1035
         Width           =   1545
      End
   End
   Begin VB.Frame frmMenu 
      BackColor       =   &H00AA7E5A&
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   1
      Left            =   225
      TabIndex        =   4
      Top             =   1485
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Image imgSubMenuFile 
         Height          =   675
         Index           =   7
         Left            =   525
         Picture         =   "frmLeftFrame.frx":4ABEA
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label lblSubmenuFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shipments"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   195
         TabIndex        =   13
         Top             =   2085
         Width           =   1245
      End
      Begin VB.Label lblSubmenuFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   255
         TabIndex        =   9
         Top             =   1035
         Width           =   1245
      End
      Begin VB.Image imgSubMenuFile 
         Height          =   720
         Index           =   3
         Left            =   570
         Picture         =   "frmLeftFrame.frx":4B49A
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdLeftMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Utilities"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   480
      MouseIcon       =   "frmLeftFrame.frx":93D24
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3675
      Width           =   1920
   End
   Begin VB.CommandButton cmdLeftMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   0
      MouseIcon       =   "frmLeftFrame.frx":9402E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   255
      Width           =   1920
   End
   Begin VB.CommandButton cmdLeftMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   -15
      MouseIcon       =   "frmLeftFrame.frx":94338
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1050
      Width           =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   -180
      Top             =   6675
   End
End
Attribute VB_Name = "frmLeftFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Const MENU_COUNT As Integer = 3
Const BUTTON_HEIGHT As Long = 420
Const FRAME_HEIGHT As Long = 5610
Const IMAGE_MENU_COUNT As Integer = 8

Private Sub cmdLeftMenu_Click(Index As Integer)
    Dim ictr%, isFrameShown As Boolean
    
   
    
    isFullForm = False
    
    For ictr = 0 To MENU_COUNT - 1
        cmdLeftMenu(ictr).Left = 0
        If Not isFrameShown Then
            cmdLeftMenu(ictr).Top = BUTTON_HEIGHT * ictr
        Else
            cmdLeftMenu(ictr).Top = (BUTTON_HEIGHT * ictr) + frmMenu(Index).Height
            
        End If
        
        frmMenu(ictr).Visible = False
        
        If ictr = Index Then
            isFrameShown = True
            cmdLeftMenu(ictr).Top = BUTTON_HEIGHT * ictr
            
            frmMenu(ictr).Visible = True
            frmMenu(ictr).Left = 0
            frmMenu(ictr).Width = frmLeftFrame.Width
            frmMenu(ictr).Top = BUTTON_HEIGHT * (ictr + 1)
        End If
    Next ictr
    
    ActiveFrame% = Index
    ActiveTop% = cmdLeftMenu(Index).Top
    isFrameShown = False
    
    
    
End Sub

Private Sub Form_Load()
    Top = 0
    Height = Screen.Height
    Left = 0
    
    lblCurrentUser = g_CurrentUser.UserName
    cmdLeftMenu_Click 0
    

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'HideSubMenus
End Sub

Private Sub frmMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgSubMenuFile_MouseMove 100, 0, 0, 0, 0

    If isClicked Then imgSubMenuFile_Click dMenu
End Sub

Public Sub imgSubMenuFile_Click(Index As Integer)
    isFullForm = False
        
    If dMenu <> 99 Then If IsLoaded(ActiveForm) Then Unload ActiveForm
    
    
    If dMenu = 98 Then
        Unload ActiveForm
        Exit Sub
    End If
    
    isClicked = False
    
    dMenu = Index
    isFullForm = False
    
    
    Select Case Index
        
        Case 0:
               
            frmSubMenuProducts.Show
            frmSubMenuProducts.SetFocus
            
        Case 1:
            
            frmSubMenuCustomers.Show
            frmSubMenuCustomers.SetFocus
        Case 2:
            
            frmSubMenuemployees.Show
            frmSubMenuemployees.SetFocus
            
            
        Case 3:
            frmSubMenuSalesOrder.Show
            frmSubMenuSalesOrder.SetFocus
        Case 4:
            If Not g_CurrentUser.BackupData Then
                DeauthorizeNotify
                Exit Sub
            End If
            frmBackUpDatabase.Show 1
        Case 5:
            frmSubMenuPackagingTypes.Show
            frmSubMenuPackagingTypes.SetFocus
            
        Case 6:
            frmSubMenuProductTypes.Show
            frmSubMenuProductTypes.SetFocus
        Case 7:
            frmSubMenuShipments.Show
            frmSubMenuShipments.SetFocus
    End Select
    

End Sub


Private Sub imgSubMenuFile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ictr%
    
    If Index <> 100 Then If imgSubMenuFile(Index).BorderStyle = 1 Then Exit Sub
    For ictr = 0 To IMAGE_MENU_COUNT - 1
        If Index = ictr Then imgSubMenuFile(ictr).BorderStyle = 1 Else imgSubMenuFile(ictr).BorderStyle = 0
    Next
    
    soundIt "move"
End Sub

Private Sub Timer1_Timer()
    lblTime = Time
    lblDate = Format(Date, "medium date")
End Sub

