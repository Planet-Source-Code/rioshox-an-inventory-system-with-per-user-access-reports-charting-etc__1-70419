VERSION 5.00
Begin VB.MDIForm MainForm 
   AutoShowChildren=   0   'False
   BackColor       =   &H00A34A18&
   Caption         =   "Ordering System via SMS with Sales and Inventory System for RFM Flour Division"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MainForm.frx":0442
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub MDIForm_DblClick()
    Generate4YrData Me
    Reprice Me
    UpdateBuyer Me
End Sub

Private Sub MDIForm_Load()
    'g_CurrentUser.EmployeeID = 1
    'g_CurrentUser.UserName = "Rio"
    'g_CurrentUser.IsAdmin = True
    
    Width = Screen.Width
    
    
    frmLeftFrame.Show
    
    
    dMenu = 99
    isFullForm = False
    IsThereLoaded = True
    
    Load frmTopBar
        
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not isFullForm Then Exit Sub
    WithSound = False
    If isClicked Then frmLeftFrame.imgSubMenuFile_Click dMenu
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo) = vbNo Then Cancel = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    connRFM.Close
    Set connRFM = Nothing
    End
End Sub

