VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackUpDatabase 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Back up Database ..."
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   Icon            =   "frmBackUpDatabase.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCreateBackup 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2190
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   7020
      TabIndex        =   3
      ToolTipText     =   "Click To Select BackUp Path"
      Top             =   1425
      Width           =   375
   End
   Begin VB.TextBox txtBackUpPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2700
      TabIndex        =   6
      Top             =   2715
      Width           =   4260
   End
   Begin VB.Image Image2 
      Height          =   4830
      Left            =   0
      Picture         =   "frmBackUpDatabase.frx":0E42
      Top             =   -120
      Width           =   2490
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Path Where to Store Backup"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Path and Click on Create Backup."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "frmBackUpDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FS As New Scripting.FileSystemObject

Private Sub cmdCreateBackup_Click()
    Dim BackUpPath$
    Dim NewFileName$
        
    BackUpPath = txtBackUpPath.Text
    If BackUpPath = "" Then Exit Sub
    ProgressBar1.Visible = True
    ProgressBar1.Value = 10
    NewFileName = CStr(Format(Date, "mmddyyyy")) & CStr(Format(Time, "hhmmss")) & ".mdb"
        
    
    Screen.MousePointer = vbHourglass
    lblStatus.Caption = App.Path & "\report file to " & BackUpPath
    DoEvents
    FS.CopyFile App.Path & "\RFM.mdb", BackUpPath, True
    ProgressBar1.Value = 50
    
    lblStatus.Caption = "Backing up SQL Server Data"
    DoEvents
    connRFM.Execute _
        "BACKUP DATABASE RFM " & _
        "TO DISK = 'c:\RFM Backup\" & Left(NewFileName, Len(NewFileName) - 4) & ".bak" & "'"


    
    Name BackUpPath & "RFM.mdb" As BackUpPath & NewFileName
    lblStatus.Caption = "Backup Process Finished"
    DoEvents
    ProgressBar1.Value = 100
    Screen.MousePointer = vbDefault
        
    MsgBox "Done backing up system database. It is advised that you backup the system database regularly for you to be able to restore your data in case of file corruption .", vbInformation
    ProgressBar1.Visible = False
    lblStatus.Caption = ""
    
    'logtrail "Back Up database"
    
    Unload Me
End Sub

Private Sub Command1_Click()
    frmSelectPath.Show 1
End Sub


Private Sub Form_Activate()
    'updatestatus "Backing up database"
End Sub

Private Sub Form_Load()
    If Not FS.FolderExists(App.Path & "\Backup") Then FS.CreateFolder (App.Path & "\Backup")
End Sub
