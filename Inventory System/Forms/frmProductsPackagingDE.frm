VERSION 5.00
Begin VB.Form frmProductsPackagingDE 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   Caption         =   "\"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1875
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
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
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1875
      Width           =   960
   End
   Begin VB.TextBox txtPackagingType 
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
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   4740
   End
   Begin VB.TextBox txtPackagingTypeID 
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
      Left            =   240
      TabIndex        =   0
      Top             =   750
      Width           =   2460
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Packaging Code"
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
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Packaging"
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
      Left            =   225
      TabIndex        =   5
      Top             =   1200
      Width           =   2190
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Product Packaging"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   120
      Width           =   3945
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1845
      Left            =   60
      Top             =   450
      Width           =   5190
   End
End
Attribute VB_Name = "frmProductsPackagingDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strSQL$, msg$
    
    If Trim(txtPackagingType.Text) = "" Then
        MsgBox "Product packaging cannot be empty..", vbExclamation
        txtPackagingType.SetFocus
        Exit Sub
    End If
    
    On Error GoTo cmdSave_Click_ERR
    
    Select Case Tag
        Case TO_ADD:
            strSQL$ = _
                "INSERT INTO tblPackagingTypes (PackagingType) VALUES (" & _
                    Enquote(txtPackagingType) & ")"
                    
            msg$ = "New product type added."
            
        Case TO_EDIT:
            strSQL$ = _
                "UPDATE tblProductTypes SET " & _
                "PackagingType = " & Enquote(txtPackagingType) & _
                " WHERE PackagingTypeID= " & txtPackagingTypeID
            
            msg$ = "Changes saved."
    End Select
    
    
    connRFM.Execute strSQL, adExecuteNoRecords
    MsgBox msg$, vbInformation
    
    If Tag = TO_EDIT Then
        With frmProductsPackaging
            .msfSearchResult.TextMatrix(.msfSearchResult.Row, 0) = Format(txtPackagingTypeID, "000")
            .msfSearchResult.TextMatrix(.msfSearchResult.Row, 1) = txtPackagingType
        End With
    End If

    
    Unload Me

cmdSave_Click_ERR:
    If Err.Number <> 0 Then
        MsgBox "Errors occurred, cannot perform current operation." & vbCrLf & Err.Description, vbCritical
    End If
    

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    If Not g_CurrentUser.Products Then cmdSave.Enabled = False
        CenterFrm Me
End Sub

Private Sub txtPackagingType_GotFocus()
    HighlightMe
End Sub

Private Sub txtPackagingTypeID_GotFocus()
    HighlightMe
End Sub


