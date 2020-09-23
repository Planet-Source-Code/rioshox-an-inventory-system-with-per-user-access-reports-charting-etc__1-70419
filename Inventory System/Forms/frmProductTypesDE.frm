VERSION 5.00
Begin VB.Form frmProductTypesDE 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtProductTypeID 
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
      Top             =   705
      Width           =   2460
   End
   Begin VB.TextBox txtProductType 
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
      TabIndex        =   3
      Top             =   1395
      Width           =   4740
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
      TabIndex        =   4
      Top             =   1830
      Width           =   960
   End
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
      TabIndex        =   5
      Top             =   1830
      Width           =   960
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Product Type"
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
      TabIndex        =   6
      Top             =   75
      Width           =   3945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Type"
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
      TabIndex        =   2
      Top             =   1155
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Type Code"
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
      TabIndex        =   0
      Top             =   435
      Width           =   2190
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1845
      Left            =   60
      Top             =   405
      Width           =   5070
   End
End
Attribute VB_Name = "frmProductTypesDE"
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
    
    If Trim(txtProductType.Text) = "" Then
        MsgBox "Product type cannot be empty..", vbExclamation
        txtProductType.SetFocus
        Exit Sub
    End If
    
    
    
    On Error GoTo cmdSave_Click_ERR
    
    Select Case Tag
        Case TO_ADD:
            strSQL$ = _
                "INSERT INTO tblProductTypes (ProductType) VALUES (" & _
                    Enquote(txtProductType) & ")"
                    
            msg$ = "New product type added."
            
        Case TO_EDIT:
            strSQL$ = _
                "UPDATE tblProductTypes SET " & _
                "ProductType = " & Enquote(txtProductType) & _
                " WHERE ProductTypeID = " & txtProductTypeID
            
            msg$ = "Changes saved."
    End Select
    
    
    connRFM.Execute strSQL, adExecuteNoRecords
    MsgBox msg$, vbInformation
    
    If Tag = TO_EDIT Then
        With frmProductTypes
            .msfSearchResult.TextMatrix(.msfSearchResult.Row, 0) = Format(txtProductTypeID, "000")
            .msfSearchResult.TextMatrix(.msfSearchResult.Row, 1) = txtProductType
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

Private Sub txtProductType_GotFocus()
    HighlightMe
End Sub

Private Sub txtProductTypeID_GotFocus()
    HighlightMe
End Sub
