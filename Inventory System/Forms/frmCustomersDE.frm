VERSION 5.00
Begin VB.Form frmCustomersDE 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   Caption         =   "c"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSMSNumber 
      Height          =   330
      Left            =   240
      TabIndex        =   19
      Top             =   4395
      Width           =   1920
   End
   Begin VB.TextBox txtMinimumPurchase 
      Height          =   330
      Left            =   2205
      TabIndex        =   25
      Top             =   5490
      Width           =   1920
   End
   Begin VB.TextBox txtIPin 
      Height          =   330
      Left            =   2355
      TabIndex        =   21
      Top             =   4380
      Width           =   1920
   End
   Begin VB.TextBox txtCreditLimit 
      Height          =   330
      Left            =   150
      TabIndex        =   23
      Top             =   5490
      Width           =   1920
   End
   Begin VB.TextBox txtEmail 
      Height          =   330
      Left            =   2355
      TabIndex        =   17
      Top             =   3660
      Width           =   1920
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   330
      Left            =   3885
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5910
      Width           =   960
   End
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   330
      Left            =   4875
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5910
      Width           =   960
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print"
      Height          =   330
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5910
      Width           =   960
   End
   Begin VB.TextBox txtPhone 
      Height          =   330
      Left            =   225
      TabIndex        =   15
      Top             =   3705
      Width           =   1920
   End
   Begin VB.TextBox txtZip 
      Height          =   330
      Left            =   4725
      TabIndex        =   13
      Top             =   3000
      Width           =   1050
   End
   Begin VB.TextBox txtRegion 
      Height          =   330
      Left            =   2355
      TabIndex        =   11
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox txtCity 
      Height          =   330
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1920
   End
   Begin VB.TextBox txtAddress 
      Height          =   330
      Left            =   210
      TabIndex        =   7
      Top             =   2325
      Width           =   5595
   End
   Begin VB.TextBox txtContactName 
      Height          =   330
      Left            =   1605
      TabIndex        =   5
      Top             =   1470
      Width           =   2175
   End
   Begin VB.TextBox txtCompanyName 
      Height          =   330
      Left            =   1620
      TabIndex        =   3
      Top             =   990
      Width           =   3945
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   330
      Left            =   1620
      TabIndex        =   1
      Top             =   510
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cellphone Number (SMS Ordering)"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   18
      Top             =   4185
      Width           =   1770
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Purchase"
      Height          =   255
      Index           =   11
      Left            =   2205
      TabIndex        =   24
      Top             =   5280
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SMS I-Pin"
      Height          =   255
      Index           =   10
      Left            =   2355
      TabIndex        =   20
      Top             =   4170
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit"
      Height          =   255
      Index           =   9
      Left            =   150
      TabIndex        =   22
      Top             =   5280
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
      Height          =   255
      Index           =   8
      Left            =   2355
      TabIndex        =   16
      Top             =   3450
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
      Height          =   255
      Index           =   7
      Left            =   225
      TabIndex        =   14
      Top             =   3495
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ZIP"
      Height          =   255
      Index           =   6
      Left            =   4725
      TabIndex        =   12
      Top             =   2790
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Index           =   5
      Left            =   225
      TabIndex        =   6
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   2790
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   10
      Top             =   2790
      Width           =   1530
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   3075
      Left            =   75
      Top             =   2070
      Width           =   5850
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Code"
      Height          =   255
      Index           =   2
      Left            =   105
      TabIndex        =   0
      Top             =   600
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1050
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1530
      Width           =   1530
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Customer"
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
      Left            =   60
      TabIndex        =   29
      Top             =   75
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1665
      Left            =   75
      Top             =   345
      Width           =   5850
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1065
      Left            =   75
      Top             =   5220
      Width           =   5850
   End
End
Attribute VB_Name = "frmCustomersDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strSQL$, dmsg$
    If Trim(txtCustomerID) = "" Then
        txtCustomerID.SetFocus
        MsgBox "Customer Code cannot be empty...", vbCritical
        Exit Sub
    End If
    
    
    If Trim(txtCompanyName) = "" Then
        If Trim(txtContactName) = "" Then
            MsgBox "Company name and contact name cannot be BOTH empty...", vbCritical
            txtCompanyName.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txtPhone) = "" Then
        txtPhone.SetFocus
        MsgBox "Contact number cannot be empty...", vbCritical
        Exit Sub
    End If
    
     If Trim(txtIPin) = "" Then
        txtIPin.SetFocus
        MsgBox "SMS I-Pin is needed to enable ordering thru text", vbCritical
        Exit Sub
    End If
    
    If Not IsNumeric(txtCreditLimit) Then
        MsgBox "Invalid credit limit.", vbExclamation
        txtCreditLimit = 0#
        txtCreditLimit.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtMinimumPurchase) Then
        MsgBox "Invalid minimum puchase.", vbExclamation
        txtMinimumPurchase = 0#
        txtMinimumPurchase.SetFocus
        Exit Sub
    End If
    
    On Error GoTo cmdSave_Click_ERR
    Select Case Tag
        Case TO_ADD
            strSQL$ = _
                "INSERT INTO tblCustomers (SMSNumber,SMSIPin,MinimumPurchase,EmailAddress,CustomerID,CreditLimit,CompanyName ,ContactName,Address,City,Region,PostalCode,Phone) VALUES (" & _
                Enquote(txtSMSNumber, True) & _
                Enquote(txtIPin, True) & _
                EnNone(txtMinimumPurchase, True) & _
                Enquote(txtEmail, True) & _
                Enquote(txtCustomerID, True) & _
                EnNone(txtCreditLimit, True) & _
                Enquote(txtCompanyName, True) & _
                Enquote(txtContactName, True) & _
                Enquote(txtAddress, True) & _
                Enquote(txtCity, True) & _
                Enquote(txtRegion, True) & _
                Enquote(txtZIP, True) & _
                Enquote(txtPhone) & ")"
            dmsg$ = "New customer added."
        Case TO_EDIT
            strSQL$ = _
                "UPDATE tblCustomers SET " & _
                "SMSNumber = " & Enquote(txtSMSNumber, True) & _
                "SMSIPin = " & Enquote(txtIPin, True) & _
                "MinimumPurchase = " & EnNone(txtMinimumPurchase, True) & _
                "EmailAddress = " & Enquote(txtEmail, True) & _
                "CreditLimit = " & EnNone(txtCreditLimit, True) & _
                "CompanyName = " & Enquote(txtCompanyName, True) & _
                "ContactName = " & Enquote(txtContactName, True) & _
                "Address = " & Enquote(txtAddress, True) & _
                "City = " & Enquote(txtCity, True) & _
                "Region = " & Enquote(txtRegion, True) & _
                "PostalCode = " & Enquote(txtZIP, True) & _
                "Phone = " & Enquote(txtPhone) & "WHERE CustomerID = " & _
                Enquote(txtCustomerID)
                
            dmsg$ = "Changes to customer saved."
        
    End Select
    
    
    
    connRFM.Execute strSQL
    MsgBox dmsg, vbInformation
    
    Unload Me
    
cmdSave_Click_ERR:
    If Err.Number <> 0 Then MsgBox "Cannot continue current operation because an error occurred." & vbCrLf & Err.Description, vbCritical
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub



Private Sub Form_Load()
    If Not g_CurrentUser.Customer Then cmdSave.Enabled = False
    CenterFrm Me
End Sub

Private Sub txtAddress_GotFocus()
    HighlightMe
End Sub

Private Sub txtAddress_LostFocus()
    txtAddress = Propercase(txtAddress)
End Sub

Private Sub txtCity_GotFocus()
    HighlightMe
End Sub

Private Sub txtCompanyName_GotFocus()
    HighlightMe
End Sub


Private Sub txtCompanyName_LostFocus()
    txtCompanyName = Propercase(txtCompanyName)
End Sub

Private Sub txtContactName_GotFocus()
    HighlightMe
End Sub

Private Sub txtContactName_LostFocus()
    txtContactName = Propercase(txtContactName)
End Sub

Private Sub txtCountry_GotFocus()
    HighlightMe
End Sub

Private Sub txtCreditLimit_GotFocus()
    HighlightMe
End Sub

Private Sub txtCustomerID_GotFocus()
    HighlightMe
End Sub

Private Sub txtEmail_GotFocus()
    HighlightMe
End Sub

Private Sub txtIPin_GotFocus()
    HighlightMe
End Sub


Private Sub txtMinimumPurchase_GotFocus()
    HighlightMe
End Sub

Private Sub txtPhone_GotFocus()
    HighlightMe
End Sub

Private Sub txtZIP_GotFocus()
    HighlightMe
End Sub
