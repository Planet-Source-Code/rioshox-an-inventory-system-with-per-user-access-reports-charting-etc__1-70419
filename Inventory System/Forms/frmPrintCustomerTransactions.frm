VERSION 5.00
Begin VB.Form frmPrintCustomerTransactions 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   Caption         =   "0"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2985
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkGroupMonthly 
      BackColor       =   &H00F5EADB&
      Caption         =   "Print In Group (Per Products)"
      Height          =   240
      Left            =   4125
      TabIndex        =   14
      Top             =   2160
      Width           =   2760
   End
   Begin VB.ComboBox cboDailyGroup 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmPrintCustomerTransactions.frx":0000
      Left            =   570
      List            =   "frmPrintCustomerTransactions.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2055
      Width           =   2085
   End
   Begin VB.CheckBox chkGroupDaily 
      BackColor       =   &H00F5EADB&
      Caption         =   "Print In Group"
      Height          =   240
      Left            =   600
      TabIndex        =   5
      Top             =   1755
      Width           =   2760
   End
   Begin VB.TextBox txtTransactionYear 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5490
      TabIndex        =   13
      Top             =   1800
      Width           =   1830
   End
   Begin VB.ComboBox cboEndMonth 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5490
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1350
      Width           =   1830
   End
   Begin VB.ComboBox cboStartMonth 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5490
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   930
      Width           =   1830
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print"
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
      Left            =   9180
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2535
      Width           =   960
   End
   Begin VB.TextBox txtStartYear 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9120
      TabIndex        =   17
      Top             =   975
      Width           =   1830
   End
   Begin VB.TextBox txtEndYear 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9120
      TabIndex        =   19
      Top             =   1350
      Width           =   1830
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
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
      Left            =   10245
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2535
      Width           =   960
   End
   Begin VB.OptionButton optPrintWhat 
      BackColor       =   &H00F5EADB&
      Caption         =   "Yearly Sales"
      Height          =   300
      Index           =   2
      Left            =   7770
      TabIndex        =   15
      Top             =   645
      Width           =   3315
   End
   Begin VB.OptionButton optPrintWhat 
      BackColor       =   &H00F5EADB&
      Caption         =   "Monthly Sales"
      Height          =   300
      Index           =   1
      Left            =   3900
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton optPrintWhat 
      BackColor       =   &H00F5EADB&
      Caption         =   "Daily Sales"
      Height          =   300
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   615
      Value           =   -1  'True
      Width           =   1980
   End
   Begin VB.TextBox txtEndDate 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1650
      TabIndex        =   4
      Top             =   1320
      Width           =   1830
   End
   Begin VB.TextBox txtStartDate 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1650
      TabIndex        =   2
      Top             =   900
      Width           =   1830
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   4140
      TabIndex        =   12
      Top             =   1845
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   4140
      TabIndex        =   8
      Top             =   990
      Width           =   1830
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4140
      TabIndex        =   10
      Top             =   1425
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   7980
      TabIndex        =   16
      Top             =   1005
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   7980
      TabIndex        =   18
      Top             =   1410
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   510
      TabIndex        =   3
      Top             =   1350
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   510
      TabIndex        =   1
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Transactions History"
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
      Left            =   150
      TabIndex        =   22
      Top             =   165
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1875
      Left            =   240
      Top             =   570
      Width           =   3510
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1890
      Left            =   7710
      Top             =   570
      Width           =   3495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1890
      Left            =   3840
      Top             =   555
      Width           =   3765
   End
End
Attribute VB_Name = "frmPrintCustomerTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ToPrint%



Private Sub cmdPrint_Click()
    Dim tmpDate$, tmpMonth%
    Dim dRpt$, dFilter$
    On Error GoTo cmdPrint_Click_ERR
    Select Case ToPrint
    
        Case 0:
            If Not IsDate(txtStartDate) Then
                txtStartDate.SetFocus
                MsgBox "Invalid start date!", vbCritical
                Exit Sub
            End If
            
            
            If Not IsDate(txtEndDate) Then
                txtEndDate.SetFocus
                MsgBox "Invalid end date!", vbCritical
                Exit Sub
            End If
            
            If CDate(txtEndDate) < CDate(txtStartDate) Then
                tmpDate$ = txtEndDate
                txtEndDate = txtStartDate
                txtStartDate = tmpDate$
            End If
            
            If chkGroupDaily.Value = Checked Then
                If cboDailyGroup.ListIndex = 0 Then
                    dRpt$ = "rptSalesByDatePerCustomer"
                Else
                    dRpt$ = "rptSalesByDatePerProduct"
                End If
            
            Else
                dRpt$ = "rptSalesByDate"
            End If
            dFilter$ = "OrderDate BETWEEN " & EnPound(txtStartDate) & " AND " & EnPound(txtEndDate)
            
        Case 1:
            If Not IsNumeric(txtTransactionYear) Then
                txtTransactionYear.SetFocus
                MsgBox "Invalid transaction year!", vbCritical
                Exit Sub
            End If
            
            If cboEndMonth.ListIndex < cboStartMonth.ListIndex Then
                tmpMonth% = cboEndMonth.ListIndex
                cboEndMonth.ListIndex = cboStartMonth.ListIndex
                cboStartMonth.ListIndex = tmpMonth%
            End If
                        
            If chkGroupMonthly.Value = Checked Then
                dRpt$ = "rptSalesByMonthPerProduct"
                dFilter$ = "(SalesMonth BETWEEN " & cboStartMonth.ListIndex + 1 & " AND " & cboEndMonth.ListIndex + 1 & ") AND SalesYear = " & txtTransactionYear
            Else
                dRpt$ = "rptSalesByMonth"
                dFilter$ = "(dMonth BETWEEN " & cboStartMonth.ListIndex + 1 & " AND " & cboEndMonth.ListIndex + 1 & ") AND dYear = " & Enquote(txtTransactionYear)
            End If
            
        Case 2:
            If Not IsNumeric(txtStartYear) Then
                txtStartYear.SetFocus
                MsgBox "Invalid start year!", vbCritical
                Exit Sub
            End If
            
            If Not IsNumeric(txtEndYear) Then
                txtEndYear.SetFocus
                MsgBox "Invalid end year!", vbCritical
                Exit Sub
            End If
            
            If txtEndYear < txtStartYear Then
                tmpMonth% = txtEndYear
                txtEndYear = txtStartYear
                txtStartYear = tmpMonth%
            End If
            
            dRpt$ = "rptSalesByYear"
            dFilter$ = "SalesYear BETWEEN " & Enquote(txtStartYear) & " AND " & Enquote(txtEndYear)
    End Select
    
        ShowReport dRpt, dFilter
    
cmdPrint_Click_ERR:
    If Err.Number <> 0 Then
        MsgBox "Cannot continue current operation because an error occurred." & vbCrLf & Err.Description, vbCritical
    End If
    
    'ShowReport "rptTransactionsDetailed", "OrderDate BETWEEN " & EnPound(txtStartDate) & " AND " & EnPound(txtEndDate)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Dim ictr%
    
    txtStartDate = Format(Date - 7, "mm/dd/yyyy")
    txtEndDate = Format(Date, "mm/dd/yyyy")
    
    CenterFrm Me
    
    For ictr% = 1 To 12
        cboStartMonth.AddItem MonthName(ictr)
        cboEndMonth.AddItem MonthName(ictr)
    Next
    
    cboDailyGroup.ListIndex = 0
End Sub

Private Sub optPrintWhat_Click(Index As Integer)
    ToPrint% = Index
End Sub

Private Sub txtEndDate_GotFocus()
    HighlightMe
End Sub

Private Sub txtEndDate_LostFocus()
    If IsDate(txtEndDate) Then txtEndDate = Format(txtEndDate, "mm/dd/yyyy")
End Sub

Private Sub txtEndYear_GotFocus()
    HighlightMe
End Sub

Private Sub txtStartDate_GotFocus()
    HighlightMe
End Sub

Private Sub txtStartDate_LostFocus()
    If IsDate(txtStartDate) Then txtStartDate = Format(txtStartDate, "mm/dd/yyyy")
End Sub


Private Sub txtStartYear_GotFocus()
    HighlightMe
End Sub

Private Sub txtTransactionYear_GotFocus()
    HighlightMe
End Sub
