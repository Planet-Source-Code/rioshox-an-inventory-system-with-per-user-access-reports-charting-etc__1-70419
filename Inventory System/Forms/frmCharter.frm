VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmCharter 
   BackColor       =   &H00C18B59&
   BorderStyle     =   0  'None
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   315
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Chart"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4695
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1260
      Width           =   1155
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
      Left            =   660
      TabIndex        =   3
      Top             =   1290
      Width           =   1620
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
      Left            =   2595
      TabIndex        =   5
      Top             =   1305
      Width           =   1620
   End
   Begin VB.OptionButton optDMode 
      BackColor       =   &H00F5EADB&
      Caption         =   "Compare to Years"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1965
      TabIndex        =   1
      Top             =   735
      Width           =   2310
   End
   Begin VB.OptionButton optDMode 
      BackColor       =   &H00F5EADB&
      Caption         =   "View From To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   450
      TabIndex        =   0
      Top             =   735
      Value           =   -1  'True
      Width           =   2310
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6030
      Left            =   345
      OleObjectBlob   =   "frmCharter.frx":0000
      TabIndex        =   7
      Top             =   1950
      Width           =   10275
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Chart"
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
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   7935
   End
   Begin VB.Label lblEnd 
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
      Height          =   210
      Left            =   2595
      TabIndex        =   4
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lblStart 
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
      Height          =   210
      Left            =   660
      TabIndex        =   2
      Top             =   1065
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EADB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F5EADB&
      Height          =   1200
      Left            =   300
      Top             =   615
      Width           =   10305
   End
End
Attribute VB_Name = "frmCharter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dMode%

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not IsNumeric(txtStartYear) Then
        txtStartYear.SetFocus
        MsgBox "Invalid start year"
        Exit Sub
    End If
    
    If Not IsNumeric(txtEndYear) Then
        txtEndYear.SetFocus
        MsgBox "Invalid end year"
        Exit Sub
    End If
    
    If CDbl(txtStartYear) > CDbl(txtEndYear) Then
        Tag = txtEndYear
        txtEndYear = txtStartYear
        txtStartYear = Tag
    End If
    
    hGlass True
    RefreshChart
    
    If dMode = 0 Then
        lblTitle = "Comparative sales chart from  " & txtStartYear & " to " & txtEndYear
    Else
        lblTitle = "Comparative sales chart of " & txtStartYear & " VS " & txtEndYear
    End If
    
    hGlass False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    optDMode_Click 0
    CenterFrm Me
End Sub


Private Sub RefreshChart()
    Dim rsChartData As New ADODB.Recordset
    Dim rsChartDataYear As New ADODB.Recordset
    
    If dMode = 0 Then
        rsChartDataYear.Open "EXEC spChartDataYear " & txtStartYear & "," & txtEndYear, connRFM, adOpenStatic
        rsChartData.Open "EXEC spChartDataMonthly " & txtStartYear & "," & txtEndYear, connRFM, adOpenStatic
    Else
        rsChartDataYear.Open "EXEC spChartDataYearCompare " & txtStartYear & "," & txtEndYear, connRFM, adOpenStatic
        rsChartData.Open "EXEC spChartDataMonthlyCompare " & txtStartYear & "," & txtEndYear, connRFM, adOpenStatic
    End If
    
    If NoRecord(rsChartDataYear) Then
        MsgBox "No record for the years specified found", vbCritical
        
        ClearRS rsChartData
        ClearRS rsChartDataYear
        Exit Sub
    End If
    
    Dim nYear As Variant
    Dim nMonth As Integer
    Dim nRowCnt As Integer
    Dim index1 As Integer
    Dim index2 As Integer
    Dim index3 As Integer
    Dim index4 As Integer
    Dim index5 As Integer
    Dim nYearArray(50)
    Dim nDataArray() 'array With data from service details - the 2 is variable as it represents the No of years In the array
    Dim nYrscnt As Integer
    Dim bFirstTime As Boolean
    
    nYrscnt = rsChartDataYear.RecordCount
    
    ReDim nDataArray(rsChartDataYear.RecordCount, 12)
    
    MSChart1.Visible = True
    bFirstTime = True
    
    'temporary data to simulate data in an a
    '     rray (2 rows = 2 years and 12 months per
    '     year)
    
    Dim ictr%
    
    If Not NoRecord(rsChartDataYear) Then
        rsChartDataYear.MoveFirst
        While Not rsChartDataYear.EOF
            ictr% = ictr% + 1
            nYearArray(ictr) = rsChartDataYear("dYear")
            rsChartDataYear.MoveNext
        Wend
    End If
    
    ictr% = 0
    Dim ictr2%
    
    rsChartDataYear.MoveFirst
    While Not rsChartDataYear.EOF
        ictr% = ictr% + 1
        
        For ictr2% = 1 To 12
            rsChartData.Filter = "dYear = " & rsChartDataYear("dYear") & " AND dMonth=" & ictr2
            If Not rsChartData.EOF Then
                nDataArray(ictr, ictr2) = rsChartData("MonthlyTotal")
            Else
                nDataArray(ictr, ictr2) = 0
            End If
        Next ictr2
        rsChartDataYear.MoveNext
    Wend
    
    

    With MSChart1
        .chartType = VtChChartType2dBar
        .ColumnCount = 12 'create 12 collumns (within Each collumn) For each year


        For nYear = 1 To nYrscnt
            .RowCount = nYear 'create the main collumn For Each year (eg: If 2 years data Then 2 collumns will be created)


            For nMonth = 1 To 12
                .Column = nMonth 'will create 12 collumns (for the months) within Each main collumn
                .Row = nYear 'set the display row To the current row (year) that you are currently on
                .Data = nDataArray(nYear, nMonth)
                .RowLabel = nYearArray(nYear) 'set the bottom of Each collumn to display the current year (eg: collum1 will be 1998, collum2 will be 2000, etc)
                


                If bFirstTime = True Then 'set the collumn labels once only
                    'determine the current year and month,so
                    '     as to set the label description
                    .ColumnLabelIndex = nYear 'determine the collumn (year) that you are currently working On
                    .Column = nMonth 'determine the month that you are currently on


                    Select Case nMonth
                        Case 1
                        .ColumnLabel = "Jan"
                        Case 2
                        .ColumnLabel = "Feb"
                        Case 3
                        .ColumnLabel = "Mar"
                        Case 4
                        .ColumnLabel = "Apr"
                        Case 5
                        .ColumnLabel = "May"
                        Case 6
                        .ColumnLabel = "Jun"
                        Case 7
                        .ColumnLabel = "Jul"
                        Case 8
                        .ColumnLabel = "Aug"
                        Case 9
                        .ColumnLabel = "Sep"
                        Case 10
                        .ColumnLabel = "Oct"
                        Case 11
                        .ColumnLabel = "Nov"
                        Case 12
                        .ColumnLabel = "Dec"
                    End Select

            End If

        Next nMonth

        bFirstTime = False
    Next nYear

    
    ' shows the legends
    .ShowLegend = True
    
    End With
    
    ClearRS rsChartData
    ClearRS rsChartDataYear

End Sub




Private Sub optDMode_Click(Index As Integer)
    dMode = Index
    
    Select Case Index
        Case 0:
            lblStart = "Start Year to "
            lblEnd = "End Year"
        Case 1:
            lblStart = "This Year"
            lblEnd = "And This Year"
    End Select
    
    
End Sub


Private Sub txtEndYear_GotFocus()
    HighlightMe
End Sub

Private Sub txtStartYear_GotFocus()
    HighlightMe
End Sub
