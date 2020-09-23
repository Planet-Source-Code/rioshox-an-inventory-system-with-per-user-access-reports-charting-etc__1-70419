Attribute VB_Name = "modADOTools"
Option Explicit
Public Function NoRecord(rs As ADODB.Recordset) As Boolean
    If rs.BOF And rs.EOF Then
        NoRecord = True
    Else
        NoRecord = False
    End If
     
End Function

Public Function Enquote(strName As String, Optional bAppendComma As Boolean) As String
    Dim strLine As String, strTemp As String
    Dim strSLine As String, strSTemp As String
    Dim DBlQuote As Integer, SglQuote As Integer, letterIdx As Integer
    
    'If Trim(strName) = "" Then
    '    Enquote = "''"
    '    Exit Function
    'End If
    
    strTemp = strName
    strLine = ""
    DBlQuote = InStr(1, strTemp, Chr(34))
    SglQuote = InStr(strTemp, Chr(39))
    
    If DBlQuote = 0 And SglQuote = 0 Then
        Enquote = "'" & Trim(strName) & "'"
        If bAppendComma Then Enquote = Enquote & ","
        Exit Function
    End If
    
    letterIdx = 0
    While DBlQuote > 0
        strLine = strLine & Left$(strTemp, DBlQuote) & Chr(34)
        strTemp = Mid$(strTemp, DBlQuote + 1, Len(strTemp) - 1)
        DBlQuote = InStr(CStr(strTemp), Chr(34))
        letterIdx = letterIdx + 1
    Wend
    
    
    strSTemp = strName
    
    letterIdx = 1
    SglQuote = InStr(strSTemp, Chr(39))
    
    If SglQuote = 0 Then
        Enquote = "'" & Trim(strSTemp) & "'"
        If bAppendComma Then Enquote = Enquote & ","
        Exit Function
    End If
    
    While SglQuote > 0
        strSLine = strSLine & Left$(strSTemp, SglQuote) & Chr(39)
        strSTemp = Mid$(strSTemp, SglQuote + 1, Len(strSTemp) - 1)
        SglQuote = InStr(CStr(strSTemp), Chr(39))
    Wend
    
    If SglQuote = 0 And strSLine <> "" Then
        strSLine = strSLine & strSTemp
    End If
    
    Enquote = "'" & strSLine & "'"
    If bAppendComma Then Enquote = Enquote & ","
End Function

Public Function EnNone(DNum As String, Optional bAppendComma As Boolean) As String
    DNum = Format(DNum, "GENERAL number")
    EnNone = CStr(DNum)
    If bAppendComma Then EnNone = EnNone & ","
End Function

Public Function EnPound(dDate As String, Optional bAppendComma As Boolean) As String
    EnPound = "#" & CStr(dDate) & "#"
    If bAppendComma Then EnPound = EnPound & ","
End Function

Public Sub MoveCursor(rs As ADODB.Recordset, Index As Integer)
    On Error Resume Next
    With rs
        If .BOF And .EOF Then
            'MsgBox "No record yet.", vbCritical, "Cannot navigate"
            Exit Sub
        End If
        Select Case Index
            Case 0:
                .MoveFirst
            Case 1:
                .MovePrevious
                If .BOF Then .MoveFirst
                While .Status = adRecDBDeleted And Not .BOF
                    .MovePrevious
                    If .BOF Then .MoveFirst
                Wend
                
            Case 2:
                .MoveNext
                If .EOF Then .MoveLast
                While .Status = adRecDBDeleted And Not .EOF
                    .MoveNext
                    If .EOF Then .MoveLast
                Wend
            Case 3:
                .MoveLast
        End Select
    End With
    
End Sub

Public Sub PopulateCboBox(rs As ADODB.Recordset, dFld As String, CboBox As ComboBox, Optional None As Boolean)
    Dim rsCloner As ADODB.Recordset
    Set rsCloner = rs.Clone
    
    rsCloner.Requery
    CboBox.Clear
    With rsCloner
        .MoveFirst
        While Not .EOF
            CboBox.AddItem rsCloner.Fields(dFld) & ""
            .MoveNext
        Wend
    End With
    
    If None Then CboBox.AddItem "None"
    
    CboBox.ListIndex = 0
    rsCloner.Close
    Set rsCloner = Nothing
End Sub

Public Sub CompactDB(DSource As String, dTemp As String)
    'Dim JRO As JRO.JetEngine
    
    
    'Set JRO = New JRO.JetEngine
    
    'JRO.CompactDatabase _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & App.Path & "\Backup\Payroll.mdb;" & _
        "Provider=Microsoft.Jet.OLEDB.4.0;", _
        "Data Source=" & App.Path & "\Backup\" & CStr(Format(Date, "mmddyyyy")) & CStr(Format(Time, "hhmmss")) & ".mdb;" & _
        "Jet OLEDB:Engine Type=4;"
End Sub

Public Sub ClearRS(rs As ADODB.Recordset)
    On Error Resume Next
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Public Sub LockFields(dfrm As Form, bMode As Boolean)
    Dim ctrl
    
    For Each ctrl In dfrm.Controls
        
    Next
End Sub

