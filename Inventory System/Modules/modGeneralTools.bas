Attribute VB_Name = "modGeneralTools"
Option Explicit

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SC_MAXIMIZE = &HF030&
Private Const WM_SYSCOMMAND = &H112
Private Const WM_CLOSE = &H10

Public Const CB_FINDSTRING = &H14C

Private bNoLookUp As Boolean

Public Sub HighlightMe()
    SendKeys "{HOME}" + "+{END}"
End Sub

'************************************************************
'*                 AutoComplete ComboBox                    *
'*  Creates a ComboBox that supports automatic completion.  *
'*                                                          *
'*                                                          *
'* Usage:                                                   *
'*   1.  Call ComboChange() in the Change event of the      *
'*         combo box.  Pass in the combo box                *
'*   2.  Call ComboKeyDown() in the Keydown event of the    *
'*         combo box.  Pass in as a parameter the keycode   *
'*         from the original combobox KeyDown event.        *
'************************************************************

Public Sub ComboChange(Combo As ComboBox)
Const Location = "ComboChange"
Dim pos As Long

On Error GoTo MyError

  If bNoLookUp = True Then
    bNoLookUp = False
    Exit Sub
  End If
  
  pos = Combo.SelStart
  Combo.ListIndex = SendMessage(Combo.hWnd, CB_FINDSTRING, -1, ByVal CStr(Combo.Text))
  If Combo.ListIndex = -1 Then
    pos = Combo.SelStart
  Else
    Combo.SelStart = pos
    Combo.SelLength = Len(Combo.Text) - pos
  End If
  
Exit Sub

MyError:
    Debug.Print "Error: " & Err.Description & ", " & Err.Number & " in module '" & Location & "'"
    Resume Next
End Sub
Public Sub ComboKeyDown(KeyCode As Integer)
Const Location = "ComboKeyDown"

On Error GoTo MyError

  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    bNoLookUp = True
  End If

Exit Sub

MyError:
    Debug.Print "Error: " & Err.Description & ", " & Err.Number & " in modul '" & Location & "'"
    Resume Next
End Sub

Public Sub ShowReport(strReportName As String, Optional strFilterCriteria As String, Optional intViewMode As Integer = acViewPreview, Optional strSource As String)
    Dim objAccess As New Access.Application
    Dim mhWndAccess As Long
    
    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
        
    objAccess.DoCmd.RunCommand acCmdAppMaximize
    objAccess.OpenCurrentDatabase App.Path & "\RFM.mdb"
    
    If Err.Number <> 0 Then
        objAccess.DoCmd.Quit acQuitSaveNone
        Set objAccess = Nothing
        MsgBox "Errors occured. Cannot continue printing" & vbCrLf & Err.Description, vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If strFilterCriteria & "" = "" Then
        objAccess.DoCmd.OpenReport strReportName, intViewMode
    Else
        objAccess.DoCmd.OpenReport strReportName, intViewMode, strSource, strFilterCriteria
    End If
       
    
    objAccess.DoCmd.SelectObject acReport, strReportName
    objAccess.Visible = True
    
    'Store so we can close the window when we close
    mhWndAccess = objAccess.hWndAccessApp
    
    'Maximize Access
    SendMessage objAccess.hWndAccessApp, WM_SYSCOMMAND, (SC_MAXIMIZE And &HFFF0), 0&
    
    DoEvents

    'Set objAccess = Nothing
    Screen.MousePointer = vbDefault

End Sub

Public Function TwoDecimals(nValue As Double) As String
    If Not IsNumeric(nValue) Then
        TwoDecimals = nValue
    Else
        TwoDecimals = Format(nValue, "Standard")
    End If
End Function

'Public Sub CenterFrm(dFrm As Form)
'    dFrm.Left = (MainForm.Width - frmLeftFrame.Width) / 2'
'End Sub

Public Sub hGlass(bMode As Boolean)
    If bMode Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbDefault
    End If
End Sub

Public Function Encrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = False) As String
    Dim i
    On Error GoTo ErrorHandler
    Dim Char As String
    Encrypt = ""
    
    For i = 1 To Len(StringToEncrypt)
        Char = Asc(Mid(StringToEncrypt, i, 1))
        Encrypt = Encrypt & Len(Char) & Char
    Next i
    
    If AlphaEncoding Then
    
        StringToEncrypt = Encrypt
        Encrypt = ""
        
        For i = 1 To Len(StringToEncrypt)
            Encrypt = Encrypt & Chr(Mid(StringToEncrypt, i, 1) + 147)
        Next i
        
    End If
    Exit Function
ErrorHandler:
    Encrypt = "Error encrypting string"
End Function

Public Function Decrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = False) As String
    Dim i
    On Error GoTo ErrorHandler
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    
    If AlphaDecoding Then
    
        Decrypt = StringToDecrypt
        StringToDecrypt = ""
        
        For i = 1 To Len(Decrypt)
            StringToDecrypt = StringToDecrypt & (Asc(Mid(Decrypt, i, 1)) - 147)
        Next i
        
    End If
    
    Decrypt = ""
    
    Do
    
        CharPos = Left(StringToDecrypt, 1)
        StringToDecrypt = Mid(StringToDecrypt, 2)
        CharCode = Left(StringToDecrypt, CharPos)
        StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
        Decrypt = Decrypt & Chr(CharCode)
        
    Loop Until StringToDecrypt = ""
    Exit Function
ErrorHandler:
    Decrypt = "Error decrypting string"
End Function

Public Sub CenterFrm(dfrm As Form)
    dfrm.Top = (Screen.Height - dfrm.Height) / 2
    dfrm.Left = ((Screen.Width - dfrm.Width) / 2) + 1200
End Sub

Public Function TrimAll(TheString As String) As String
    Dim i%
    Dim LastChar As String
    Dim NextChar As String
    LastChar = Left(TheString, 1)
    TrimAll = LastChar

    For i = 2 To Len(TheString)
    NextChar = Mid(TheString, i, 1)


    If NextChar = " " And LastChar = " " Then
    Else
        TrimAll = TrimAll & NextChar

End If

LastChar = NextChar
Next i

End Function

Public Function Propercase(TheString As String) As String
    Dim i%
    
    TheString = TrimAll(TheString)
    Propercase = UCase(Left(TheString, 1))


    For i = 2 To Len(TheString)


        If Mid(TheString, i - 1, 1) = " " Then
            Propercase = Propercase & UCase(Mid(TheString, i, 1))
        Else
            Propercase = Propercase & LCase(Mid(TheString, i, 1))
        End If

    Next i
    
End Function

Public Function IsLoaded(ByVal pObjForm As Form) As Boolean

    Dim tmpForm As Form


    For Each tmpForm In Forms

        If tmpForm Is pObjForm Then
            IsLoaded = True
            Exit For
        End If

    Next

End Function


Public Sub ShrinkForm(dfrm As Form)
    Dim num%, i%
    
    num = 1


    For i = 0 To dfrm.Height
        
        
        dfrm.Height = dfrm.Height - num


        DoEvents
        Next i


        For i = 0 To dfrm.Width
            dfrm.Width = dfrm.Width - num


            DoEvents
            Next i


            For i = 0 To dfrm.Left
                dfrm.Left = dfrm.Left - num
                num = num + 1


                DoEvents
                Next i
    
End Sub


Public Sub ResizeGrid(dGrid As MSFlexGrid, OrigWidth As Integer, dSize As Integer)
    If dSize < 10 Then
        dGrid.Width = OrigWidth
    Else
        dGrid.Width = OrigWidth + 250
    End If
End Sub


Public Sub soundIt(whatSound As String)

    Dim dumInt  As Variant
    Dim soundtoPlay$
    
    soundtoPlay = App.Path & "\Sounds\" & whatSound & ".WAV"
    dumInt = PlaySound(soundtoPlay, 10, 1&)
    
End Sub

