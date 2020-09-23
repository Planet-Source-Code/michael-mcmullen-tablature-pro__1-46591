Attribute VB_Name = "Code"
Public strFont As String
Public intFontSize As Integer
Public strUserName As String
Public strRegCode As String
Public blnRegistered As Boolean

Public Sub WriteIni()
    Open App.Path & "\gtp.ini" For Output As #1
    Write #1, strFont, intFontSize, blnEditMode
    Close #1
End Sub

Public Sub ReadIni()
    On Error GoTo CreateNew
    
    Open App.Path & "\gtp.ini" For Input As #1
    Input #1, strFont, intFontSize, blnEditMode
    Close #1
    strUserName = GetSetting("Tablature Pro", "Main", "User Name", "")
    strRegCode = GetSetting("Tablature Pro", "Main", "Registration Code", "")
    Call CheckReg
    Exit Sub
    
CreateNew:
    Close #1
    strFont = "Lucida Console"
    intFontSize = 8
    blnEditMode = False
    blnRegistered = False
    Call WriteIni
End Sub

Public Sub CreateReg()
    
    strUserName = frmRegister.txtName.Text
    Dim strPass As String
    Dim strletter As String
    Dim i As Integer
    strPass = "TP"
    strletter = Left(strUserName, 1)
    
    For i = 1 To Len(strUserName)
        If strletter = "" Then
            Exit For
        End If
        strPass = strPass & (Str(Asc(strletter)) * Len(strUserName))
        strletter = Right(strUserName, Len(strUserName) - i)
    Next i
    strRegCode = strPass
End Sub

Public Sub CheckReg()
    
    Dim strPass As String
    Dim strletter As String
    Dim i As Integer
    strPass = "TP"
    strletter = Left(strUserName, 1)
    
    For i = 1 To Len(strUserName)
        If strletter = "" Then
            Exit For
        End If
        strPass = strPass & (Str(Asc(strletter)) * Len(strUserName))
        strletter = Right(strUserName, Len(strUserName) - i)
    Next i
    If strPass <> strRegCode Then
        blnRegistered = False
    Else
        blnRegistered = True
    End If
End Sub
