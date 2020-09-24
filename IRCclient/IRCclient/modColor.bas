Attribute VB_Name = "modColor"
Option Explicit

Public Type udtColorVariables
    colorAction     As String * 6
    colorCTCP       As String * 6
    colorJoin       As String * 6
    colorPart       As String * 6
    colorKick       As String * 6
    colorQuit       As String * 6
    colorMode       As String * 6
    colorNotice     As String * 6
    colorOwn        As String * 6
    colorNick       As String * 6
    colorUser       As String * 6
    colorInvite     As String * 6
    colorTopic      As String * 6
    colorWhois      As String * 6
    colorChat       As String * 6
    colorOther      As String * 6
    colorListText   As String * 6
    colorEditText   As String * 6
    colorEdit       As String * 6
    colorFrame      As String * 6
    colorList       As String * 6
    DoColor         As String * 6
End Type


Public udtColor As udtColorVariables

Public Sub DoColor(RTF As RichTextBox, a As String)

Dim b As String
Dim n As Integer
Dim n2 As Integer
Dim fgcolor As Integer
Dim bgcolor As Integer
Dim savefg As Integer
Dim savebg As Integer
Dim bReverse As Boolean

    Dim Color(0 To 15) As Long
    Color(0) = vbWhite 'white
    Color(1) = vbBlack 'black
    Color(2) = RGB(0, 0, 140) 'dark blue
    Color(3) = RGB(0, 140, 0) 'dark green
    Color(4) = vbRed 'red
    Color(5) = RGB(110, 65, 0) 'brown
    Color(6) = RGB(140, 0, 140) 'purple
    Color(7) = RGB(248, 146, 0) 'orange
    Color(8) = RGB(255, 255, 0) 'yellow
    Color(9) = vbGreen 'light green
    Color(10) = RGB(0, 140, 140) 'dark blue green
    Color(11) = RGB(0, 255, 255) 'light blue green
    Color(12) = vbBlue 'light blue
    Color(13) = vbMagenta 'magenta
    Color(14) = RGB(140, 140, 140) 'grey
    Color(15) = RGB(200, 200, 200) 'light grey
    
RTF.SelBold = False
RTF.SelUnderline = False

fgcolor = 1
bgcolor = 0
bReverse = False
savefg = fgcolor
savebg = bgcolor

For n = 1 To Len(a)
   b = Mid(a, n, 1)
   If b = Chr(3) Then
    'Parse Colours
    If IsNumeric(Mid(a, n + 1, 1)) Then
       If IsNumeric(Mid(a, n + 2, 1)) Then
        If Mid(a, n + 3, 1) = "," Then
            If IsNumeric(Mid(a, n + 4, 1)) Then
                If IsNumeric(Mid(a, n + 5, 1)) Then
                    '@##,##
                    fgcolor = CInt(Mid(a, n + 1, 2))
                    bgcolor = CInt(Mid(a, n + 4, 2))
                    n = n + 5
                Else
                    '@##,#
                    fgcolor = CInt(Mid(a, n + 1, 2))
                    bgcolor = CInt(Mid(a, n + 4, 1))
                    n = n + 4
                End If
            Else
                '@##,
                fgcolor = CInt(Mid(a, n + 1, 2))
                n = n + 3
            End If
        Else
            '@##
            fgcolor = CInt(Mid(a, n + 1, 2))
            n = n + 2
        End If
           ElseIf Mid(a, n + 2, 1) = "," Then
        If IsNumeric(Mid(a, n + 3, 1)) Then
            If IsNumeric(Mid(a, n + 4, 1)) Then
                '@#,##
                fgcolor = CInt(Mid(a, n + 1, 1))
                bgcolor = CInt(Mid(a, n + 3, 2))
                n = n + 4
            Else
                '@#,#
                fgcolor = CInt(Mid(a, n + 1, 1))
                bgcolor = CInt(Mid(a, n + 3, 1))
                n = n + 3
            End If
        Else
            '@#,
            fgcolor = CInt(Mid(a, n + 1, 1))
            n = n + 2
        End If
           Else
        '@#
        fgcolor = CInt(Mid(a, n + 1, 1))
        n = n + 1
       End If
       If fgcolor > 15 Then
           fgcolor = 1
       End If

       If bgcolor > 15 Then
           bgcolor = 0
       End If
       RTF.SelColor = Color(fgcolor)
       'RTF.FontBackColour = Color(bgcolor)
        Else
           RTF.SelColor = Color(1)
           'RTF.FontBackColour = Color(0)
        End If
   ElseIf b = Chr(2) Then
    RTF.SelBold = Not (RTF.SelBold)
'   if bBold then'
'       'Turn Bold off
'       bBold = False
'       RTF.FontBold = False
'   else
'       'Turn Bold on
'       bBold = True
'       RTF.FontBold = True
'   endif
   ElseIf b = Chr(31) Then
        RTF.SelUnderline = Not (RTF.SelUnderline)
'   if bUnderline then
'       'Turn underline off
'       bUnderline = False
'       RTF.FontUnderline = False
'   else
'       'Turn underline on
'       bUnderline = True
'       RTF.FontUnderline = True
'   endif
   ElseIf b = Chr(22) Then
    'Reverse Foreground / Background colors
'    n2 = bgcolor
'    bgcolor = fgcolor
'    fgcolor = n2
'    RTF.FontColour = color(fgcolor)
'    RTF.FontBackColour = color(bgcolor)
    'Set the colors to the reverse standard colour set.
    If bReverse Then
        bReverse = False
        fgcolor = savefg
        bgcolor = savebg
    Else
        bReverse = True
        savefg = fgcolor
        savebg = bgcolor
        fgcolor = 0
        bgcolor = 1
        
    End If
    RTF.SelColor = Color(fgcolor)
    'RTF.FontBackColour = Color(bgcolor)
    
   Else
    RTF.SelText = b
   End If
Next n

    
'    For i = 1 To Len(strColor)
'        RTF.InsertContents SF_TEXT, Asc(Mid(strColor, i, 1)) & "|"
'    Next i
    
    RTF.SelColor = Color(1)
    'RTF.FontBackColour = Color(0)
    RTF.SelBold = False
    RTF.SelUnderline = False

'    RTF.InsertContents SF_TEXT, vbCrLf
End Sub


Public Sub LoadColor()

    Dim strApp As String
    Dim intFile As Integer
    intFile = FreeFile
    strApp = App.Path & "\Color.nvf"
    
    'Load color from file
    Open strApp For Random Access Read As #intFile Len = Len(udtColor)
        Get #intFile, , udtColor
    Close #intFile
    
    'If there are no color setting, then give default color
    With udtColor
    'If .colorAction = String(6, Chr(0)) Then
            .colorAction = "9C009C"
            .colorCTCP = "FF0000"
            .colorJoin = "009300"
            .colorPart = "009300"
            .colorKick = "009300"
            .colorQuit = "00007F"
            .colorMode = "009300"
            .colorNotice = "7F0000"
            .colorOwn = "000000"
            .colorNick = "000000"
            .colorUser = "000000"
            .colorInvite = "009300"
            .colorTopic = "009300"
            .colorWhois = "000000"
            .colorChat = "000000"
            .colorOther = "000000"
            .colorListText = "000000"
            .colorEditText = "000000"
            .colorEdit = "FFFFFF"
            .colorFrame = "FFFFFF"
            .colorList = "FFFFFF"
        'End If
    End With
    
End Sub

Public Sub SaveColor()
    'Save current color setting to file
    Dim strApp As String
    Dim intFile As Integer
    intFile = FreeFile
    strApp = App.Path & "\Color.nvf"
    
    Open strApp For Random Access Write As #intFile Len = Len(udtColor)
        Put #intFile, , udtColor
    Close #intFile
End Sub
Public Function RGBtoHEX(RGB) As String
    'Convert rgb format to hex
    Dim strMsg As String
    Dim intCounter As Integer
    strMsg = Hex(RGB)
    intCounter = Len(strMsg)
    
    Select Case intCounter
        Case 1
            strMsg = String(5, "0") & strMsg
        Case 2
            strMsg = String(4, "0") & strMsg
        Case 3
            strMsg = String(3, "0") & strMsg
        Case 4
            strMsg = String(2, "0") & strMsg
        Case 5
            strMsg = String(1, "0") & strMsg
    End Select
    RGBtoHEX = strMsg
End Function

Public Sub ChangeObjectColor(ctrlObject As Control, strColor As String, intNum As Integer)
    'This sub is universal to change an object color.

    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    
    intRed = Val("&H" & Right(strColor, 2))
    intGreen = Val("&H" & Mid(strColor, 3, 2))
    intBlue = Val("&H" & Left(strColor, 2))
    
    Select Case intNum
        Case 1
            ctrlObject.BackColor = RGB(intRed, intGreen, intBlue)
        Case 2
            ctrlObject.ForeColor = RGB(intRed, intGreen, intBlue)
        Case 3
            ctrlObject.SelColor = RGB(intRed, intGreen, intBlue)
    End Select
End Sub

