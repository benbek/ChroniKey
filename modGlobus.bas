Attribute VB_Name = "modGlobus"
' CAUTION: ChroniKey does not check whether the films' names are longer or shorter than what Chronica accepts
' "\" = Integer Division; "/" = Real Division
Option Explicit
Option Base 1
Public Const MaxTimes As Integer = 10 ',MaxCinemas As Integer = 150, MaxMovies As Integer = 1000
Public Const AuthEMail As String = "chronikey@freaktalk.com", ChroniKeySite As String = "chronikey.freaktalk.com"
Public StartDate As Date, EndDate As Date, Cinemas() As String, intCinemas As Variant, _
       Movies() As String, intMovies As Integer, Data() As String, intMaxCine As _
       Integer, intMaxMov As Integer, intNumCine As Integer, intNumMov As Integer, _
       DataFile As String, DateFormat As String, CinemaObtainURL As String, MovieObtainURL As String
Public Running9x As Boolean
Public BeingStopped As Boolean
    
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    czCSDVersion As String * 128
End Type

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158

'Taken from http://www.dr-vb.co.il/dws.php?id=281
'Written by Yaniv Drukman
Public Const SHACF_DEFAULT = &H0
Public Const SHACF_FILESYSTEM = &H1
Public Const SHACF_FILESYS_ONLY = &H10
Public Const SHACF_FILEALL = (SHACF_FILESYSTEM Or SHACF_FILESYS_ONLY)
Public Const SHACF_URLHISTORY = &H2
Public Const SHACF_URLMRU = &H4
Public Const SHACF_URLALL = (SHACF_URLHISTORY Or SHACF_URLMRU)
Public Const SHACF_USETAB = &H8
Declare Sub SHAutoComplete Lib "shlwapi.dll" (ByVal hwndEdit As Long, ByVal dwFlags As Long)

Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

'Dimension a variable as the user-defined data type.
'Public nID As NOTIFYICONDATA

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'For displaying text-boxes' balloons
'Taken from http://vbnet.mvps.org/index.html?code/textapi/showballoontiptext.htm
Public Const ECM_FIRST As Long = &H1500
Public Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Public Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)

Public Type EDITBALLOONTIP
   cbStruct As Long
   pszTitle As String
   pszText As String
   ttiIcon As Long
End Type

Public Enum TextboxBalloonTipIconConstants
   TTI_NONE = 0
   TTI_INFO = 1
   TTI_WARNING = 2
   TTI_ERROR = 3
End Enum

'For displaying combo-boxes' balloons
'Taken from http://vbnet.mvps.org/index.html?code/textapi/showballoontiptext.htm
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton As Long
   hwndCombo As Long
   hwndEdit As Long
   hwndList As Long
End Type

Private Declare Function GetComboBoxInfo Lib "user32" _
  (ByVal hwndCombo As Long, _
   CBInfo As COMBOBOXINFO) As Long

'For visual XP styles (link to ComCtl32.dll)
'Taken from http://www.vbaccelerator.com/home/VB/Code/Libraries/XP_Visual_Styles/Using_XP_Visual_Styles_in_VB/article.asp
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
   
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Function GetComboEditHWND(ctl As ComboBox) As Long
   Dim CBI As COMBOBOXINFO

   CBI.cbSize = Len(CBI)
   Call GetComboBoxInfo(ctl.hwnd, CBI)
   GetComboEditHWND = CBI.hwndEdit
End Function

Public Sub ShowTextboxBalloonTip(m_hWnd As Long, ByVal m_sText$, ByVal m_sTitle$, m_eIcon%)
   Dim ebt As EDITBALLOONTIP
   
   With ebt
        ebt.cbStruct = Len(ebt)
        ebt.pszText = StrConv(m_sText, vbUnicode) ' Text to show
        ebt.pszTitle = StrConv(m_sTitle, vbUnicode) ' Title to show
        ebt.ttiIcon = m_eIcon 'One of the balloon tip icons
   End With
   
   Call SendMessage(m_hWnd, EM_SHOWBALLOONTIP, 0, ebt)
End Sub

Public Sub HideTextboxBalloonTip(m_hWnd As Long)
   Dim lR As Long
   
   Call SendMessage(m_hWnd, EM_HIDEBALLOONTIP, 0, 0)
End Sub

Public Sub FindInCombo(ByVal strSearch$, cbxComb As ComboBox)
    Dim intI%, Found As Boolean
    
    intI = -1
    Found = False
    If strSearch = "" Then Exit Sub
    Do While intI + 1 < cbxComb.ListCount And Not Found
        intI = intI + 1
        Found = cbxComb.ItemData(intI) = strSearch
    Loop
    If Found Then cbxComb.ListIndex = intI
End Sub

Public Function HTMLessString(ByVal strData$) As String
    'Cleans the specified string line
    'Returns the string without any html tags
    Dim intiStart, intiEnd As Integer
    Dim strFinal As String
    
    intiStart = InStr(1, strData, "<", vbTextCompare)
    intiEnd = InStr(intiStart + 1, strData, ">", vbTextCompare)
    While intiStart > 0 And intiEnd > 0
        strFinal = Left(strData, intiStart - 1)
        strFinal = strFinal & Right(strData, Len(strData) - intiEnd)
        strData = strFinal
        intiStart = InStr(1, strData, "<", vbTextCompare)
        intiEnd = InStr(intiStart + 1, strData, ">", vbTextCompare)
    Wend
    intiStart = InStr(1, strData, "&nbsp;", vbTextCompare)
    While intiStart > 0
        strFinal = Left(strData, intiStart - 1)
        strFinal = strFinal & " " & Mid(strData, intiStart + Len("&nbsp;"))
        strData = strFinal
        intiStart = InStr(1, strData, "&nbsp;", vbTextCompare)
    Wend
    intiStart = InStr(1, strData, "&quot;", vbTextCompare)
    While intiStart > 0
        strFinal = Left(strData, intiStart - 1)
        strFinal = strFinal & """" & Mid(strData, intiStart + Len("&quot;"))
        strData = strFinal
        intiStart = InStr(1, strData, "&quot;", vbTextCompare)
    Wend
    HTMLessString = Trim(strData)
End Function

Public Function NoQuotesString(ByVal strData$) As String
    'Cleans the specified string line
    'Returns the string without any quotation marks
    Dim intiStart%
    Dim strFinal As String
    
    intiStart = InStr(1, strData, """", vbTextCompare)
    While intiStart > 0
        strFinal = Left(strData, intiStart - 1)
        strFinal = strFinal & Mid(strData, intiStart + 1)
        strData = strFinal
        intiStart = InStr(1, strData, """", vbTextCompare)
    Wend
    NoQuotesString = Trim(strData)
End Function

Public Function CommaString(ByVal strData$) As String
    'Cleans the specified string line
    'Returns the string with the commas replaced with spaces
    Dim intiStart%
    Dim strFinal As String
    
    intiStart = InStr(1, strData, ",", vbTextCompare)
    While intiStart > 0
        strFinal = Left(strData, intiStart - 1)
        strFinal = strFinal & " " & Mid(strData, intiStart + 1)
        strData = strFinal
        intiStart = InStr(1, strData, ",", vbTextCompare)
    Wend
    CommaString = Trim(strData)
End Function

Public Function DashString(ByVal strData$) As String
    'Cleans the specified string line
    'Returns the string with the extended-dashes replaced with regular ones
    Dim intiStart%
    Dim strFinal As String
    
    intiStart = InStr(1, strData, "–", vbTextCompare)
    While intiStart > 0
        strFinal = Left(strData, intiStart - 1)
        strFinal = strFinal & "-" & Mid(strData, intiStart + 1)
        strData = strFinal
        intiStart = InStr(1, strData, "–", vbTextCompare)
    Wend
    DashString = Trim(strData)
End Function

Public Function TabString(ByVal strData$) As String
    'Cleans the specified string line
    'Returns the string with the the tabs replaced with lots of spaces
    Dim intiStart%
    Dim strFinal As String
    
    intiStart = InStr(1, strData, vbTab, vbTextCompare)
    While intiStart > 0
        strFinal = Left(strData, intiStart - 1)
        strFinal = strFinal & "   " & Mid(strData, intiStart + 1)
        strData = strFinal
        intiStart = InStr(1, strData, vbTab, vbTextCompare)
    Wend
    TabString = Trim(strData)
End Function

Public Function QuotLessString(ByVal strData$) As String
    'Cleans the specified string line
    'Returns the string without any quotation marks
    'Specially written for Globus Group
    
    strData = Trim(strData)
    strData = Mid(strData, 2, Len(strData) - 1)
    QuotLessString = strData
End Function

Public Function NoDoubleSpaceString(ByVal strData$) As String
    'Cleans the specified string line
    'Returns the string without any double spaces ("  ") inside the string
    Dim intiStart%
    Dim strFinal As String
    
    intiStart = InStr(1, strData, "  ", vbTextCompare)
    While intiStart > 0
        strFinal = Left(strData, intiStart - 1)
        strFinal = strFinal & Mid(strData, intiStart + 1)
        strData = strFinal
        intiStart = InStr(1, strData, "  ", vbTextCompare)
    Wend
    NoDoubleSpaceString = Trim(strData)
End Function

Public Function CleanString(ByVal strData$) As String
    'Cleans the specified string line
    'Returns the string but with no characters before "=" (including "=") and from any ";"
    Dim intInteger
    
    intInteger = InStr(1, strData, "=", vbTextCompare)
    If intInteger > 0 Then strData = Mid(strData, intInteger + 1)
    intInteger = InStr(1, strData, ";", vbTextCompare)
    If intInteger > 0 Then strData = Left(strData, intInteger - 1)
    intInteger = InStr(1, strData, Chr(9), vbTextCompare) 'Search for Tabs
    If intInteger > 0 Then strData = Left(strData, intInteger - 1)
    
    CleanString = Trim(strData)
End Function

'Public Function RemFirChar(ByVal strData$) As String
    'Cleans the specified string line by removing any even character
'    Dim intInteger%, strFinal$
'
'    For intInteger = 1 To Len(strData) Step 2
'        strFinal = strFinal & Mid(strData, intInteger, 1)
'    Next intInteger
'    strData = strFinal
'
'    intInteger = InStr(1, strData, "‎", vbTextCompare)
'    While intInteger > 0
'        strFinal = Left(strData, intInteger - 1)
'        strFinal = strFinal & Mid(strData, intInteger + 1)
'        strData = strFinal
'        intInteger = InStr(1, strData, "‎", vbTextCompare)
'    Wend
'    RemFirChar = Trim(strData)
'End Function

Public Function RemHenChar(ByVal strData$) As String
    'Cleans the specified string line by removing the "‎" character
    Dim intInteger%, strFinal$, Spacer%, Mone%
    
    'For intInteger = 2 To Len(strData) Step 2
    '    strFinal = strFinal & Mid(strData, intInteger, 1)
    'Next intInteger
    'strData = strFinal
    
    'strData = CropString(strData, Chr(227))
    strData = CropString(strData, Chr(5))
    strData = CropString(strData, Chr(211))
    
    intInteger = InStr(1, strData, "‎", vbTextCompare)
    If intInteger = 0 Then intInteger = InStr(1, strData, Chr(14), vbTextCompare)
    While intInteger > 0
        strFinal = Left(strData, intInteger - 1)
        strFinal = strFinal & Mid(strData, intInteger + 1)
        strData = strFinal
        intInteger = InStr(1, strData, "‎", vbTextCompare)
        If intInteger = 0 Then intInteger = InStr(1, strData, Chr(14), vbTextCompare)
    Wend
    
    'Handling the extra space that exists in the newer Rav-Hen boards
    'strData = Trim(strData)
    'intInteger = 1
    'While Mid(strData, intInteger, 1) = Chr(32) And intInteger <= Len(strData)
    '    intInteger = intInteger + 1
    'Wend
    ''intInteger contains the first spot in the string that contains data (not spaces)
    'Spacer = 1 + (intInteger Mod 2)
    '
    ''Count the number of times there is a space in every second spot in the string
    'Mone = 0
    'For intInteger = Spacer To Len(strData) Step 2
    '    If Mid(strData, intInteger, 1) = Chr(32) Then Mone = Mone + 1
    'Next intInteger
    '
    'If Mone = Len(strData) \ 2 Then 'Delete every second space
    '    For intInteger = Spacer To Len(strData) Step 2
    '        strFinal = Left(strData, intInteger - 1)
    '        strFinal = strFinal & Mid(strData, intInteger + 1)
    '        strData = strFinal
    '    Next intInteger
    'End If
    
    RemHenChar = Trim(strData)
End Function

Public Function CropString(ByVal strData$, ByVal Remove$) As String
    'Cleans the specified string line
    'Returns the string without the wanted-to-be-removed part
    Dim intiStart%
    Dim strFinal As String
    
    intiStart = InStr(1, strData, Remove, vbTextCompare)
    While intiStart > 0
        strFinal = Left(strData, intiStart - 1)
        strFinal = strFinal & Mid(strData, intiStart + Len(Remove))
        strData = strFinal
        intiStart = InStr(1, strData, Remove, vbTextCompare)
    Wend
    CropString = Trim(strData)
End Function

Public Sub RemoveHour(str_Data$, strRemove$)
    'Removes the hours written in strRemove from str_Data
    Dim iI1%, iI2%, toRemove() As String
    
    'iI1 = 1
    'iI2 = InStr(iI1, strRemove, " ", vbTextCompare)
    'If Len(strRemove) >= 5 And iI2 = 0 Then iI2 = -999
    'Do While iI2 <> 0
    '    If iI2 = -999 Then iI2 = 0
    '    str_Data = CropString(str_Data, Mid(strRemove, iI1, Len(strRemove) - iI2), True)
    '    iI1 = iI2 + 1
    '    iI2 = InStr(iI1, strRemove, " ", vbTextCompare)
    'Loop
    'str_Data = NoDoubleSpaceString(str_Data)
    toRemove = Split(strRemove, " ")
    For iI1 = 0 To UBound(toRemove)
        str_Data = CropString(str_Data, toRemove(iI1))
    Next iI1
    str_Data = NoDoubleSpaceString(str_Data)
End Sub

Public Function ReverseText(ByVal strData$) As String
    'Reverses the given text
    ' Hello -> olleH
    Dim intInteger%, strFinal$, strChar$
    
    For intInteger = Len(strData) To 1 Step -1
        strChar = Mid(strData, intInteger, 1)
        If (strChar >= Chr(48) And strChar <= Chr(57)) Then 'Don't reverse numbers
            strFinal = strChar & strFinal
        'ElseIf (strChar >= Chr(65) And strChar <= Chr(90)) Then 'Don't reverse English
        '    strFinal = strChar & strFinal
        Else
            strFinal = strFinal & strChar
        End If
    Next intInteger
    strData = strFinal
    ReverseText = strData
End Function

Public Function CutEnglish(ByVal strData$) As String
    'Crops all of the capital english letters and numbers from the string
    Dim intInteger%
    Dim strFinal As String
    
    For intInteger = 1 To Len(strData)
        If (Mid(strData, intInteger, 1) >= Chr(65) And Mid(strData, intInteger, 1) <= Chr(90)) Or (Mid(strData, intInteger, 1) >= Chr(48) And Mid(strData, intInteger, 1) <= Chr(57)) Or (Mid(strData, intInteger, 1) >= Chr(97) And Mid(strData, intInteger, 1) <= Chr(122)) Then
           strFinal = Left(strData, intInteger - 1)
           strData = strFinal & Mid(strData, intInteger + 1)
           intInteger = intInteger - 1
           If intInteger = Len(strData) Then Exit For
        End If
    Next intInteger
    CutEnglish = Trim(strData)
End Function

Public Function LeftRight(ByVal strData$, ByVal FirstMovieInRow As Boolean) As String
    
    On Error Resume Next
    
    If InStr(1, strData, "|", vbTextCompare) = 0 Then
        LeftRight = strData
        Exit Function
    End If
    
    'If the string is Hebrew, make sure that what the cutting will fit the
    'right-to-left reading order
    If (Mid(Trim(strData), 1, 1) >= Chr(48) And Mid(Trim(strData), 1, 1) <= Chr(57)) Or Mid(strData, 1, 1) = Chr(44) Or Mid(Trim(strData), 1, 1) = "|" Or (HaveHour(strData) And (InStr(1, strData, "מוגבל מגיל", vbTextCompare) > 0 Or InStr(1, strData, "ליגמ לבגומ", vbTextCompare) > 0 Or InStr(1, strData, "בכורה ארצית", vbTextCompare) > 0 Or InStr(1, strData, "תיצרא הרוכב", vbTextCompare) > 0 Or InStr(1, strData, "טרום בכורה", vbTextCompare) > 0 Or InStr(1, strData, "הרוכב םורט", vbTextCompare) > 0) Or InStr(1, strData, "בכורה עולמית", vbTextCompare) > 0 Or InStr(1, strData, "תימלוע הרוכב", vbTextCompare) > 0) Then
        FirstMovieInRow = Not FirstMovieInRow
    End If
    'If (Mid(Trim(strData), 1, 1) >= Chr(224) And Mid(Trim(strData), 1, 1) <= Chr(250)) Or (Mid(Trim(strData), 1, 1) >= Chr(48) And Mid(Trim(strData), 1, 1) <= Chr(57)) Then
    '    FirstMovieInRow = Not FirstMovieInRow
    'End If
    
    Dim sStr() As String
    sStr() = Split(strData, "|")
    
    If FirstMovieInRow Then
    '    LeftRight = Trim(Mid(strData, 1, InStr(1, strData, "|", vbTextCompare) - 1))
         LeftRight = Trim(sStr(0))
    Else
    '    LeftRight = Trim(Mid(strData, InStr(1, strData, "|", vbTextCompare) + 1))
         LeftRight = Trim(sStr(1))
    End If

End Function

Public Function HaveHour(ByVal str_Data$) As Boolean
   'Before adding the new "repair-it-so-VB-could-detect-it" featured in WhereHour,
   'one must check whether it's necessary and if Rav-Hen support of ChroniKey isn't
   'built on this result (w/o the feature)
   On Error Resume Next
   Dim intIndex%, IsHour As Boolean, TempDate As Date

   If Len(str_Data) = 0 Then
       HaveHour = False
       Exit Function
   End If

   IsHour = False
   intIndex = 1
   Do While intIndex < Len(str_Data) - 2 And Not IsHour
       TempDate = FormatDateTime(Mid(str_Data, intIndex, 5), vbShortTime)
       IsHour = Not TempDate = "00:00:00"
       intIndex = intIndex + 1
   Loop

   HaveHour = IsHour
End Function

Public Function HaveEnglish(ByVal str_Data$) As Boolean
   On Error Resume Next
   Dim intIndex%, IsEnglish As Boolean

   If Len(str_Data) = 0 Then
       HaveEnglish = False
       Exit Function
   End If

   IsEnglish = False
   intIndex = 1
   Do While intIndex < Len(str_Data) And Not IsEnglish
       IsEnglish = (Mid(str_Data, intIndex, 1) >= Chr(65) And Mid(str_Data, intIndex, 1) <= Chr(90)) Or (Mid(str_Data, intIndex, 1) >= Chr(97) And Mid(str_Data, intIndex, 1) <= Chr(122))
       intIndex = intIndex + 1
   Loop

   HaveEnglish = IsEnglish
End Function

Public Function WhereHour(ByVal str_Data$) As Integer
   On Error Resume Next
   Dim intIndex%, IsHour As Boolean, TempDate As Date, tempStr$

   If Len(str_Data) = 0 Then
       WhereHour = False
       Exit Function
   End If

   IsHour = False
   intIndex = 1
   Do While intIndex < Len(str_Data) - 2 And Not IsHour
       tempStr = Mid(str_Data, intIndex, 5)
       'Repair it so VB can detect it as an hour
       If tempStr Like "24:??" Then tempStr = "19:30" '"Just" an hour
       TempDate = FormatDateTime(tempStr, vbShortTime)
       IsHour = Not TempDate = "00:00:00"
       If InStr(1, tempStr, ".") Then IsHour = False
       intIndex = intIndex + 1
   Loop

   WhereHour = intIndex - 1
End Function

Public Sub FormCosmetics(StartUp As Boolean)
    Dim intIntel%
    With frmGlobus
        For intIntel = 0 To .Count - 1
            If Not TypeOf .Controls(intIntel) Is CommonDialog _
            And Not TypeOf .Controls(intIntel) Is Inet _
            And Not TypeOf .Controls(intIntel) Is Frame _
            And Not TypeOf .Controls(intIntel) Is StatusBar _
            Then
                .Controls(intIntel).Enabled = Not .Controls(intIntel).Enabled
            End If
        Next intIntel
        'Special treat
        If StartUp Then
            .txtStartAt.Enabled = False
            .txtEndAt.Enabled = False
        Else
            .txtStartAt.Enabled = .chkPosition.Enabled
            .txtEndAt.Enabled = .chkPositionEnd.Enabled
        End If
        .StatBar.Visible = StartUp
        .StatBarIdle.Visible = Not StartUp 'False when starting-up the process, True when shutting-down
        
        .StatBarIdle.Panels(1).Text = App.Title & " " & App.Major & "." & App.Minor & " ready." 'It should  always be ready...
        
        .cmdGo.Visible = Not StartUp
        .cmdExit.Visible = Not StartUp
        .cmdStop.Visible = StartUp
        .cmdMinimize.Visible = StartUp
        
        .MousePointer = IIf(StartUp, 11, 0)
    End With
End Sub

Public Function FindLast(ByVal strStr As String, ByVal strWhat As String, Optional ByVal Leng As Integer) As Integer
    Dim intNumber%, intIndex%
    
    If Trim(strStr) <> "" And strWhat <> "" Then
        If Leng = Empty Then Leng = Len(strStr)
        intIndex = 1
        intNumber = 0
        Do While intIndex + intNumber <= Leng - Len(strWhat)
            If Mid(strStr, intIndex + intNumber, Len(strWhat)) = strWhat Then
                intIndex = intIndex + intNumber
                intNumber = 1
            Else
                intNumber = intNumber + 1
            End If
        Loop
    End If
    If intIndex <> 0 And Mid(strStr, intIndex, Len(strWhat)) = strWhat Then
        FindLast = intIndex
    Else
        FindLast = 0
    End If
End Function

Public Function FixAvail(ByVal iCine%, ByVal iMov, ByVal StartDay, ByVal EndDay) As Integer
    Dim iI%, Bool As Boolean
    
    iI = 1
    Bool = False
    With frmWorkShop
    Do While Not Bool And iI <= .fraRepairs.UBound
    If .cbxCinemas(iI + 1).ListIndex > -1 And .cbxMovies(iI + 1).ListIndex > -1 Then
        If iCine = .cbxCinemas(iI + 1).ItemData(.cbxCinemas(iI + 1).ListIndex) Then
            If iMov = .cbxMovies(iI + 1).ItemData(.cbxMovies(iI + 1).ListIndex) Then
                If StartDay >= .DTPicker(iI - 1).Value And EndDay <= .DTPickerEnd(iI - 1).Value Then
                    Bool = Trim(.txtAdd(iI - 1).Text) <> "" Or Trim(.txtDelete(iI - 1).Text) <> ""
                End If
            End If
        End If
    End If
    iI = iI + 1
    Loop
    End With
    If Bool Then
        FixAvail = iI - 1
    Else
        FixAvail = -1
    End If
End Function

Public Sub Decode(inArray() As String, outArray() As String, Count As Integer)
    Dim intNumber%, intIndex%, intSmallNum%, intBigNum%
    On Error Resume Next
    
    Count = 1
    'Scan for the smallest and greatest index number
    intNumber = InStr(1, inArray(Count), ">", vbTextCompare) - 1
    intIndex = Val(QuotLessString(Mid(inArray(Count), 1, intNumber)))
    intSmallNum = intIndex
    intBigNum = intIndex
    Do
        Count = Count + 1
        intNumber = InStr(1, inArray(Count), ">", vbTextCompare) - 1
        intIndex = Val(QuotLessString(Mid(inArray(Count), 1, intNumber)))
        If intIndex < intSmallNum Then intSmallNum = intIndex
        If intIndex > intBigNum Then intBigNum = intIndex
    Loop Until IsEmpty(inArray(Count))
    ReDim outArray(intSmallNum To intBigNum)
    Count = 1
    Do
        intNumber = InStr(1, inArray(Count), ">", vbTextCompare) - 1
        intIndex = Val(QuotLessString(Mid(inArray(Count), 1, intNumber)))
        outArray(intIndex) = Right(inArray(Count), Len(inArray(Count)) - intNumber - 1)
        Count = Count + 1
    Loop Until IsEmpty(inArray(Count))
    'ReDim Preserve outArray(Count)
    Count = Count - 1
End Sub

Public Sub CinemasList(inMar() As String, outMar() As String, _
                       int_Start%, int_End%, str_Prefix$)

    On Error Resume Next
    Dim str_Text$
    
    'Obtains and parses the list of cinemas
    With frmGlobus
        .StatBar.Panels(1).Text = "Obtaining the list of cinemas..."
        DoEvents
        If Trim(CinemaObtainURL) <> "" Then
            str_Text = ExtractFileToString(CinemaObtainURL)
        Else
            str_Text = .Inet.OpenURL(str_Prefix + .txtObtain.Text)
        End If
        '.txtINetStatus.Text = ""
        'Troubles: If InStr(1, str_Text, "401 Authorization Required", vbTextCompare) > 0 Then Err.Raise 9210, , "Bad Username or Password"
        .StatBar.Panels(1).Text = "Populating the list of cinemas..."
        DoEvents
        int_Start = InStr(1, str_Text, "cinema_id", vbTextCompare) + 11
        int_End = InStr(int_Start, str_Text, "</select>", vbTextCompare)
        'strText = Right(strText, Len(strText) - intEzer)
        inMar = Split(Mid(str_Text, int_Start, int_End - int_Start), "<option value=")
        Decode inMar, outMar, intMaxCine
    End With
End Sub

Public Sub MoviesList(inMar() As String, outMar() As String, int_Start%, int_End%, _
                      str_Prefix$)
    
    On Error Resume Next
    Dim str_Text$
    
    'Obtains and parses the list of movies
    With frmGlobus
        .StatBar.Panels(1).Text = "Obtaining the list of movies..."
        DoEvents
        If Trim(MovieObtainURL) <> "" Then
            str_Text = ExtractFileToString(MovieObtainURL)
        Else
            str_Text = .Inet.OpenURL(str_Prefix + .txtObtain.Text)
        End If
        '.txtINetStatus.Text = ""
        'Troubles: If InStr(1, str_Text, "401 Authorization Required", vbTextCompare) > 0 Then Err.Raise 9210, , "Bad Username or Password"
        .StatBar.Panels(1).Text = "Populating the list of movies..."
        DoEvents
        int_Start = InStr(1, str_Text, "movie_id[0]", vbTextCompare) + 12
        int_End = InStr(int_Start, str_Text, "</select>", vbTextCompare)
        inMar = Split(Mid(str_Text, int_Start, int_End - int_Start), "<option value=")
        Decode inMar, outMar, intMaxMov
    End With
End Sub

Function GetCinema(Arr() As String, ByVal Cine As String) As Integer
    Dim intMoo%, Freef%, stText$, WasHereBefore As Boolean

    WasHereBefore = False
Bereshit:
    intMoo = AnalyzeCineMov(Arr, Cine)
    If intMoo = 0 Then
      'Getting to the information file
      If WasHereBefore Then GoTo Query 'Mostly happens due to outdated source file
      Freef = FreeFile
      On Error GoTo ErrHandler
      Open DataFile For Input As #Freef
      Do While Not EOF(Freef)
        Line Input #Freef, stText
        'If InStr(1, stText, Cine, vbTextCompare) = 1 Then
        If Cine = Left(stText, FindLast(stText, "=") - 1) Then
            Cine = CleanString(stText)
            Close Freef 'Closes reading from file
            WasHereBefore = True
            GoTo Bereshit
        End If
      Loop
      Close Freef 'Closes reading from file
Query:
      With frmManual
        'Could not identify
        .Caption = .Caption & "cinema"
        .lblCannot.Caption = .lblCannot.Caption & "cinema:"
        .cbxCombo.Clear
        For intMoo = LBound(Arr) To UBound(Arr)
            If Not IsNumeric(Trim(Arr(intMoo))) And Trim(Arr(intMoo)) <> "" Then
                .cbxCombo.AddItem Arr(intMoo)
                .cbxCombo.ItemData(.cbxCombo.NewIndex) = intMoo
            End If
        Next intMoo
        .txtMissing.Text = Cine
        .txtQuick.Text = Cine
        'If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
        NotificationIcon True, "has a question", "What is the cinema """ & Cine & """?", NIIF_INFO
        .Show vbModal, frmGlobus
        NotificationIcon False
        'If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
        Select Case .Tag
            Case "Make", "Revive"
                MsgBox "This option is meant only for movies.", vbExclamation + vbOKOnly, "Error"
                GoTo ForceUpdate
            Case "Update"
ForceUpdate:
                ResetArray Cinemas
                CinemasList Data, Cinemas, 1, 1, "http://" _
                            + frmGlobus.txtUserName.Text + _
                            ":" + frmGlobus.txtPassword.Text _
                            + "@"
                frmGlobus.StatBar.Panels(4).Text = ""
                frmGlobus.StatBar.Panels(1) = "Determining cinema..."
                DoEvents
                Unload frmManual
                GoTo Bereshit
            Case "Cont"
                intMoo = .cbxCombo.ItemData(.cbxCombo.ListIndex)
                Unload frmManual
            Case "Save"
                Freef = FreeFile
                Open DataFile For Append As #Freef
                Print #Freef, Cine & "=" & .cbxCombo.List(.cbxCombo.ListIndex)
                Close Freef
                intMoo = .cbxCombo.ItemData(.cbxCombo.ListIndex)
                Unload frmManual
            Case "Skip"
                intMoo = -1
                Unload frmManual
            Case "Stop"
                intMoo = -2
                Unload frmManual
        End Select
      End With
    End If
    GetCinema = intMoo
    Exit Function
ErrHandler:
    If Err.Number = 76 Or Err.Number = 53 Then GoTo Query 'If not such file or directory
End Function

Function AnalyzeCineMov(Mar() As String, ByVal ToMatch As String) As Integer
    Dim Found As Boolean, intCount%
    
    Found = False
    intCount = LBound(Mar())
    Do While Not Found And intCount <= UBound(Mar())
        Found = Mar(intCount) = ToMatch
        intCount = intCount + 1
    Loop
    If Found Then
        AnalyzeCineMov = intCount - 1
    Else
        AnalyzeCineMov = 0
    End If
End Function

Function GetMovie(Arr() As String, ByVal Mov As String) As Integer
    Dim intMoo%, Freef%, stText$, WasHereBefore As Boolean

    WasHereBefore = False
Bereshit:
    intMoo = AnalyzeCineMov(Arr, Mov)
    If intMoo = 0 Then
      'Getting to the information file
      If WasHereBefore Then GoTo Query 'Mostly happens due to outdated source file
      Freef = FreeFile
      On Error GoTo ErrHandler
      Open DataFile For Input As #Freef
      Do While Not EOF(Freef)
        Line Input #Freef, stText
        'If InStr(1, stText, Mov, vbTextCompare) = 1 Then
        If Mov = Left(stText, FindLast(stText, "=") - 1) Then
            Mov = CleanString(stText)
            Close Freef 'Closes reading from file
            WasHereBefore = True
            GoTo Bereshit
        End If
      Loop
      Close Freef 'Closes reading from file
Query:
      With frmManual
        'Could not identify
        .Caption = .Caption & "movie"
        .lblCannot.Caption = .lblCannot.Caption & "movie:"
        .cbxCombo.Clear
        For intMoo = LBound(Arr) To UBound(Arr)
            If Not IsNumeric(Trim(Arr(intMoo))) And Trim(Arr(intMoo)) <> "" Then
                .cbxCombo.AddItem Arr(intMoo)
                .cbxCombo.ItemData(.cbxCombo.NewIndex) = intMoo
            End If
        Next intMoo
        .txtMissing.Text = Mov
        .txtQuick.Text = Mov
        'If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
        NotificationIcon True, "has a question", "What is the movie """ & Mov & """?", NIIF_INFO
        .Show vbModal, frmGlobus
        NotificationIcon False
        'If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
        Select Case .Tag
            Case "Make"
                .txtMissing.Tag = InputBox("Please enter the name of the movie to be added:", "Make a new entry", .txtMissing.Text)
                If Trim(.txtMissing.Tag) <> "" Then
                    .txtMissing.Tag = "http://" + frmGlobus.txtUserName.Text + ":" + frmGlobus.txtPassword.Text + "@" + frmGlobus.txtUpdateMovies + _
                                    "?id=-99999&name=" + .txtMissing.Tag + "&description=&review_id=0&playing=1&" + _
                                    "cat01=0&cat02=0&cat03=0&cat04=0&cat05=0&cat06=0&cat07=0&cat08=0&cat09=0&cat10=0&cat11=0&cat12=0&cat13=0&cat14=0&cat15=0&cat16=0&cat17=0&cat18=0&cat19=0&cat20=0&cat21=0&cat22=0&cat23=0&cat24=0&cat25=0&cat26=0&cat27=0&cat28=0&cat29=0&cat30=0&cat31=0&cat32=0&cat33=0&cat34=0&cat35=0&cat36=0&cat37=0&cat38=0&cat39=0&cat40=0" + _
                                    "&FormMeta[action]=שמור"
                    .txtQuick.Tag = frmGlobus.Inet.OpenURL(.txtMissing.Tag)
                End If
                
                .Tag = "Update" 'So the list will be updated and contain the new item
                GoTo ForceUpdate 'So that we'll really get there. The "Select Case" statement has built-in "break" statement and we're not allowed to skip from case to case
            Case "Revive"
                .txtMissing.Tag = InputBox("Please enter the ID number of the movie to revive:", "Revive an entry")
                If Trim(.txtMissing.Tag) <> "" Then
                    Dim ID$, reviewID$, MovieName$, Description$, intProgress1%, intProgress2%
                    
                    ID = .txtMissing.Tag
                    
                    .txtMissing.Tag = "http://" + frmGlobus.txtUserName.Text + ":" + frmGlobus.txtPassword.Text + "@" + frmGlobus.txtUpdateMovies + _
                                        "?FormMeta[skip_to]=" + ID + "&FormMeta[action]=עבור"
                    .txtQuick.Tag = frmGlobus.Inet.OpenURL(.txtMissing.Tag)
                    
                    intProgress1 = InStr(1, .txtQuick.Tag, """name""")
                    If intProgress1 = 0 Then GoTo ForceUpdate 'Skip the following
                    intProgress2 = InStr(intProgress1 + Len("""name"" value=""") + 1, .txtQuick.Tag, """")
                    MovieName = NoQuotesString(Mid(.txtQuick.Tag, intProgress1 + Len("""name"" value="""), intProgress2 - (intProgress1 + Len("""name"" value="""))))
                    
                    intProgress1 = InStr(1, .txtQuick.Tag, """review_id""")
                    If intProgress1 = 0 Then GoTo ForceUpdate 'Skip the following
                    intProgress2 = InStr(intProgress1 + Len("""review_id"" value=""") + 1, .txtQuick.Tag, """")
                    reviewID = NoQuotesString(Mid(.txtQuick.Tag, intProgress1 + Len("""review_id"" value="""), intProgress2 - (intProgress1 + Len("""review_id"" value="""))))
                    
                    intProgress1 = InStr(1, .txtQuick.Tag, """description""")
                    If intProgress1 = 0 Then GoTo ForceUpdate 'Skip the following
                    intProgress1 = InStr(intProgress1, .txtQuick.Tag, ">")
                    intProgress2 = InStr(intProgress1 + 1, .txtQuick.Tag, "</textarea>")
                    Description = Mid(.txtQuick.Tag, intProgress1 + 1, intProgress2 - (intProgress1 + 1))
                    
                    'Finished extracting the information, now update the entry with "playing=1"
                    .txtMissing.Tag = "http://" + frmGlobus.txtUserName.Text + ":" + frmGlobus.txtPassword.Text + "@" + frmGlobus.txtUpdateMovies + _
                                    "?id=" + ID + "&name=" + MovieName + "&description=" + Description + "&review_id=" + reviewID + "&playing=1&" + _
                                    "cat01=0&cat02=0&cat03=0&cat04=0&cat05=0&cat06=0&cat07=0&cat08=0&cat09=0&cat10=0&cat11=0&cat12=0&cat13=0&cat14=0&cat15=0&cat16=0&cat17=0&cat18=0&cat19=0&cat20=0&cat21=0&cat22=0&cat23=0&cat24=0&cat25=0&cat26=0&cat27=0&cat28=0&cat29=0&cat30=0&cat31=0&cat32=0&cat33=0&cat34=0&cat35=0&cat36=0&cat37=0&cat38=0&cat39=0&cat40=0" + _
                                    "&FormMeta[action]=שמור"
                    .txtQuick.Tag = frmGlobus.Inet.OpenURL(.txtMissing.Tag)
                End If
                
                .Tag = "Update" 'So the list will be updated and contain the new item
                GoTo ForceUpdate 'So that we'll really get there. The "Select Case" statement has built-in "break" statement and we're not allowed to skip from case to case
            Case "Update"
ForceUpdate:
                ResetArray Movies
                MoviesList Data, Movies, 1, 1, "http://" _
                            + frmGlobus.txtUserName.Text + _
                            ":" + frmGlobus.txtPassword.Text _
                            + "@"
                frmGlobus.StatBar.Panels(4).Text = ""
                frmGlobus.StatBar.Panels(1) = "Determining movie..."
                DoEvents
                Unload frmManual
                GoTo Bereshit
            Case "Cont"
                intMoo = .cbxCombo.ItemData(.cbxCombo.ListIndex)
                Unload frmManual
            Case "Save"
                Freef = FreeFile
                Open DataFile For Append As #Freef
                Print #Freef, Mov & "=" & .cbxCombo.List(.cbxCombo.ListIndex)
                Close Freef
                intMoo = .cbxCombo.ItemData(.cbxCombo.ListIndex)
                Unload frmManual
            Case "Skip"
                intMoo = -1
                Unload frmManual
            Case "Stop"
                intMoo = -2
                Unload frmManual
        End Select
      End With
    End If
    GetMovie = intMoo
    Exit Function
ErrHandler:
    If Err.Number = 76 Or Err.Number = 53 Then GoTo Query 'If not such file or directory
End Function

Public Sub NotificationIcon(Show As Boolean, Optional ByRef tipText As String, Optional ByRef tipBalloon As String, Optional ByVal BalloonIcon As Long = &H0)
    If Show And frmGlobus.cmdMinimize.Tag = "1" Then
        'Set the individual values of the NOTIFYICONDATA data type
        'nID.cbSize = Len(nID)
        'nID.hwnd = frmGlobus.hwnd
        'nID.uID = vbNull
        'nID.uFlags = NIF_ICON Or NIF_TIP Or NIF_INFO
        'nID.uCallbackMessage = WM_MOUSEMOVE
        'nID.hIcon = frmWorkShop.Icon
        'nID.szTip = App.Title & " " & tipText & " (dbl-click to open)" & vbNullChar
        'nID.szInfoTitle = App.Title & " " & tipText & vbNullChar
        'nID.szInfo = tipBalloon & vbNullChar
        'nID.uTimeout = 15000 'in milliseconds
        'nID.dwInfoFlags = BalloonIcon
        
        ModifyNotifyIcon frmWorkShop.Icon.Handle, App.Title & " " & tipText
        
        ShowBalloonTip tipBalloon, App.Title & " " & tipText, NIIF_INFO
        
        frmGlobus.cmdMinimize.Tag = "2" 'For Form_MouseMove
        
        While frmGlobus.cmdMinimize.Tag = "2" 'Waiting for the user to react
            DoEvents
        Wend
        
        frmGlobus.Show
    ElseIf Not Show And frmGlobus.cmdMinimize.Tag = "1" Then
        'Set the individual values of the NOTIFYICONDATA data type
        'nID.cbSize = Len(nID)
        'nID.hwnd = frmGlobus.hwnd
        'nID.uID = vbNull
        'nID.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO 'NIF_INFO to remove the balloon
        'nID.uCallbackMessage = WM_MOUSEMOVE
        'nID.hIcon = frmGlobus.Icon
        'nID.szTip = App.Title & " (dbl-click to open)" & vbNullChar
        'nID.szInfoTitle = vbNullChar
        'nID.szInfo = vbNullChar
        'nID.uTimeout = 0
        'nID.dwInfoFlags = &H0
        
        ModifyNotifyIcon frmGlobus.Icon.Handle, App.Title
        
        frmGlobus.Hide
    End If
End Sub

Public Function GetState(s As Integer) As String
'from http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnexpvb/html/addinginternettransfercontrol.asp
Select Case s
   Case 0
      GetState = "No state information is available."
   
   Case 1
      GetState = "Looking up the IP address for the remote server."
   
   Case 2
      GetState = "Found the IP address for the remote server."
   
   Case 3
      GetState = "Connecting to the remote server."
   
   Case 4
      GetState = "Connected to the remote server."
   
   Case 5
      GetState = "Requesting information from the remote server."
   
   Case 6
      GetState = "The request was sent successfully to the remote server."
   
   Case 7
      GetState = "Receiving a response from the remote server."
   
   Case 8
      GetState = "The response was received successfully from the " & _
         "remote server."
   
   Case 9
      GetState = "Disconnecting from the remote server."
   
   Case 10
      GetState = "Disconnected from the remote server."
   
   Case 11
      GetState = "An error has occurred while communicating with the " & _
         "remote server."
   
   Case 12
      GetState = "The request was completed, all data has been received."
   
   Case Else
      GetState = "Unknown state: " & FormatNumber(s, 0)
   
   End Select
   
End Function

Public Function FindBiggest(ArMar()) As Integer
    Dim i%
    On Error GoTo Quit
    
    i = 1 'And not LBound(ArMar())! LBound doesn't contain anything
    Do While Not ArMar(i) = "" And i <= UBound(ArMar())
        i = i + 1
    Loop
    
Quit:
    FindBiggest = i - 1
End Function

Public Sub ResetArray(Maarach() As String)
    Dim i%
    On Error GoTo Quit
    
    i% = LBound(Maarach())
    Do While (Not IsEmpty(Maarach(i)) Or Not IsNull(Maarach(i))) And i <= UBound(Maarach())
        Maarach(i) = ""
        i = i + 1
    Loop

'    For i = 1 To UBound(Maarach())
'        Maarach(i) = ""
'    Next i

Quit:
End Sub

Public Sub ResetVarArray(Maarach() As Variant)
    Dim i%
    On Error GoTo Quit
    
    i = LBound(Maarach())
    Do While (Not IsEmpty(Maarach(i)) Or Not IsNull(Maarach(i))) And i <= UBound(Maarach())
        Maarach(i) = ""
        i = i + 1
    Loop

'    For i = 1 To UBound(Maarach())
'        Maarach(i) = ""
'    Next i

Quit:
End Sub

Public Function InDates(arDates(), serDate As Date) As Boolean
    Dim i%, IsExist As Boolean
    On Error GoTo Quit
    
    i = LBound(arDates())
    IsExist = False
    
    Do While (Not IsEmpty(arDates(i)) Or Not IsNull(arDates(i))) And arDates(i) <> "" And i <= UBound(arDates()) And Not IsExist
        If arDates(i) = serDate Then IsExist = True
        i = i + 1
    Loop
    
    InDates = IsExist
    
Quit:
End Function

Public Function AnalyzeDay(ByVal CurDay$) As Integer
    Dim NumAdd%
    
    'Make it ביעי and not רביעי, just in case
    If InStr(1, CurDay, "ה", vbTextCompare) > 0 Or InStr(1, CurDay, "ח", vbTextCompare) > 0 Or InStr(1, CurDay, "חמישי", vbTextCompare) > 0 Then
        NumAdd = 0
    ElseIf InStr(1, CurDay, "שישי", vbTextCompare) > 0 Or (InStr(1, CurDay, "ו", vbTextCompare) > 0 And InStr(1, CurDay, "ראשון", vbTextCompare) = 0 And InStr(1, CurDay, "מוצ", vbTextCompare) = 0) Or InStr(1, CurDay, "שישי", vbTextCompare) > 0 Then
        NumAdd = 1
    ElseIf InStr(1, CurDay, "שבת", vbTextCompare) > 0 Or (InStr(1, CurDay, "ש", vbTextCompare) > 0 And InStr(1, CurDay, "ראשון", vbTextCompare) = 0 And InStr(1, CurDay, "שני", vbTextCompare) = 0 And InStr(1, CurDay, "שלישי", vbTextCompare) = 0 And InStr(1, CurDay, "חמישי", vbTextCompare) = 0 And InStr(1, CurDay, "שישי", vbTextCompare) = 0) Or InStr(1, CurDay, "שבת", vbTextCompare) > 0 Then
        NumAdd = 2
    ElseIf InStr(1, CurDay, "א", vbTextCompare) > 0 Or InStr(1, CurDay, "ן", vbTextCompare) > 0 Or InStr(1, CurDay, "ראשון", vbTextCompare) > 0 Then
        NumAdd = 3
    ElseIf (InStr(1, CurDay, "ב", vbTextCompare) > 0 And InStr(1, CurDay, "ביעי", vbTextCompare) = 0 And InStr(1, CurDay, "שבת", vbTextCompare) = 0) Or InStr(1, CurDay, "שני", vbTextCompare) > 0 Then
        NumAdd = 4
    ElseIf InStr(1, CurDay, "ג", vbTextCompare) > 0 Or InStr(1, CurDay, "שלישי", vbTextCompare) > 0 Then
        NumAdd = 5
    ElseIf InStr(1, CurDay, "ד", vbTextCompare) > 0 Or InStr(1, CurDay, "ביעי", vbTextCompare) > 0 Then
        NumAdd = 6
    Else
        NumAdd = -1
        'Err.Raise 5253, , "Invalid range for dates"
    End If
    
    AnalyzeDay = NumAdd
End Function

Public Function NewAnalyzeDay(ByVal CurDay$) As Integer
    Dim ThisDay%
    
    'Make it ביעי and not רביעי, just in case
    If InStr(1, CurDay, "ה", vbTextCompare) > 0 Or InStr(1, CurDay, "ח", vbTextCompare) > 0 Or InStr(1, CurDay, "חמישי", vbTextCompare) > 0 Then
        ThisDay = vbThursday
    ElseIf InStr(1, CurDay, "שישי", vbTextCompare) > 0 Or (InStr(1, CurDay, "ו", vbTextCompare) > 0 And InStr(1, CurDay, "ראשון", vbTextCompare) = 0 And InStr(1, CurDay, "מוצ", vbTextCompare) = 0) Or InStr(1, CurDay, "שישי", vbTextCompare) > 0 Then
        ThisDay = vbFriday
    ElseIf InStr(1, CurDay, "שבת", vbTextCompare) > 0 Or (InStr(1, CurDay, "ש", vbTextCompare) > 0 And InStr(1, CurDay, "ראשון", vbTextCompare) = 0 And InStr(1, CurDay, "שני", vbTextCompare) = 0 And InStr(1, CurDay, "שלישי", vbTextCompare) = 0 And InStr(1, CurDay, "חמישי", vbTextCompare) = 0 And InStr(1, CurDay, "שישי", vbTextCompare) = 0) Or InStr(1, CurDay, "שבת", vbTextCompare) > 0 Then
        ThisDay = vbSaturday
    ElseIf InStr(1, CurDay, "א", vbTextCompare) > 0 Or InStr(1, CurDay, "ן", vbTextCompare) > 0 Or InStr(1, CurDay, "ראשון", vbTextCompare) > 0 Then
        ThisDay = vbSunday
    ElseIf (InStr(1, CurDay, "ב", vbTextCompare) > 0 And InStr(1, CurDay, "ביעי", vbTextCompare) = 0 And InStr(1, CurDay, "שבת", vbTextCompare) = 0) Or InStr(1, CurDay, "שני", vbTextCompare) > 0 Then
        ThisDay = vbMonday
    ElseIf InStr(1, CurDay, "ג", vbTextCompare) > 0 Or InStr(1, CurDay, "שלישי", vbTextCompare) > 0 Then
        ThisDay = vbTuesday
    ElseIf InStr(1, CurDay, "ד", vbTextCompare) > 0 Or InStr(1, CurDay, "ביעי", vbTextCompare) > 0 Then
        ThisDay = vbWednesday
    Else
        ThisDay = -1
        'Err.Raise 5253, , "Invalid range for dates"
    End If
    
    NewAnalyzeDay = ThisDay
End Function

Function dhNextDOW(intDOW As Integer, _
 Optional dtmDate As Date = 0) As Date ' Taken from MSDN, http://msdn.microsoft.com/library/en-us/dnvbadev/html/findingnextorpreviousweekday.asp
    ' Find the next specified day of the week
    ' after the specified date.
    ' MODIFIED: Does NOT returns the next specified day of the week,
    '           if the given date IS the specified day of week.
    '           It returns the given date instead.
    Dim intTemp As Integer
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    intTemp = Weekday(dtmDate)
    If intTemp = intDOW Then
        dhNextDOW = dtmDate
    Else
        dhNextDOW = dtmDate - intTemp + intDOW + _
         IIf(intTemp < intDOW, 0, 7)
    End If
End Function

Public Function ExtractFileToString(ByVal FileName As String) As String
    'On Error Resume Next
    Dim FreeFileNum As Integer, LineText As String
    
    ExtractFileToString = ""
    FreeFileNum = FreeFile
    Open FileName For Input As #FreeFileNum
    Do While Not EOF(FreeFileNum)
        Line Input #FreeFileNum, LineText
        ExtractFileToString = ExtractFileToString & LineText '& vbCrLf
    Loop
    Close #FreeFileNum
End Function

Public Function ValidInput(Optional ByVal SkipSource As Boolean = False) As Boolean
    Dim checkBtn As Boolean, optBtn As OptionButton
    
  With frmGlobus
    If Trim(.txtUserName.Text) = "" Then
        'MsgBox "Please input the username.", vbCritical + vbOKOnly + vbApplicationModal, "Error"
        ShoutAtUser frmGlobus, .txtUserName, "Please input the username."
        '.txtUserName.SetFocus
        ValidInput = False
        Exit Function
    ElseIf Trim(.txtPassword.Text) = "" Then
        'MsgBox "Please input the password.", vbCritical + vbOKOnly + vbApplicationModal, "Error"
        '.txtPassword.SetFocus
        ShoutAtUser frmGlobus, .txtPassword, "Please input the password."
        ValidInput = False
        Exit Function
    ElseIf Not SkipSource And Trim(.txtSource.Text) = "" Then
        'MsgBox "Please specify the source file.", vbCritical + vbOKOnly + vbApplicationModal, "Error"
        '.txtSource.SetFocus
        ShoutAtUser frmGlobus, .txtSource, "Please specify the source file."
        ValidInput = False
        Exit Function
    ElseIf Not SkipSource And _
    ((LCase(Right(.txtSource.Text, 4)) <> ".htm" And LCase(Right(.txtSource.Text, 5)) <> ".html" And LCase(Right(.txtSource.Text, 4)) <> ".txt") _
    Or (StrComp(LCase(Dir(.txtSource.Text)), "", vbTextCompare) = 0)) Then 'The last term checks to see whether the file exists
        'MsgBox "Please enter a valid source file.", vbCritical + vbOKOnly + vbApplicationModal, "Error"
        '.txtSource.SetFocus
        ShoutAtUser frmGlobus, .txtSource, "Please enter a valid source file."
        ValidInput = False
        Exit Function
    ElseIf Trim(.txtAddress.Text) = "" Then
        'MsgBox "Please specify the destination address.", vbCritical + vbOKOnly + vbApplicationModal, "Error"
        '.txtAddress.SetFocus
        ShoutAtUser frmGlobus, .txtAddress, "Please specify the destination address."
        ValidInput = False
        Exit Function
    ElseIf Trim(.txtObtain.Text) = "" Then
        ShoutAtUser frmGlobus, .txtObtain, "Please specify the address for the list of cinemas and movies."
        ValidInput = False
        Exit Function
    ElseIf Trim(.txtUpdateCinemas.Text) = "" Then
        ShoutAtUser frmGlobus, .txtUpdateCinemas, "Please specify the address for the list of cinemas."
        ValidInput = False
        Exit Function
    ElseIf Trim(.txtUpdateMovies.Text) = "" Then
        ShoutAtUser frmGlobus, .txtUpdateMovies, "Please specify the address for the list of movies."
        ValidInput = False
        Exit Function
    ElseIf .chkPosition.Value = vbChecked And Not IsNumeric(.txtStartAt.Text) Then
        'MsgBox "Please input a valid cinema number.", vbCritical + vbOKOnly + vbApplicationModal, "Error"
        '.txtStartAt.SetFocus
        ShoutAtUser frmGlobus, .txtStartAt, "Please input a valid cinema number."
        '.txtStartAt.SelText = .txtStartAt.Text
        ValidInput = False
        Exit Function
    ElseIf .chkPositionEnd.Value = vbChecked And Not IsNumeric(.txtEndAt.Text) Then
        'MsgBox "Please input a valid cinema number.", vbCritical + vbOKOnly + vbApplicationModal, "Error"
        '.txtEndAt.SetFocus
        '.txtEndAt.SelText = .txtEndAt.Text
        ShoutAtUser frmGlobus, .txtEndAt, "Please input a valid cinema number."
        ValidInput = False
        Exit Function
    End If
    If Trim(.txtStartAt.Text) = "" Then .chkPosition.Value = False
    If Trim(.txtEndAt.Text) = "" Then .chkPositionEnd.Value = False
  End With 'end of frmGlobus references
    ValidInput = True
End Function

Public Sub ShoutAtUser(frm As Form, ctrl As Control, Optional textMessage$ = "Invalid input.", Optional titleText$ = "Error", Optional MsgBoxSeverity As VbMsgBoxStyle = vbCritical, Optional BalloonSeverity As TextboxBalloonTipIconConstants = TTI_ERROR)
    On Error Resume Next
    
    With frm
        .ctrl.SelText = .ctrl
    End With
    If Running9x Then
        ctrl.SetFocus
        MsgBox textMessage, MsgBoxSeverity + vbOKOnly + vbApplicationModal, titleText
    Else
        Dim ctrlhwnd As Long
        
        ctrlhwnd = ctrl.hwnd
        If TypeOf ctrl Is ComboBox Then ctrlhwnd = GetComboEditHWND(frm.ctrl)
        ShowTextboxBalloonTip ctrlhwnd, textMessage, titleText, (BalloonSeverity)
    End If
End Sub

Public Function RunningUnderNT() As Boolean
    Dim OSV As OSVERSIONINFO
    OSV.dwOSVersionInfoSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        If OSV.dwPlatformID = VER_PLATFORM_WIN32_NT Then
            RunningUnderNT = True
        End If
    End If
End Function

Sub Main()
    Dim Confirm As Long
    
    InitCommonControlsVB
    
    If App.PrevInstance Then 'if the program is already running and user wants to run it again,
        MsgBox """" & App.Title & """ is already running.", vbSystemModal + vbExclamation, "Cannot launch multiple instances"
        AppActivate App.Title 'Set focus to the already-running program
        Exit Sub 'Terminate this application
    End If
    If Right(App.Path, 2) = ":\" Then 'Data file is in the root folder
        DataFile = App.Path & "Globus.dat"
        App.StartLogging App.Path & "Error.log", vbLogToFile
    Else
        DataFile = App.Path & "\Globus.dat"
        App.StartLogging App.Path & "\Error.log", vbLogToFile
    End If
    
    argProcessCMDLine 'Residents in modArgs,
    If argSwitchPresent("/reversedate") Then
        DateFormat = "dd/MM/yyyy"
    Else
        DateFormat = "MM/dd/yyyy"
    End If
'    If InStr(1, LCase(Command), "/open", vbTextCompare) > 0 Then Shell "Notepad " & DataFile, vbMinimizedFocus
    If argSwitchPresent("/open") Then Shell "Notepad " & DataFile, vbMinimizedFocus
    'If InStr(1, LCase(Command), "/erase", vbTextCompare) > 0 Then
    If argSwitchPresent("/erase") Then
        Confirm = MsgBox("Are you sure you want to erase the information file?" & vbCrLf & "Click ""Yes"" to erase, click ""No"" to exit and click" & vbCrLf & """Cancel"" to load the program normally.", vbSystemModal + vbQuestion + vbYesNoCancel, "Erasing the information file")
        If Confirm = vbYes Then
            Kill DataFile
            Exit Sub
        ElseIf Confirm = vbNo Then
            Exit Sub
        End If
    End If
'    If Command <> "" And InStr(1, LCase(Command), "/open", vbTextCompare) = 0 And InStr(1, LCase(Command), "/erase", vbTextCompare) = 0 Then
'        frmGlobus.txtSource.Text = NoQuotesString(Command)
'    End If
'    If InStr(1, LCase(Command), "/pwd", vbTextCompare) > 0 Then
'        Confirm = InStr(1, LCase(Command), "/pwd", vbTextCompare) + 4
'        Confirm = InStr(Confirm, Command, " ", vbTextCompare) - InStr(1, LCase(Command), "/pwd", vbTextCompare)
'        frmGlobus.txtPassword.Text = Mid(Command, InStr(1, LCase(Command), "/pwd", vbTextCompare) + 4, Confirm)
'    End If
    If argSwitchPresent("/pwd*", Confirm, True) Then
        frmGlobus.txtPassword.Text = Mid(argv(Confirm), InStr(1, argv(Confirm), "/pwd", vbTextCompare) + 4)
    End If
    If argSwitchPresent("/cinemaobtain:*", Confirm, True) Then
        CinemaObtainURL = NoQuotesString(Mid(argv(Confirm), InStr(1, argv(Confirm), "/cinemaobtain:", vbTextCompare) + 14))
        frmGlobus.txtObtain.Text = "[Custom]"
    End If
    If argSwitchPresent("/movieobtain:*", Confirm, True) Then
        MovieObtainURL = NoQuotesString(Mid(argv(Confirm), InStr(1, argv(Confirm), "/movieobtain:", vbTextCompare) + 13))
        frmGlobus.txtObtain.Text = "[Custom]"
    End If
    For Confirm = 1 To argc - 1& 'Bypassing the c:\...Globus.Exe argument
        If Trim(argv(Confirm)) <> "" And InStr(1, argv(Confirm), "/open", vbTextCompare) = 0 And InStr(1, argv(Confirm), "/erase", vbTextCompare) = 0 And InStr(1, argv(Confirm), "/pwd", vbTextCompare) = 0 And InStr(1, argv(Confirm), "/reversedate", vbTextCompare) = 0 And InStr(1, argv(Confirm), "/cinemaobtain", vbTextCompare) = 0 And InStr(1, argv(Confirm), "/movieobtain", vbTextCompare) = 0 Then
            frmGlobus.txtSource.Text = NoQuotesString(argv(Confirm))
            Exit For
        End If
    Next Confirm
    ResetArray Cinemas
    ResetArray Movies
    Running9x = Not RunningUnderNT 'Check to see what operating system version is running
    BeingStopped = False
    frmGlobus.Show
End Sub
