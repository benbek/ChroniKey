Attribute VB_Name = "modDoer"
Option Explicit

Private Enum Distributers
    Automat = 0
    Compiled = 1
    GG = 2
    OldLev                  'and so on..
    LevExcel
    RavHen
    London
    Haifa
    Jerusalem
    Dizengof
End Enum

Public Sub SetMovieWeek(ByRef Date1 As Date, ByRef Date2 As Date)
    'Search for next thursday
    Dim AddDay As Integer, Date0 As Date
    
    Date0 = Format(Date$, DateFormat)
    AddDay = Weekday(Date0)
    Select Case AddDay
        Case vbSunday To vbThursday
            AddDay = vbThursday - AddDay
        Case vbFriday
            AddDay = 6
        Case vbSaturday
            AddDay = 5
    End Select
    Date1 = DateAdd("d", AddDay, Date0)
    Date2 = DateAdd("d", 6, Date1)
End Sub

Public Function LetsDoIt(Simulate As Boolean) As Boolean
    Dim strText$, strPrefix$, intStart%, intEnd%, FreeFileNum%, _
        Times$(MaxTimes), Days$(MaxTimes), Dates(MaxTimes * 2 + 1), _
        intCount%, SkipCinemas As Boolean, MovStartDate$, MovEndDate$, _
        MoviePos As Long, HMP&, SkipMovies As Boolean, DontFilmMore As Boolean, _
        DontDoMore As Boolean, intI%, MarTimes$(MaxTimes), intTimes% ', _
        'CinemaCity(MaxMovies, MaxTimes), CCDates() 'Cinema City special handling
    Dim RavHenFirstMovieInBoard As Boolean, RavHenHasJustExited As Boolean
    On Error GoTo Oops
    
    LetsDoIt = False 'Set initial value
    
    If Not ValidInput Then Exit Function
    
    GetUpdate 'Check to see whether update is available
    
    With frmGlobus

    strPrefix = "http://" + .txtUserName.Text + ":" _
                + .txtPassword.Text + "@"
    SkipCinemas = .chkPosition.Value
    SkipMovies = frmWorkShop.chkStartMovie.Value And .chkPosition.Value 'Valid only if both happen
    DontDoMore = False
    DontFilmMore = False
    
    'Interface changes
    FormCosmetics True
    '.cmdGo.Enabled = False
    '.cmdExit.Enabled = False
    '.fraStatus.Enabled = True

    If Simulate Then
        .cmdSimulate.Tag = .Caption
        .Caption = .Caption & " [Simulating]"
    End If
    
    CinemasList Data, Cinemas, intStart, intEnd, strPrefix
    .StatBar.Panels(4).Text = ""
    MoviesList Data, Movies, intStart, intEnd, strPrefix
    .StatBar.Panels(4).Text = ""
    If UBound(Cinemas) < 2 Or UBound(Movies) < 2 Then Err.Raise 5111, , "No cinemas or no movies have been found"
    
    RavHenFirstMovieInBoard = True
    RavHenHasJustExited = False
    intNumCine = 0
    intNumMov = 0
    intCinemas = 0
    intMovies = 0
    .StatBar.Panels(1).Text = "Opening the source file..."
    DoEvents
    FreeFileNum = FreeFile
    Open .txtSource.Text For Input As #FreeFileNum
    Do While Not EOF(FreeFileNum)
      If .cbxType.ListIndex = Automat Then
        DetermineSourceType FreeFileNum
        If EOF(FreeFileNum) Then Err.Raise 5252, , "The cinema was not detected"
      End If
      If .cbxType.ListIndex = GG Then 'Globus Group
        Line Input #FreeFileNum, strText
        .StatBar.Panels(2).Text = ""
        Do While InStr(1, HTMLessString(strText), "סרטים לשבוע " & EndDate & " - " & StartDate, vbTextCompare) > 0 And InStr(1, HTMLessString(strText), "הנהלה", vbTextCompare) = 0
                 .StatBar.Panels(1).Text = "Found matching dates."
                 DoEvents
GGStartCinema:
                Line Input #FreeFileNum, strText
                 .StatBar.Panels(1).Text = "Determining cinema..."
                 .StatBar.Panels(2).Text = ""
                 DoEvents
                 'strText = Right(strText, Len(strText) - InStr(1, strText, ":", vbTextCompare))
                 strText = HTMLessString(strText)
                 intCinemas = GetCinema(Cinemas, strText)
                 If intCinemas = -1 Then Exit Do 'User chose to skip in frmManual
                 If intCinemas = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(2).Text = Cinemas(intCinemas)
                 If DontDoMore Then
                    Exit Do
                 End If
                 If SkipCinemas And intCinemas <> .txtStartAt.Text Then
                    Exit Do
                 End If
                 SkipCinemas = False
                 If .chkPositionEnd.Value And intCinemas = .txtEndAt.Text Then
                    DontDoMore = True
                 End If
                 For intCount = 1 To 9
                    Line Input #FreeFileNum, strText
                 Next intCount
                 If InStr(1, strText, "<H1", vbTextCompare) > 0 Then Exit Do
            Do
GGStartMovie:
                 .StatBar.Panels(1).Text = "Searching for times..."
                 DoEvents
                 Do While InStr(1, strText, "<td align=right>", vbTextCompare) = 0 And InStr(1, strText, "<td align=middle>", vbTextCompare) = 0
                    Line Input #FreeFileNum, strText
                 Loop
                 If InStr(1, strText, "<td align=middle>", vbTextCompare) > 0 Then GoTo GGStartCinema
                 .StatBar.Panels(1).Text = "Found appearance of times for cinema #" & intCinemas & "."
                 DoEvents
                 Line Input #FreeFileNum, strText 'Get to the times
                 ResetArray Times
                 intCount = 1
                 Do While InStr(1, strText, "</td>", vbTextCompare) = 0 And intCount <= MaxTimes
                    strText = DashString(HTMLessString(strText))
                    Times(intCount) = strText
                    Line Input #FreeFileNum, strText
                    Do While InStr(1, strText, "<br", vbTextCompare) = 0 And InStr(1, strText, "</td>", vbTextCompare) = 0
                        Times(intCount) = Times(intCount) & " " & strText
                        Line Input #FreeFileNum, strText
                    Loop
                    intCount = intCount + 1
                 Loop
                 .StatBar.Panels(1).Text = "Extracted times; searching for days..."
                 DoEvents
                 Do While InStr(1, strText, "<td align=right>", vbTextCompare) = 0
                    Line Input #FreeFileNum, strText
                 Loop
                 .StatBar.Panels(1).Text = "Found appearance of days for cinema #" & intCinemas & "."
                 DoEvents
                 Line Input #FreeFileNum, strText 'Get to the dates
                 ResetArray Days
                 intCount = 1
                 Do While InStr(1, strText, "</td>", vbTextCompare) = 0 And intCount <= MaxTimes
                    strText = HTMLessString(strText)
                    Days(intCount) = strText
                    Line Input #FreeFileNum, strText
                    Do While InStr(1, strText, "<br", vbTextCompare) = 0 And InStr(1, strText, "</td>", vbTextCompare) = 0
                        Days(intCount) = Days(intCount) & " " & strText
                        Line Input #FreeFileNum, strText
                    Loop
                    intCount = intCount + 1
                 Loop
                 .StatBar.Panels(1).Text = "Extracted days; searching for movie..."
                 DoEvents
                 Do While InStr(1, strText, "<td align=middle>", vbTextCompare) = 0
                    Line Input #FreeFileNum, strText
                 Loop
                 Do
                    Line Input #FreeFileNum, strText
                 Loop Until Not HaveHour(strText)
                 .StatBar.Panels(1).Text = "Determining movie..."
                 .StatBar.Panels(3).Text = ""
                 DoEvents
                 intMovies = GetMovie(Movies, Trim(strText))
                 If intMovies = -1 Then GoTo GGStartMovie 'User chose to skip in frmManual
                 If intMovies = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(3).Text = Movies(intMovies)
                 If DontFilmMore Then GoTo GGStartMovie
                 If SkipMovies Then
                    If intMovies <> frmWorkShop.cbxMovies(0).ItemData(frmWorkShop.cbxMovies(0).ListIndex) Then GoTo GGStartMovie
                 End If
                 SkipMovies = False
                 If frmWorkShop.chkEndMovie.Value = 1 Then
                    If intMovies = frmWorkShop.cbxMovies(1).ItemData(frmWorkShop.cbxMovies(1).ListIndex) Then DontDoMore = True
                 End If
                 ResetArray MarTimes
                 ResetVarArray Dates
                 '.StatBar.Panels(1).Text = "Handling Cinema City, if applicable..."
                 'intCount = 1
                 'Do While intCinemas = 95 And Days(intCount) <> "" And intCount <= MaxTimes
                 '   If CinemaCity(intMovies, intCount) = "" Then
                 '       CinemaCity(intMovies, intCount) = Days(intCount) & "/\" & Trim(Times(intCount))
                 '   Else
                 '       CinemaCity(intMovies, intCount) = CinemaCity(intMovies, intCount) & " " & Trim(Times(intCount))
                 '   End If
                 'Loop
                 .StatBar.Panels(1).Text = "Filling up the buffer..."
                 DoEvents
                 intCount = 1
                 intI = 1
                 intTimes = 1
                 Do While intCount <= MaxTimes And Days(intCount) <> ""
                    If InStr(1, Days(intCount), "א - ד", vbTextCompare) > 0 Then
                        Dates(intI) = DateAdd("d", 3, StartDate)
                        intI = intI + 1
                        Dates(intI) = EndDate
                        intI = intI + 1
                        FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                        MarTimes(intTimes) = Times(intCount)
                        intTimes = intTimes + 1
                        'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Times(intCount)
                    Else
                        If InStr(1, Days(intCount), "-", vbTextCompare) > 0 Then
                            'Do
                                frmDates.txtMissing.Text = Days(intCount)
                                If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
                                frmDates.Show frmGlobus, vbModal
                                'Dates(intI) = InputBox("Please enter the ""from"" date for " & Days(intCount) & " (dd/mm/yyyy):", "Cannot analyze date", StartDate)
                                Dates(intI) = FormatDateTime(frmDates.DTPicker(1).Value, vbShortDate)
                            'Loop Until Trim(Dates(intI)) <> "" And IsDate(Dates(intI))
                            'intI = intI + 1
                            'Do
                            '    Dates(intI) = InputBox("Please enter the ""to"" date for " & Days(intCount) & " (dd/mm/yyyy):", "Cannot analyze date", EndDate)
                            'Loop Until Trim(Dates(intI)) <> "" And IsDate(Dates(intI))
                            Dates(intI + 1) = FormatDateTime(frmDates.DTPicker(2).Value, vbShortDate)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Times(intCount)
                        Else
                        If InStr(1, Days(intCount), "ה", vbTextCompare) > 0 Then
                            Dates(intI) = StartDate
                            intI = intI + 1
                            Dates(intI) = StartDate
                            intI = intI + 1
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Times(intCount)
                        End If
                        If InStr(1, Days(intCount), "שישי", vbTextCompare) > 0 Or InStr(1, Days(intCount), "ו", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 1, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Times(intCount)
                        End If
                        If InStr(1, Days(intCount), "שבת", vbTextCompare) > 0 Or InStr(1, Days(intCount), "ש", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 2, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Times(intCount)
                        End If
                        If InStr(1, Days(intCount), "א", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 3, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Times(intCount)
                        End If
                        If InStr(1, Days(intCount), "ב", vbTextCompare) > 0 Then 'And InStr(1, Days(intCount), "שבת", vbTextCompare) = 0 Then
                            If Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ש" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ת" Then
                                Dates(intI) = DateAdd("d", 4, StartDate)
                                Dates(intI + 1) = Dates(intI)
                                intI = intI + 2
                                FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                                MarTimes(intTimes) = Times(intCount)
                                intTimes = intTimes + 1
                                'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                    Dates, Trim(Str(intMovies)), Times(intCount)
                            End If
                        End If
                        If InStr(1, Days(intCount), "ג", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 5, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Times(intCount)
                        End If
                        If InStr(1, Days(intCount), "ד", vbTextCompare) > 0 Then
                            Dates(intI) = EndDate
                            Dates(intI + 1) = EndDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Times(intCount)
                        End If
                      End If
                    End If
                 intCount = intCount + 1
                 Loop
            If Not Simulate Then
                POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                        Dates, Trim(Str(intMovies)), MarTimes
            End If
            intNumMov = intNumMov + 1
            .StatBar.Panels(4).Text = ""
            .StatBar.Panels(1).Text = "Finished uploading data."
            DoEvents
            For intCount = 1 To 4
                Line Input #FreeFileNum, strText
            Next intCount
            .StatBar.Panels(3).Text = ""
            If InStr(1, strText, "style", vbTextCompare) > 0 Then
                Line Input #FreeFileNum, strText
                If InStr(1, strText, "<H1", vbTextCompare) > 0 Then Exit Do
            End If
            If InStr(1, strText, "/html", vbTextCompare) > 0 Then Exit Do
            Loop
            intNumCine = intNumCine + 1
        Loop
      ElseIf .cbxType.ListIndex = OldLev Then 'Lev
        MovStartDate = Format(StartDate, "d/m")
        MovEndDate = Format(EndDate, "d/m")
        Line Input #FreeFileNum, strText
        strText = Trim(strText)
        Do While InStr(1, strText, MovStartDate & "-" & MovEndDate, vbTextCompare) > 0 Or InStr(1, strText, MovEndDate & "-" & MovStartDate, vbTextCompare) > 0
                 .StatBar.Panels(1).Text = "Found matching dates; Determining cinema..."
                 .StatBar.Panels(2).Text = ""
                 DoEvents
                 If DontDoMore Then
                    Exit Do
                 End If
                 strText = Trim(strText)
                 strText = Mid(strText, 1, Len(strText) - Len(MovStartDate & "-" & MovEndDate))
                 intCinemas = GetCinema(Cinemas, Trim(strText))
                 If intCinemas = -1 Then Exit Do 'User chose to skip in frmManual
                 If intCinemas = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(2).Text = Cinemas(intCinemas)
                 If SkipCinemas And intCinemas <> .txtStartAt.Text Then
                    Exit Do
                 End If
                 SkipCinemas = False
                 If .chkPositionEnd.Value And intCinemas = .txtEndAt.Text Then
                    DontDoMore = True
                 End If
                 .StatBar.Panels(1).Text = "Searching for movie..."
                 DoEvents
                 'Do While InStr(1, strText, "1-700", vbTextCompare) = 0 And InStr(1, strText, "700-1", vbTextCompare) = 0
                 '   Line Input #FreeFileNum, strText
                 'Loop
                 Do While Trim(strText) <> ""
                    Line Input #FreeFileNum, strText
                 Loop
                 Do
                    Line Input #FreeFileNum, strText
                 Loop Until Trim(strText) <> ""
            Do
                 .StatBar.Panels(1).Text = "Determining movie..."
                 .StatBar.Panels(3).Text = ""
                 DoEvents
                 strText = NoQuotesString(strText)
                 intMovies = GetMovie(Movies, strText)
                 If intMovies = -1 Then Exit Do 'User chose to skip in frmManual
                 If intMovies = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(3).Text = Movies(intMovies)
                 If DontFilmMore Then Exit Do
                 If SkipMovies Then
                    If intMovies <> frmWorkShop.cbxMovies(0).ItemData(frmWorkShop.cbxMovies(0).ListIndex) Then Exit Do
                 End If
                 SkipMovies = False
                 If frmWorkShop.chkEndMovie.Value = 1 Then
                    If intMovies = frmWorkShop.cbxMovies(1).ItemData(frmWorkShop.cbxMovies(1).ListIndex) Then DontDoMore = True
                 End If
                 .StatBar.Panels(1).Text = "Searching for days & times..."
                 DoEvents
                 Do
                    Line Input #FreeFileNum, strText
                 Loop Until Trim(strText) <> ""
                 .StatBar.Panels(1).Text = "Extracting days & times for movie #" & intMovies & "..."
                 DoEvents
                 ResetArray Times
                 ResetArray Days
                 ResetArray MarTimes
                 ResetVarArray Dates
                 intCount = 1
                 Do While Trim(strText) <> "" And intCount <= MaxTimes
                    strText = DashString(NoDoubleSpaceString(CommaString(strText)))
                    If InStr(1, strText, "'", vbTextCompare) > 0 Then
                        Days(intCount) = Left(strText, InStr(1, strText, ":", vbTextCompare) - 1)
                    Else
                        Days(intCount) = Left(strText, InStr(1, strText, ":", vbTextCompare))
                    End If
                    Times(intCount) = Mid(strText, InStr(1, strText, ":", vbTextCompare) + 1)
                    Line Input #FreeFileNum, strText
                    strText = TabString(strText)
                    intCount = intCount + 1
                 Loop
                 .StatBar.Panels(1).Text = "Filling up the buffer..."
                 DoEvents
                 intCount = 1
                 intI = 1
                 intTimes = 1
                 Do While intCount <= MaxTimes And Days(intCount) <> ""
                   If InStr(1, Days(intCount), "-", vbTextCompare) > 0 Then
                        If InStr(1, Days(intCount), "-", vbTextCompare) = 2 Then
                            If InDates(Dates(), dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 1, 1)), StartDate)) Then
                                Dates(intI) = dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 1, 1)), StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 1, 1)), StartDate)
                            End If
                            If InDates(Dates(), dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)) Then
                                Dates(intI + 1) = dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI + 1) = dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)
                            End If
                            Days(intCount) = CropString(Days(intCount), Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 1, 3))
                        ElseIf InStr(1, Days(intCount), "-", vbTextCompare) >= 3 Then
                            If InDates(Dates(), dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 2, 1)), StartDate)) Then
                                Dates(intI) = dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 2, 1)), StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 2, 1)), StartDate)
                            End If
                            If InDates(Dates(), dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)) Then
                                Dates(intI + 1) = dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI + 1) = dhNextDOW(NewAnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)
                            End If
                            Days(intCount) = CropString(Days(intCount), Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 2, 4))
                        End If
                        If (InStr(1, Days(intCount), "-", vbTextCompare) < 2 And InStr(1, Days(intCount), "-", vbTextCompare) > 0) Or DateDiff("d", StartDate, Dates(intI)) <= -1 Or DateDiff("d", StartDate, Dates(intI + 1)) <= -1 Then
                            'Do
                            '    Dates(intI) = InputBox("Please enter the ""from"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", StartDate)
                            'Loop Until Trim(Dates(intI)) <> "" And IsDate(Dates(intI))
                            'Do
                            '    Dates(intI + 1) = InputBox("Please enter the ""to"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", EndDate)
                            'Loop Until Trim(Dates(intI + 1)) <> "" And IsDate(Dates(intI + 1))
                            frmDates.txtMissing.Text = Days(intCount)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
                            frmDates.Show frmGlobus, vbModal
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
                            Dates(intI) = FormatDateTime(frmDates.DTPicker(1).Value, vbShortDate)
                            Dates(intI + 1) = FormatDateTime(frmDates.DTPicker(2).Value, vbShortDate)
                        End If
                        intI = intI + 2
                        FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                        MarTimes(intTimes) = Trim(Times(intCount))
                        intTimes = intTimes + 1
                        'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ה", vbTextCompare) > 0 Or InStr(1, Days(intCount), "חמישי", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbThursday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbThursday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbThursday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Trim(Times(intCount))
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "שישי", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbFriday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbFriday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbFriday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Trim(Times(intCount))
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "שבת", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbSaturday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbSaturday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbSaturday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Trim(Times(intCount))
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "א", vbTextCompare) > 0 Or InStr(1, Days(intCount), "ראשון", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbSunday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbSunday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbSunday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Trim(Times(intCount))
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ב", vbTextCompare) > 0 Then 'And InStr(1, Days(intCount), "שבת", vbTextCompare) = 0) Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                        If (Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ש" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ת") Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbMonday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbMonday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbMonday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Trim(Times(intCount))
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                        End If
                   End If
                   If InStr(1, Days(intCount), "ג", vbTextCompare) > 0 Or InStr(1, Days(intCount), "שלישי", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbTuesday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbTuesday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbTuesday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Trim(Times(intCount))
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ד", vbTextCompare) > 0 Or InStr(1, Days(intCount), "רביעי", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbWednesday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbWednesday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbWednesday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Trim(Times(intCount))
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                 intCount = intCount + 1
                 Loop
            If Not Simulate Then
                POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                        Dates, Trim(Str(intMovies)), MarTimes
            End If
            intNumMov = intNumMov + 1
            .StatBar.Panels(4).Text = ""
            .StatBar.Panels(1).Text = "Finished uploading data."
            DoEvents
            Line Input #FreeFileNum, strText
            Do While Not EOF(FreeFileNum) And Trim(strText) = ""
                Line Input #FreeFileNum, strText
            Loop
            If InStr(1, strText, MovStartDate & "-" & MovEndDate, vbTextCompare) > 0 Or InStr(1, strText, MovEndDate & "-" & MovStartDate, vbTextCompare) > 0 Or EOF(FreeFileNum) Then Exit Do
            Loop
        intNumCine = intNumCine + 1
        Loop
      ElseIf .cbxType.ListIndex = RavHen Then 'Rav-Hen
        MovStartDate = Format(StartDate, "dd.mm.yy")
        MovEndDate = Format(EndDate, "dd.mm.yy")
        HMP = Seek(FreeFileNum)
        Line Input #FreeFileNum, strText
        strText = RemHenChar(HTMLessString(strText))
        'Don't expect to see a "|" character that implies of a movie entry (and not cinema), unless we haven't began yet (and then, "|" can appear in X-Mailer writings and such) but while we last dealt with a second-column movie (MoviePos=0)
        Do While InStr(1, strText, MovEndDate & " - " & MovStartDate, vbTextCompare) > 0 Or (Not RavHenFirstMovieInBoard And InStr(1, RemHenChar(HTMLessString(strText)), "|", vbTextCompare) > 0 And MoviePos = 0 And Not RavHenHasJustExited)
                 If Not RavHenFirstMovieInBoard And InStr(1, RemHenChar(HTMLessString(strText)), "|", vbTextCompare) > 0 And MoviePos = 0 And Not RavHenHasJustExited Then 'Check for that condition
                    'We must have encountered a movie! Moving to the movie-treating division
                    intNumCine = intNumCine - 1
                    MoviePos = HMP 'It should be in the first row
                    GoTo HenDetermineMovie
                 End If
                 RavHenHasJustExited = False
                 .StatBar.Panels(1).Text = "Found matching dates."
                 DoEvents
                 If DontDoMore Then
                    Exit Do
                 End If
                 Do While InStr(1, ReverseText(RemHenChar(HTMLessString(strText))), "קולנועפון", vbTextCompare) = 0
                    Line Input #FreeFileNum, strText
                 Loop
                 .StatBar.Panels(1).Text = "Determining cinema..."
                 .StatBar.Panels(2).Text = ""
                 DoEvents
                 strText = CutEnglish(ReverseText(RemHenChar(HTMLessString(strText))))
                 strText = Trim(Left(strText, InStr(1, strText, "טל", vbTextCompare) - 1))
                 intCinemas = GetCinema(Cinemas, strText)
                 If intCinemas = -1 Then Exit Do 'User chose to skip in frmManual
                 If intCinemas = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(2).Text = Cinemas(intCinemas)
                 If SkipCinemas And intCinemas <> .txtStartAt.Text Then
                    Exit Do
                 End If
                 SkipCinemas = False
                 If .chkPositionEnd.Value And intCinemas = .txtEndAt.Text Then
                    DontDoMore = True
                 End If
                 .StatBar.Panels(1).Text = "Searching for movie..."
                 DoEvents
                 Do While InStr(1, RemHenChar(HTMLessString(strText)), "------------", vbTextCompare) = 0 And InStr(1, RemHenChar(HTMLessString(strText)), "____________", vbTextCompare) = 0
                    Line Input #FreeFileNum, strText
                 Loop
                 RavHenFirstMovieInBoard = False
                 MoviePos = 0
                 HMP = Seek(FreeFileNum)
                 Line Input #FreeFileNum, strText
                 Do While Trim(strText) = ""
                    HMP = Seek(FreeFileNum)
                    Line Input #FreeFileNum, strText
                 Loop
            Do
                 If MoviePos = 0 Then
                    MoviePos = HMP
                 Else
                    MoviePos = 0
                 End If
HenDetermineMovie:
                 .StatBar.Panels(1).Text = "Determining movie..."
                 .StatBar.Panels(3).Text = ""
                 DoEvents
                 strText = CutEnglish(ReverseText(RemHenChar(HTMLessString(strText))))
                 strText = LeftRight(strText, MoviePos)
                 If strText = "" Then
                    RavHenHasJustExited = True
                    Exit Do
                 End If
                 intMovies = GetMovie(Movies, strText)
                 If intMovies = -1 Then GoTo HenMovieSearch 'User chose to skip in frmManual
                 If intMovies = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(3).Text = Movies(intMovies)
                 RavHenHasJustExited = False
                 If DontFilmMore Then GoTo HenMovieSearch
                 If SkipMovies Then
                    If intMovies <> frmWorkShop.cbxMovies(0).ItemData(frmWorkShop.cbxMovies(0).ListIndex) Then GoTo HenMovieSearch
                 End If
                 SkipMovies = False
                 If frmWorkShop.chkEndMovie.Value = 1 Then
                    If intMovies = frmWorkShop.cbxMovies(1).ItemData(frmWorkShop.cbxMovies(1).ListIndex) Then DontDoMore = True
                 End If
                 .StatBar.Panels(1).Text = "Searching for days & times..."
                 DoEvents
                 Do
                    Line Input #FreeFileNum, strText
                    strText = LeftRight(RemHenChar(HTMLessString(strText)), MoviePos)
                 Loop Until InStr(1, strText, "םוי", vbTextCompare) > 0
                 .StatBar.Panels(1).Text = "Extracting days & times for movie #" & intMovies & "..."
                 DoEvents
                 ResetArray Times
                 ResetArray Days
                 ResetArray MarTimes
                 ResetVarArray Dates
                 intCount = 1
                 Do While Trim(strText) <> "" And intCount <= MaxTimes And InStr(1, strText, "------------", vbTextCompare) = 0 And InStr(1, strText, "___________", vbTextCompare) = 0 And InStr(1, RemHenChar(strText), "<hr>", vbTextCompare) = 0
                    strText = LeftRight(CommaString(RemHenChar(HTMLessString(strText))), MoviePos)
                    Times(intCount) = Left(strText, InStr(1, strText, "  ", vbTextCompare) - 1)
                    Days(intCount) = Mid(strText, InStr(1, strText, "  ", vbTextCompare) + 1, InStr(1, strText, "םוי", vbTextCompare) - InStr(1, strText, "  ", vbTextCompare)) 'Trim out the "םוי"
                    Line Input #FreeFileNum, strText
                    strText = LeftRight(CommaString(RemHenChar(HTMLessString(strText))), MoviePos)
                    Do While InStr(1, strText, "םוי", vbTextCompare) = 0 And strText <> "" And InStr(1, strText, "------------", vbTextCompare) = 0 And InStr(1, strText, "___________", vbTextCompare) = 0 And InStr(1, RemHenChar(strText), "<hr>", vbTextCompare) = 0
                        Times(intCount) = Times(intCount) & " " & NoDoubleSpaceString(strText)
                        Line Input #FreeFileNum, strText
                        strText = LeftRight(CommaString(RemHenChar(HTMLessString(strText))), MoviePos)
                    Loop
                    intCount = intCount + 1
                 Loop
                 .StatBar.Panels(1).Text = "Filling up the buffer..."
                 DoEvents
                 intCount = 1
                 intI = 1
                 intTimes = 1
                 Do While intCount <= MaxTimes And Days(intCount) <> ""
                   If InStr(1, Days(intCount), "-", vbTextCompare) > 0 Then
                        'Do
                        '    Dates(intI) = InputBox("Please enter the ""from"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", StartDate)
                        'Loop Until Trim(Dates(intI)) <> "" And IsDate(Dates(intI))
                        'Do
                        '    Dates(intI + 1) = InputBox("Please enter the ""to"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", EndDate)
                        'Loop Until Trim(Dates(intI + 1)) <> "" And IsDate(Dates(intI + 1))
                        frmDates.txtMissing.Text = Days(intCount)
                        If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
                        frmDates.Show frmGlobus, vbModal
                        Dates(intI) = FormatDateTime(frmDates.DTPicker(1).Value, vbShortDate)
                        Dates(intI + 1) = FormatDateTime(frmDates.DTPicker(2).Value, vbShortDate)
                        If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
                        intI = intI + 2
                        FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                        MarTimes(intTimes) = Times(intCount)
                        intTimes = intTimes + 1
                        'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ה", vbTextCompare) > 0 Then
                            Dates(intI) = dhNextDOW(vbThursday, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ו", vbTextCompare) > 0 Then
                            Dates(intI) = dhNextDOW(vbFriday, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "תבש", vbTextCompare) > 0 Then
                            Dates(intI) = dhNextDOW(vbSaturday, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "א", vbTextCompare) > 0 Then
                            Dates(intI) = dhNextDOW(vbSunday, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ב", vbTextCompare) > 0 Then 'And InStr(1, Days(intCount), "תבש", vbTextCompare) = 0 Then
                        If Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ש" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ת" Then
                            Dates(intI) = dhNextDOW(vbMonday, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                        End If
                   End If
                   If InStr(1, Days(intCount), "ג", vbTextCompare) > 0 Then
                            Dates(intI) = dhNextDOW(vbTuesday, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ד", vbTextCompare) > 0 Then
                            Dates(intI) = dhNextDOW(vbWednesday, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                 intCount = intCount + 1
                 Loop
            If Not Simulate Then
                POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                        Dates, Trim(Str(intMovies)), MarTimes
            End If
            intNumMov = intNumMov + 1
HenMovieSearch:
            .StatBar.Panels(4).Text = ""
            .StatBar.Panels(1).Text = "Finished uploading data."
            DoEvents
            If MoviePos <> 0 Then
                Seek #FreeFileNum, MoviePos
                Line Input #FreeFileNum, strText
            Else
                HMP = Seek(FreeFileNum)
                Line Input #FreeFileNum, strText
                strText = RemHenChar(HTMLessString(strText))
                Do While Not EOF(FreeFileNum) And (LeftRight(strText, 0) = "" And LeftRight(strText, 1) = "") Or (InStr(1, strText, ":", vbTextCompare) > 0 And HaveHour(strText)) And InStr(1, strText, "------------", vbTextCompare) = 0 And InStr(1, strText, "____________", vbTextCompare) = 0 And InStr(1, strText, "ISRAEL THEATERS LTD.", vbTextCompare) = 0
                    HMP = Seek(FreeFileNum)
                    Line Input #FreeFileNum, strText
                    strText = RemHenChar(HTMLessString(strText))
                Loop
                If InStr(1, strText, "------------", vbTextCompare) > 0 Or InStr(1, strText, "____________", vbTextCompare) > 0 Or InStr(1, strText, "ISRAEL THEATERS LTD.", vbTextCompare) > 0 Or EOF(FreeFileNum) Then Exit Do
            End If
            Loop
        intNumCine = intNumCine + 1
        Loop
      ElseIf .cbxType.ListIndex = London Then 'London
        MovStartDate = Format(StartDate, "d.m")
        MovEndDate = Format(EndDate, "d.m")
        .StatBar.Panels(2).Text = ""
        .StatBar.Panels(1).Text = "Cinema is automatically set."
        DoEvents
        intCinemas = 77
        .StatBar.Panels(2).Text = Cinemas(intCinemas)
        .StatBar.Panels(1).Text = "Searching for matching dates..."
        DoEvents
        Do
            Line Input #FreeFileNum, strText
        Loop Until InStr(1, strText, MovEndDate & "-" & MovStartDate, vbTextCompare) > 0 Or InStr(1, strText, MovStartDate & " - " & MovEndDate, vbTextCompare) > 0
        .StatBar.Panels(1).Text = "Found matching dates; searching for movie..."
        DoEvents
        Do
            Line Input #FreeFileNum, strText
        Loop Until Trim(strText) <> ""
        Do
            .StatBar.Panels(1).Text = "Determining movie..."
            .StatBar.Panels(3).Text = ""
            DoEvents
            strText = NoQuotesString(NoDoubleSpaceString(strText))
            intMovies = GetMovie(Movies, strText)
            If intMovies = -1 Then Exit Do 'User chose to skip in frmManual
            If intMovies = -2 Then Err.Raise 5250, , "User-generated emergency stop"
            .StatBar.Panels(3).Text = Movies(intMovies)
            If DontFilmMore Then Exit Do
            If SkipMovies Then
                If intMovies <> frmWorkShop.cbxMovies(0).ItemData(frmWorkShop.cbxMovies(0).ListIndex) Then Exit Do
            End If
            SkipMovies = False
            If frmWorkShop.chkEndMovie.Value = 1 Then
                If intMovies = frmWorkShop.cbxMovies(1).ItemData(frmWorkShop.cbxMovies(1).ListIndex) Then DontDoMore = True
            End If
            .StatBar.Panels(1).Text = "Searching for days & times..."
            DoEvents
            Do
                Line Input #FreeFileNum, strText
            Loop Until Trim(strText) <> ""
            .StatBar.Panels(1).Text = "Extracting days & times for movie #" & intMovies & "..."
            DoEvents
            ResetArray Times
            ResetArray Days
            intCount = 1
            Do While Trim(strText) <> "" And intCount <= MaxTimes
                strText = DashString(NoDoubleSpaceString(CommaString(strText)))
                Days(intCount) = Left(strText, InStr(1, strText, ":", vbTextCompare))
                Times(intCount) = Mid(strText, InStr(1, strText, ":", vbTextCompare) + 1)
                Line Input #FreeFileNum, strText
                intCount = intCount + 1
            Loop
            ResetArray MarTimes
            ResetVarArray Dates
            .StatBar.Panels(1).Text = "Filling up the buffer..."
            DoEvents
            intCount = 1
            intI = 1
            intTimes = 1
            Do While intCount <= MaxTimes And Days(intCount) <> ""
                   If InStr(1, Days(intCount), "-", vbTextCompare) > 0 Then
                        If InStr(1, Days(intCount), "-", vbTextCompare) = 2 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 1, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)
                        ElseIf InStr(1, Days(intCount), "-", vbTextCompare) > 3 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 2, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)
                        End If
                        If InStr(1, Days(intCount), "-", vbTextCompare) < 2 Or DateDiff("d", StartDate, Dates(intI)) = -1 Or DateDiff("d", StartDate, Dates(intI + 1)) = -1 Then
                            'Do
                            '    Dates(intI) = InputBox("Please enter the ""from"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", StartDate)
                            'Loop Until Trim(Dates(intI)) <> "" And IsDate(Dates(intI))
                            'Do
                            '    Dates(intI + 1) = InputBox("Please enter the ""to"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", EndDate)
                            'Loop Until Trim(Dates(intI + 1)) <> "" And IsDate(Dates(intI + 1))
                            frmDates.txtMissing.Text = Days(intCount)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
                            frmDates.Show frmGlobus, vbModal
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
                            Dates(intI) = FormatDateTime(frmDates.DTPicker(1).Value, vbShortDate)
                            Dates(intI + 1) = FormatDateTime(frmDates.DTPicker(2).Value, vbShortDate)
                        End If
                        intI = intI + 2
                        FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                        MarTimes(intTimes) = Times(intCount)
                        intTimes = intTimes + 1
                        'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ה", vbTextCompare) > 0 Or InStr(1, Days(intCount), "חמישי", vbTextCompare) > 0 Then
                            Dates(intI) = StartDate
                            Dates(intI + 1) = StartDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "שישי", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 1, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "שבת", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 2, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "א", vbTextCompare) > 0 Or InStr(1, Days(intCount), "ראשון", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 3, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ב", vbTextCompare) > 0 Then 'And InStr(1, Days(intCount), "שבת", vbTextCompare) = 0) Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                        If (Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ש" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ת") Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 4, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                        End If
                   End If
                   If InStr(1, Days(intCount), "ג", vbTextCompare) > 0 Or InStr(1, Days(intCount), "שלישי", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 5, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                   If InStr(1, Days(intCount), "ד", vbTextCompare) > 0 Or InStr(1, Days(intCount), "רביעי", vbTextCompare) > 0 Then
                            Dates(intI) = EndDate
                            Dates(intI + 1) = EndDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                            'POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                                Dates, Trim(Str(intMovies)), Trim(Times(intCount))
                   End If
                 intCount = intCount + 1
                 Loop
            If Not Simulate Then
                POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                        Dates, Trim(Str(intMovies)), MarTimes
            End If
            intNumMov = intNumMov + 1
            .StatBar.Panels(4).Text = ""
            .StatBar.Panels(1).Text = "Finished uploading data."
            DoEvents
            Do
                Line Input #FreeFileNum, strText
            Loop Until Trim(strText) <> "" Or EOF(FreeFileNum)
            If EOF(FreeFileNum) Then Exit Do
            Loop
        intNumCine = intNumCine + 1
      ElseIf .cbxType.ListIndex = Haifa Then 'Haifas
        MovStartDate = Format(StartDate, "dd.mm.yyyy")
        MovEndDate = Format(EndDate, "dd.mm.yyyy")
        .StatBar.Panels(1).Text = "Searching for cinema..."
        .StatBar.Panels(2).Text = ""
        DoEvents
        If DontDoMore Then
           Exit Do
        End If
        Do While InStr(1, strText, "מאת", vbTextCompare) = 0
            If InStr(1, strText, "בית גבריאל", vbTextCompare) > 0 Then Exit Do
            Line Input #FreeFileNum, strText
        Loop
HaifaStartCinema:
        If InStr(1, strText, "בית גבריאל", vbTextCompare) > 0 Then
            intCinemas = GetCinema(Cinemas, "בית גבריאל")
        Else
            strText = Trim(Mid(strText, InStr(1, strText, ":") + 1))
            intCinemas = GetCinema(Cinemas, strText)
        End If
        If intCinemas = -1 Then Exit Do 'User chose to skip in frmManual
        If intCinemas = -2 Then Err.Raise 5250, , "User-generated emergency stop"
        .StatBar.Panels(2).Text = Cinemas(intCinemas)
        If SkipCinemas And intCinemas <> .txtStartAt.Text Then
            Exit Do
        End If
        SkipCinemas = False
        If .chkPositionEnd.Value And intCinemas = .txtEndAt.Text Then
            DontDoMore = True
        End If
        .StatBar.Panels(1).Text = "Searching for matching dates..."
        DoEvents
        Do
            If InStr(1, strText, "בית גבריאל", vbTextCompare) = 0 Then Line Input #FreeFileNum, strText
        Loop Until InStr(1, strText, Format(EndDate, "dd.mm.yy") & " לבין " & Format(StartDate, "dd.mm.yy"), vbTextCompare) > 0 Or InStr(1, strText, Format(StartDate, "dd.mm.yy") & " לבין " & Format(EndDate, "dd.mm.yy"), vbTextCompare) > 0 Or InStr(1, strText, MovEndDate & " לבין " & MovStartDate, vbTextCompare) > 0 Or InStr(1, strText, MovStartDate & " לבין " & MovEndDate, vbTextCompare) > 0 Or InStr(1, strText, "מאת", vbTextCompare) > 0 Or (InStr(1, strText, "בית גבריאל", vbTextCompare) > 0 And InStr(1, strText, "-", vbTextCompare) > 0)
        If InStr(1, strText, "מאת", vbTextCompare) > 0 Then GoTo HaifaStartCinema
        .StatBar.Panels(1).Text = "Found matching dates; searching for movie..."
        DoEvents
HaifaStartMovie:
        Do
            Line Input #FreeFileNum, strText
        Loop Until InStr(1, strText, "הסרט:") > 0 Or InStr(1, strText, "מאת", vbTextCompare) > 0 Or (InStr(1, strText, """", vbTextCompare) > 0 And InStr(1, strText, "בית גבריאל", vbTextCompare) = 0)
        If InStr(1, strText, "מאת", vbTextCompare) > 0 Then GoTo HaifaStartCinema
        Do
            .StatBar.Panels(1).Text = "Determining movie..."
            .StatBar.Panels(3).Text = ""
            DoEvents
            strText = NoQuotesString(NoDoubleSpaceString(CropString(strText, "הסרט:")))
            intMovies = GetMovie(Movies, strText)
            If intMovies = -1 Then GoTo HaifaStartMovie 'User chose to skip in frmManual
            If intMovies = -2 Then Err.Raise 5250, , "User-generated emergency stop"
            .StatBar.Panels(3).Text = Movies(intMovies)
            If DontFilmMore Then Exit Do
            If SkipMovies Then
                If intMovies <> frmWorkShop.cbxMovies(0).ItemData(frmWorkShop.cbxMovies(0).ListIndex) Then GoTo HaifaStartMovie
            End If
            SkipMovies = False
            If frmWorkShop.chkEndMovie.Value = 1 Then
                If intMovies = frmWorkShop.cbxMovies(1).ItemData(frmWorkShop.cbxMovies(1).ListIndex) Then DontDoMore = True
            End If
            .StatBar.Panels(1).Text = "Searching for days & times..."
            DoEvents
            Do
                Line Input #FreeFileNum, strText
            Loop Until HaveHour(Trim(strText))
            .StatBar.Panels(1).Text = "Extracting days & times for movie #" & intMovies & "..."
            DoEvents
            ResetArray Times
            ResetArray Days
            intCount = 1
            Do While Trim(strText) <> "" And intCount <= MaxTimes And HaveHour(Trim(strText))
                strText = CommaString(strText)
                Days(intCount) = DashString(Mid(strText, InStr(1, strText, "ם", vbTextCompare) + 2, InStr(InStr(1, strText, "ם", vbTextCompare) + 2, strText, "  ", vbTextCompare)))
                strText = NoDoubleSpaceString(TabString(strText)) 'Must be here, and not anywhere above
                'Times(intCount) = Mid(strText, InStr(1, strText, ":", vbTextCompare) + 1, Len(strText) - FindLast(strText, " ") + 6)
                Times(intCount) = Mid(strText, WhereHour(strText))
                'If Right(Times(intCount), 7) = " יומית." Then Times(intCount) = Mid(Times(intCount), 1, Len(Times(intCount)) - 7)
                Times(intCount) = CropString(Times(intCount), "יומית")
                Times(intCount) = CropString(Times(intCount), "חצות")
                Times(intCount) = CropString(Times(intCount), "צהריים")
                Times(intCount) = CropString(Times(intCount), "צהרים")
                Times(intCount) = CropString(Times(intCount), ".")
                Line Input #FreeFileNum, strText
                If Not HaveHour(strText) Then Line Input #FreeFileNum, strText 'Try again
                intCount = intCount + 1
            Loop
            ResetArray MarTimes
            ResetVarArray Dates
            .StatBar.Panels(1).Text = "Filling up the buffer..."
            DoEvents
            intCount = 1
            intI = 1
            intTimes = 1
            Do While intCount <= MaxTimes And Days(intCount) <> ""
                   'Seems to me that there's a difference between Moria's dash and the regular dash; we check them both
                   If InStr(1, Days(intCount), " – ", vbTextCompare) > 0 Or InStr(1, Days(intCount), " - ", vbTextCompare) > 0 Then
                        If InStr(1, Days(intCount), " – ", vbTextCompare) = 2 Or InStr(1, Days(intCount), " - ", vbTextCompare) = 2 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) - 1, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) + 3, 1)), StartDate)
                        ElseIf InStr(1, Days(intCount), " – ", vbTextCompare) >= 3 Or InStr(1, Days(intCount), " - ", vbTextCompare) >= 3 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) - 2, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) + 3, 1)), StartDate)
                        End If
                        If InStr(1, Days(intCount), "-", vbTextCompare) < 2 Or DateDiff("d", StartDate, Dates(intI)) = -1 Or DateDiff("d", StartDate, Dates(intI + 1)) = -1 Then
                            'Do
                            '    Dates(intI) = InputBox("Please enter the ""from"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", StartDate)
                            'Loop Until Trim(Dates(intI)) <> "" And IsDate(Dates(intI))
                            'Do
                            '    Dates(intI + 1) = InputBox("Please enter the ""to"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", EndDate)
                            'Loop Until Trim(Dates(intI + 1)) <> "" And IsDate(Dates(intI + 1))
                            frmDates.txtMissing.Text = Days(intCount)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
                            frmDates.Show frmGlobus, vbModal
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
                            Dates(intI) = FormatDateTime(frmDates.DTPicker(1).Value, vbShortDate)
                            Dates(intI + 1) = FormatDateTime(frmDates.DTPicker(2).Value, vbShortDate)
                        End If
                        intI = intI + 2
                        FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                        MarTimes(intTimes) = Times(intCount)
                        intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ה", vbTextCompare) > 0 Or InStr(1, Days(intCount), "חמישי", vbTextCompare) > 0 Then
                            Dates(intI) = StartDate
                            Dates(intI + 1) = StartDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "שישי", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 1, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "שבת", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 2, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "א", vbTextCompare) > 0 Or InStr(1, Days(intCount), "ראשון", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 3, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   'If (InStr(1, Days(intCount), "ב", vbTextCompare) > 0 And InStr(1, Days(intCount), "שבת", vbTextCompare) = 0 And InStr(1, Days(intCount), "רביעי", vbTextCompare) = 0) Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                   If InStr(1, Days(intCount), "ב", vbTextCompare) > 0 Then
                        If (Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ש" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ת" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "י" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ר") Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 4, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                        End If
                   End If
                   'The "Lishi" is there because of trim problems when extracting the days from Beit-Gabriel boards
                   If InStr(1, Days(intCount), "ג", vbTextCompare) > 0 Or InStr(1, Days(intCount), "שלישי", vbTextCompare) > 0 Or InStr(1, Days(intCount), "לישי", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 5, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ד", vbTextCompare) > 0 Or InStr(1, Days(intCount), "רביעי", vbTextCompare) > 0 Then
                            Dates(intI) = EndDate
                            Dates(intI + 1) = EndDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                 intCount = intCount + 1
                 Loop
            If Not Simulate Then
                POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                        Dates, Trim(Str(intMovies)), MarTimes
            End If
            intNumMov = intNumMov + 1
            .StatBar.Panels(4).Text = ""
            .StatBar.Panels(1).Text = "Finished uploading data."
            DoEvents
            Do While InStr(1, strText, "הסרט:") = 0 And InStr(1, strText, "מאת") = 0 And InStr(1, strText, "מופעים", vbTextCompare) = 0 And Not EOF(FreeFileNum)
                If (InStr(1, strText, """", vbTextCompare) > 0 And InStr(1, strText, "04", vbTextCompare) = 0) Then Exit Do
                Line Input #FreeFileNum, strText
            Loop
            If EOF(FreeFileNum) Then Exit Do
            If InStr(1, strText, "מאת") > 0 Or InStr(1, strText, "מופעים") > 0 Then Exit Do
            Loop
        intNumCine = intNumCine + 1
      ElseIf .cbxType.ListIndex = Jerusalem Then 'Jerusalem Theatre
        MovStartDate = Format(StartDate, "d.m.yy")
        MovEndDate = Format(EndDate, "d.m.yy")
        .StatBar.Panels(1).Text = "Searching for cinema..."
        .StatBar.Panels(2).Text = ""
        DoEvents
        If DontDoMore Then
           Exit Do
        End If
        Do
            Line Input #FreeFileNum, strText
        'This poor line is wrong (wrote about it to LJ) - Loop Until Trim(strText) <> "" And Not isdate(strtext)
        Loop Until Trim(strText) <> "" And InStr(1, strText, ".", vbTextCompare) = 0
        intCinemas = GetCinema(Cinemas, strText)
        If intCinemas = -1 Then Exit Do 'User chose to skip in frmManual
        If intCinemas = -2 Then Err.Raise 5250, , "User-generated emergency stop"
        .StatBar.Panels(2).Text = Cinemas(intCinemas)
        If SkipCinemas And intCinemas <> .txtStartAt.Text Then
            Exit Do
        End If
        SkipCinemas = False
        If .chkPositionEnd.Value And intCinemas = .txtEndAt.Text Then
            DontDoMore = True
        End If
        .StatBar.Panels(1).Text = "Searching for matching dates..."
        DoEvents
        Do
            Line Input #FreeFileNum, strText
        Loop Until InStr(1, NoDoubleSpaceString(strText), MovEndDate & " עד יום ד' " & MovStartDate, vbTextCompare) > 0 Or InStr(1, NoDoubleSpaceString(strText), MovStartDate & " עד יום ד' " & MovEndDate, vbTextCompare) > 0
        'Do 'Yes, twice! The line that has the dates appears twice
        '    Line Input #FreeFileNum, strText
        'Loop Until InStr(1, strText, MovEndDate & " עד יום ד' " & MovStartDate, vbTextCompare) > 0 Or InStr(1, strText, MovStartDate & " עד יום ד' " & MovEndDate, vbTextCompare) > 0
        .StatBar.Panels(1).Text = "Found matching dates; searching for movie..."
        DoEvents
        Do
            Line Input #FreeFileNum, strText
        Loop Until Trim(strText) <> "" And (InStr(1, strText, """", vbTextCompare) > 0 Or HaveEnglish(strText))
        Do
            .StatBar.Panels(1).Text = "Determining movie..."
            strText = Trim(TabString(strText))
            .StatBar.Panels(3).Text = ""
            DoEvents
            .StatBarIdle.Panels(3).Tag = FindLast(strText, """")
            If .StatBarIdle.Panels(3).Tag <> "0" Then
                .cmdExit.Tag = CutEnglish(NoQuotesString(Mid(strText, FindLast(strText, """", Val(.StatBarIdle.Panels(3).Tag) - 1), Len(strText) - Val(.StatBarIdle.Panels(3).Tag) - 1)))
            Else
                .cmdExit.Tag = Trim(CutEnglish(strText))
            End If
            intMovies = GetMovie(Movies, .cmdExit.Tag)
            If intMovies = -1 Then Exit Do 'User chose to skip in frmManual
            If intMovies = -2 Then Err.Raise 5250, , "User-generated emergency stop"
            .StatBar.Panels(3).Text = Movies(intMovies)
            If DontFilmMore Then Exit Do
            If SkipMovies Then
                If intMovies <> frmWorkShop.cbxMovies(0).ItemData(frmWorkShop.cbxMovies(0).ListIndex) Then Exit Do
            End If
            SkipMovies = False
            If frmWorkShop.chkEndMovie.Value = 1 Then
                If intMovies = frmWorkShop.cbxMovies(1).ItemData(frmWorkShop.cbxMovies(1).ListIndex) Then DontDoMore = True
            End If
            .StatBar.Panels(1).Text = "Searching for days & times for movie #" & intMovies & "..."
            DoEvents
            Do While Not HaveHour(strText)
                Line Input #FreeFileNum, strText
            Loop
            .StatBar.Panels(1).Text = "Extracting days & times for movie #" & intMovies & "..."
            DoEvents
            ResetArray Times
            ResetArray Days
            intCount = 1
            Do While Trim(strText) <> "" And intCount <= MaxTimes And HaveHour(strText)
                'strText = TabString(DashString(CommaString(strText)))
                'Days(intCount) = Mid(strText, InStr(1, strText, "ם", vbTextCompare) + 2, InStr(InStr(1, strText, "ם", vbTextCompare) + 2, strText, "  ", vbTextCompare))
                'strText = NoDoubleSpaceString(strText) 'Must be here, and not anywhere above
                'Times(intCount) = Trim(Right(strText, Len(strText) - FindLast(strText, """")))
                strText = DashString(NoDoubleSpaceString(CommaString(strText)))
                .cbxType.Tag = WhereHour(strText)
                Days(intCount) = Trim(Left(strText, Val(.cbxType.Tag) - 1))
                Days(intCount) = CropString(Days(intCount), "בשעות")
                Times(intCount) = Trim(Mid(strText, Val(.cbxType.Tag)))
                Line Input #FreeFileNum, strText
                If Not HaveHour(strText) Then Line Input #FreeFileNum, strText 'Try again
                intCount = intCount + 1
                'StatBarIdle.Panels(3).Tag = FindLast(strText, vbTab)
                'fraSource.Tag = NoQuotesString(Mid(strText, FindLast(strText, """", Val(StatBarIdle.Panels(3).Tag) - 1), Len(strText) - Val(StatBarIdle.Panels(3).Tag) - 1))
                'If cmdExit.Tag <> fraSource.Tag Then
                '    fraSource.Tag = "1"
                '    Exit Do
                'End If
            Loop
            ResetArray MarTimes
            ResetVarArray Dates
            .StatBar.Panels(1).Text = "Filling up the buffer..."
            DoEvents
            intCount = 1
            intI = 1
            intTimes = 1
            Do While intCount <= MaxTimes And Days(intCount) <> ""
                   If InStr(1, Days(intCount), " - ", vbTextCompare) > 0 Then
                        If InStr(1, Days(intCount), " - ", vbTextCompare) = 2 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) - 1, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) + 3, 1)), StartDate)
                        ElseIf InStr(1, Days(intCount), " - ", vbTextCompare) > 3 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) - 2, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) + 3, 1)), StartDate)
                        End If
                        If InStr(1, Days(intCount), "-", vbTextCompare) < 2 Or DateDiff("d", StartDate, Dates(intI)) = -1 Or DateDiff("d", StartDate, Dates(intI + 1)) = -1 Then
                            'Do
                            '    Dates(intI) = InputBox("Please enter the ""from"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", StartDate)
                            'Loop Until Trim(Dates(intI)) <> "" And IsDate(Dates(intI))
                            'Do
                            '    Dates(intI + 1) = InputBox("Please enter the ""to"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", EndDate)
                            'Loop Until Trim(Dates(intI + 1)) <> "" And IsDate(Dates(intI + 1))
                            frmDates.txtMissing.Text = Days(intCount)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
                            frmDates.Show frmGlobus, vbModal
                            Dates(intI) = FormatDateTime(frmDates.DTPicker(1).Value, vbShortDate)
                            Dates(intI + 1) = FormatDateTime(frmDates.DTPicker(2).Value, vbShortDate)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
                        End If
                        intI = intI + 2
                        FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                        MarTimes(intTimes) = Times(intCount)
                        intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ה", vbTextCompare) > 0 Or InStr(1, Days(intCount), "חמישי", vbTextCompare) > 0 Then
                            Dates(intI) = StartDate
                            Dates(intI + 1) = StartDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "שישי", vbTextCompare) > 0 Or (InStr(1, Days(intCount), "ו", vbTextCompare) > 0 And InStr(1, Days(intCount), "ראשון", vbTextCompare) = 0 And InStr(1, Days(intCount), "מוצ""ש", vbTextCompare) = 0) Then
                            Dates(intI) = DateAdd("d", 1, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "שבת", vbTextCompare) > 0 Or InStr(1, Days(intCount), "מוצ""ש", vbTextCompare) > 0 Or InStr(1, Days(intCount), "מוצש", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 2, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "א", vbTextCompare) > 0 Or InStr(1, Days(intCount), "ראשון", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 3, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ב", vbTextCompare) > 0 Then 'And InStr(1, Days(intCount), "שבת", vbTextCompare) = 0) Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                        If (Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ש" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ת") Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 4, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                        End If
                   End If
                   If InStr(1, Days(intCount), "ג", vbTextCompare) > 0 Or InStr(1, Days(intCount), "שלישי", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 5, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ד", vbTextCompare) > 0 Or InStr(1, Days(intCount), "רביעי", vbTextCompare) > 0 Then
                            Dates(intI) = EndDate
                            Dates(intI + 1) = EndDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                 intCount = intCount + 1
                 Loop
            If Not Simulate Then
                POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                        Dates, Trim(Str(intMovies)), MarTimes
            End If
            intNumMov = intNumMov + 1
            .StatBar.Panels(4).Text = ""
            .StatBar.Panels(1).Text = "Finished uploading data."
            DoEvents
            'If fraSource.Tag <> "1" Then
                Do While Trim(strText) = "" And InStr(1, strText, """", vbTextCompare) = 0 And (InStr(1, strText, MovEndDate & " עד יום ד' " & MovStartDate, vbTextCompare) = 0 Or InStr(1, strText, MovStartDate & " עד יום ד' " & MovEndDate, vbTextCompare) = 0) And Not EOF(FreeFileNum)
                    Line Input #FreeFileNum, strText
                Loop
            'End If
            If InStr(1, strText, MovEndDate & " עד יום ד' " & MovStartDate, vbTextCompare) > 0 Or InStr(1, strText, MovStartDate & " עד יום ד' " & MovEndDate, vbTextCompare) > 0 Then Seek #FreeFileNum, Seek(FreeFileNum) * Seek(FreeFileNum)
            If EOF(FreeFileNum) Then Exit Do
            Loop
        intNumCine = intNumCine + 1
    ElseIf .cbxType.ListIndex = Compiled Then 'Specially-compiled
        MovStartDate = Format(StartDate, "d/m/yyyy")
        MovEndDate = Format(EndDate, "d/m/yyyy")
        Line Input #FreeFileNum, strText
        strText = Trim(strText)
        If InStr(1, strText, MovStartDate, vbTextCompare) > 0 Or intCinemas <> 0 Then 'Also check for not first time coming here
            Do
                Line Input #FreeFileNum, strText
            Loop Until Trim(strText) <> ""
            Do
                 'Protect from end-of-file error which may come on later runs of this do-loop
                 If EOF(FreeFileNum) Then Exit Do
                 .StatBar.Panels(1).Text = "Found matching dates; Determining cinema..."
                 .StatBar.Panels(2).Text = ""
                 DoEvents
                 If DontDoMore Then
                    Exit Do
                 End If
                 intCinemas = GetCinema(Cinemas, Trim(strText))
                 If intCinemas = -1 Then Exit Do 'User chose to skip in frmManual
                 If intCinemas = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(2).Text = Cinemas(intCinemas)
                 If SkipCinemas And intCinemas <> .txtStartAt.Text Then
                    Exit Do
                 End If
                 SkipCinemas = False
                 If .chkPositionEnd.Value And intCinemas = .txtEndAt.Text Then
                    DontDoMore = True
                 End If
                 .StatBar.Panels(1).Text = "Searching for movie..."
                 DoEvents
                 Do
                    Line Input #FreeFileNum, strText
                 Loop Until Trim(strText) <> ""
            Do
                 .StatBar.Panels(1).Text = "Determining movie..."
                 .StatBar.Panels(3).Text = ""
                 DoEvents
                 strText = CropString(NoDoubleSpaceString(NoQuotesString(strText)), "סרט:")
                 intMovies = GetMovie(Movies, CropString(strText, vbTab))
                 If intMovies = -1 Then Exit Do 'User chose to skip in frmManual
                 If intMovies = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(3).Text = Movies(intMovies)
                 If DontFilmMore Then Exit Do
                 If SkipMovies Then
                    If intMovies <> frmWorkShop.cbxMovies(0).ItemData(frmWorkShop.cbxMovies(0).ListIndex) Then Exit Do
                 End If
                 SkipMovies = False
                 If frmWorkShop.chkEndMovie.Value = 1 Then
                    If intMovies = frmWorkShop.cbxMovies(1).ItemData(frmWorkShop.cbxMovies(1).ListIndex) Then DontDoMore = True
                 End If
                 .StatBar.Panels(1).Text = "Searching for days & times..."
                 DoEvents
                 Do
                    Line Input #FreeFileNum, strText
                 Loop Until Trim(strText) <> "" And HaveHour(strText)
                 .StatBar.Panels(1).Text = "Extracting days & times for movie #" & intMovies & "..."
                 DoEvents
                 ResetArray Times
                 ResetArray Days
                 ResetArray MarTimes
                 ResetVarArray Dates
                 intCount = 1
                 Do While Trim(strText) <> "" And intCount <= MaxTimes And HaveHour(strText)
                    strText = DashString(NoDoubleSpaceString(CommaString(strText)))
                    .cbxType.Tag = WhereHour(strText)
                    'If InStr(1, strText, ":", vbTextCompare) > Val(cbxType.Tag) Then 'No colon between the day and the time
                        Days(intCount) = Trim(Left(strText, Val(.cbxType.Tag) - 1))
                    'Else
                    '    Days(intCount) = Trim(Left(strText, InStr(1, strText, ":", vbTextCompare) - 1))
                    'End If
                    Days(intCount) = CropString(Days(intCount), "בשעות")
                    Times(intCount) = Trim(Mid(strText, Val(.cbxType.Tag)))
                    Line Input #FreeFileNum, strText
                    Do While HaveHour(Left(Trim(strText), 5))
                        Times(intCount) = Times(intCount) & " " & NoDoubleSpaceString(CommaString(strText))
                        Line Input #FreeFileNum, strText
                    Loop
                    intCount = intCount + 1
                 Loop
                 .StatBar.Panels(1).Text = "Filling up the buffer..."
                 DoEvents
                 intCount = 1
                 intI = 1
                 intTimes = 1
                 Do While intCount <= MaxTimes And Days(intCount) <> ""
                   If InStr(1, Days(intCount), "-", vbTextCompare) > 0 Then
                        If InStr(1, Days(intCount), " - ", vbTextCompare) = 2 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) - 1, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) + 3, 1)), StartDate)
                        ElseIf InStr(1, Days(intCount), " - ", vbTextCompare) >= 3 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) - 2, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) + 3, 1)), StartDate)
                        ElseIf InStr(1, Days(intCount), "-", vbTextCompare) = 2 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 1, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)
                        ElseIf InStr(1, Days(intCount), "-", vbTextCompare) >= 3 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 2, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)
                        End If
                        If InStr(1, Days(intCount), "-", vbTextCompare) < 2 Or DateDiff("d", StartDate, Dates(intI)) = -1 Or DateDiff("d", StartDate, Dates(intI + 1)) = -1 Then
                            'Do
                            '    Dates(intI) = InputBox("Please enter the ""from"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", StartDate)
                            'Loop Until Trim(Dates(intI)) <> "" And IsDate(Dates(intI))
                            'Do
                            '    Dates(intI + 1) = InputBox("Please enter the ""to"" date for " & Trim(Days(intCount)) & " (dd/mm/yyyy):", "Cannot analyze date", EndDate)
                            'Loop Until Trim(Dates(intI + 1)) <> "" And IsDate(Dates(intI + 1))
                            frmDates.txtMissing.Text = Days(intCount)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
                            frmDates.Show frmGlobus, vbModal
                            Dates(intI) = FormatDateTime(frmDates.DTPicker(1).Value, vbShortDate)
                            Dates(intI + 1) = FormatDateTime(frmDates.DTPicker(2).Value, vbShortDate)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
                        End If
                        intI = intI + 2
                        FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                        MarTimes(intTimes) = Times(intCount)
                        intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ה", vbTextCompare) > 0 Or InStr(1, Days(intCount), "חמישי", vbTextCompare) > 0 Then
                            Dates(intI) = StartDate
                            Dates(intI + 1) = StartDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "שישי", vbTextCompare) > 0 Or (InStr(1, Days(intCount), "ו", vbTextCompare) > 0 And InStr(1, Days(intCount), "ראשון", vbTextCompare) = 0) Then
                            Dates(intI) = DateAdd("d", 1, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   'A very long check!
                   If InStr(1, Days(intCount), "שבת", vbTextCompare) > 0 Or (InStr(1, Days(intCount), "ש", vbTextCompare) > 0 And InStr(1, Days(intCount), "חמישי", vbTextCompare) = 0 And InStr(1, Days(intCount), "שישי", vbTextCompare) = 0 And InStr(1, Days(intCount), "ראשון", vbTextCompare) = 0 And InStr(1, Days(intCount), "שני", vbTextCompare) = 0 And InStr(1, Days(intCount), "שלישי", vbTextCompare) = 0) Then
                            Dates(intI) = DateAdd("d", 2, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "א", vbTextCompare) > 0 Or InStr(1, Days(intCount), "ראשון", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 3, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ב", vbTextCompare) > 0 Then 'And InStr(1, Days(intCount), "שבת", vbTextCompare) = 0) Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                        If (Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ש" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ת") Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 4, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                        End If
                   End If
                   If InStr(1, Days(intCount), "ג", vbTextCompare) > 0 Or InStr(1, Days(intCount), "שלישי", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 5, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ד", vbTextCompare) > 0 Or InStr(1, Days(intCount), "רביעי", vbTextCompare) > 0 Then
                            Dates(intI) = EndDate
                            Dates(intI + 1) = EndDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                 intCount = intCount + 1
                 Loop
            If Not Simulate Then
                POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                        Dates, Trim(Str(intMovies)), MarTimes
            End If
            intNumMov = intNumMov + 1
            .StatBar.Panels(4).Text = ""
            .StatBar.Panels(1).Text = "Finished uploading data."
            DoEvents
            .fraSource.Tag = ""
            Do
                .txtAddress.Tag = strText 'Back-up
                Do While Not EOF(FreeFileNum) And Trim(strText) = ""
                     MoviePos = Seek(FreeFileNum)
                     Line Input #FreeFileNum, strText
                     .fraSource.Tag = "Been here"
                Loop
                'Check to see whether it's a cinema
                If Not EOF(FreeFileNum) Then Line Input #FreeFileNum, strText
                If HaveHour(strText) Then
                    'That was a movie
                    .cbxType.Tag = "-1"
                End If
                If .fraSource.Tag = "Been here" Then
                    Seek #FreeFileNum, MoviePos
                    Line Input #FreeFileNum, strText
                Else
                    strText = .txtAddress.Tag
                End If
            Loop Until Trim(strText) <> "" Or EOF(FreeFileNum) Or .cbxType.Tag = "-1"
            If EOF(FreeFileNum) Or .cbxType.Tag <> "-1" Then Exit Do
            Loop
        intNumCine = intNumCine + 1
        Loop
        End If 'Inner "End if", for checking the dates
      ElseIf .cbxType.ListIndex = Dizengof Then 'Dizengof
        MovStartDate = Format(StartDate, "d/m")
        MovEndDate = Format(EndDate, "d/m")
        .StatBar.Panels(1).Text = "Searching for cinema..."
        .StatBar.Panels(2).Text = ""
        DoEvents
        Do
            Line Input #FreeFileNum, strText
        Loop Until Trim(strText) <> ""
        If DontDoMore Then
            Exit Do
        End If
        strText = Trim(strText)
        intCinemas = GetCinema(Cinemas, strText)
        If intCinemas = -1 Then Exit Do 'User chose to skip in frmManual
        If intCinemas = -2 Then Err.Raise 5250, , "User-generated emergency stop"
        .StatBar.Panels(2).Text = Cinemas(intCinemas)
        If SkipCinemas And intCinemas <> .txtStartAt.Text Then
            Exit Do
        End If
        SkipCinemas = False
        If .chkPositionEnd.Value And intCinemas = .txtEndAt.Text Then
            DontDoMore = True
        End If
        Do
            Line Input #FreeFileNum, strText
        Loop Until IsDate(Trim(Right(strText, 8))) Or IsDate(Trim(Right(strText, 7))) 'It can be shorted to 6 digits, such as "6/9/04 "
        strText = Trim(Right(strText, 8))
        If Not IsDate(strText) Then strText = Trim(Right(strText, 7))
        intEnd = Weekday(strText)
        Select Case intEnd
            Case vbSunday To vbThursday
                intEnd = vbThursday - intEnd
            Case vbFriday
                intEnd = 6
            Case vbSaturday
                intEnd = 5
        End Select
        .cmdWorkShop.Tag = DateAdd("d", intEnd, strText)
        Do While StartDate = .cmdWorkShop.Tag
                 .StatBar.Panels(1).Text = "Found matching dates; Searching for movie..."
                 DoEvents
                 If EOF(FreeFileNum) Then Exit Do
                 Do
                    Line Input #FreeFileNum, strText
                 Loop Until Trim(strText) <> "" Or EOF(FreeFileNum)
                 If EOF(FreeFileNum) Then Exit Do
            Do
                 .StatBar.Panels(1).Text = "Determining movie..."
                 .StatBar.Panels(3).Text = ""
                 DoEvents
                 strText = NoQuotesString(strText)
                 intMovies = GetMovie(Movies, strText)
                 If intMovies = -1 Then Exit Do 'User chose to skip in frmManual
                 If intMovies = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(3).Text = Movies(intMovies)
                 If DontFilmMore Then Exit Do
                 If SkipMovies Then
                    If intMovies <> frmWorkShop.cbxMovies(0).ItemData(frmWorkShop.cbxMovies(0).ListIndex) Then Exit Do
                 End If
                 SkipMovies = False
                 If frmWorkShop.chkEndMovie.Value = 1 Then
                    If intMovies = frmWorkShop.cbxMovies(1).ItemData(frmWorkShop.cbxMovies(1).ListIndex) Then DontDoMore = True
                 End If
                 .StatBar.Panels(1).Text = "Searching for days & times..."
                 DoEvents
                 Do
                    Line Input #FreeFileNum, strText
                 Loop Until Trim(strText) <> ""
                 .StatBar.Panels(1).Text = "Extracting days & times for movie #" & intMovies & "..."
                 DoEvents
                 ResetArray Times
                 ResetArray Days
                 ResetArray MarTimes
                 ResetVarArray Dates
                 intCount = 1
                 Do While Trim(strText) <> "" And intCount <= MaxTimes And HaveHour(strText)
                    strText = NoDoubleSpaceString(CommaString(strText))
                    .cbxType.Tag = WhereHour(strText)
                    Days(intCount) = Trim(DashString(Left(strText, Val(.cbxType.Tag) - 1)))
                    If Trim(Right(Days(intCount), 2)) = "-" Then Days(intCount) = CropString(Days(intCount), "-")
                    Days(intCount) = CropString(CutEnglish(Days(intCount)), "/")
                    Days(intCount) = NoDoubleSpaceString(CropString(Days(intCount), "יום"))
                    If Left(Days(intCount), 1) = "-" Then Days(intCount) = Trim(Right(Days(intCount), Len(Days(intCount)) - 1))
                    Times(intCount) = Trim(Mid(strText, Val(.cbxType.Tag)))
                    Times(intCount) = CropString(Times(intCount), ".")
                    Line Input #FreeFileNum, strText
                    intCount = intCount + 1
                 Loop
                 .StatBar.Panels(1).Text = "Filling up the buffer..."
                 DoEvents
                 intCount = 1
                 intI = 1
                 intTimes = 1
                 Do While intCount <= MaxTimes And Days(intCount) <> ""
                   If InStr(1, Days(intCount), "-", vbTextCompare) > 0 Or InStr(1, Days(intCount), " - ", vbTextCompare) > 0 Or InStr(1, Days(intCount), " – ", vbTextCompare) > 0 Then
                        If InStr(1, Days(intCount), " - ", vbTextCompare) = 2 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) - 1, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), " - ", vbTextCompare) + 3, 1)), StartDate)
                        ElseIf InStr(1, Days(intCount), " - ", vbTextCompare) >= 3 Then
                            intEnd = FindLast(Days(intCount), " - ") - 1
                            intStart = FindLast(Days(intCount), " ", intEnd) + 1
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), intStart, Len(Days(intCount)) - intEnd)), StartDate)
                            intStart = FindLast(Days(intCount), " - ") + 4
                            intEnd = InStr(intStart, Days(intCount), " ", vbTextCompare)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), intStart, Len(Days(intCount)) - intEnd)), StartDate)
                        ElseIf InStr(1, Days(intCount), "-", vbTextCompare) = 2 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 1, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)
                        ElseIf InStr(1, Days(intCount), "-", vbTextCompare) >= 3 Then
                            Dates(intI) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) - 2, 1)), StartDate)
                            Dates(intI + 1) = DateAdd("d", AnalyzeDay(Mid(Days(intCount), InStr(1, Days(intCount), "-", vbTextCompare) + 1, 1)), StartDate)
                        End If
                        If InStr(1, Days(intCount), "-", vbTextCompare) < 2 Or DateDiff("d", StartDate, Dates(intI)) = -1 Or DateDiff("d", StartDate, Dates(intI + 1)) = -1 Then
                            frmDates.txtMissing.Text = Days(intCount)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Show
                            frmDates.Show 1
                            Dates(intI) = FormatDateTime(frmDates.DTPicker(1).Value, vbShortDate)
                            Dates(intI + 1) = FormatDateTime(frmDates.DTPicker(2).Value, vbShortDate)
                            If frmGlobus.cmdMinimize.Tag = "1" Then frmGlobus.Hide
                        End If
                        intI = intI + 2
                        FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                        MarTimes(intTimes) = Times(intCount)
                        intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ה", vbTextCompare) > 0 Or InStr(1, Days(intCount), "חמישי", vbTextCompare) > 0 Then
                            Dates(intI) = StartDate
                            Dates(intI + 1) = StartDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "שישי", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 1, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "שבת", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 2, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "א", vbTextCompare) > 0 Or InStr(1, Days(intCount), "ראשון", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 3, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   'If (InStr(1, Days(intCount), "ב", vbTextCompare) > 0 And InStr(1, Days(intCount), "שבת", vbTextCompare) = 0 And InStr(1, Days(intCount), "רביעי", vbTextCompare) = 0) Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                   If InStr(1, Days(intCount), "ב", vbTextCompare) > 0 Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then 'And InStr(1, Days(intCount), "שבת", vbTextCompare) = 0) Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                        If (Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ש" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ת" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "י" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ר") Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 4, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                        End If
                   End If
                   If InStr(1, Days(intCount), "ג", vbTextCompare) > 0 Or InStr(1, Days(intCount), "שלישי", vbTextCompare) > 0 Then
                            Dates(intI) = DateAdd("d", 5, StartDate)
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ד", vbTextCompare) > 0 Or InStr(1, Days(intCount), "רביעי", vbTextCompare) > 0 Then
                            Dates(intI) = EndDate
                            Dates(intI + 1) = EndDate
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                 intCount = intCount + 1
                 Loop
            If Not Simulate Then
                POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                        Dates, Trim(Str(intMovies)), MarTimes
            End If
            intNumMov = intNumMov + 1
            .StatBar.Panels(4).Text = ""
            .StatBar.Panels(1).Text = "Finished uploading data."
            DoEvents
            If Not EOF(FreeFileNum) Then Line Input #FreeFileNum, strText
            Do While Not EOF(FreeFileNum) And Trim(strText) = ""
                Line Input #FreeFileNum, strText
            Loop
            If InStr(1, strText, .cmdWorkShop.Tag, vbTextCompare) > 0 Or EOF(FreeFileNum) Then Exit Do
            Loop
        intNumCine = intNumCine + 1
        Loop
      ElseIf .cbxType.ListIndex = LevExcel Then 'Lev (Excel)
        MovStartDate = Format(StartDate, "dd/mm/yyyy")
        MovEndDate = Format(EndDate, "dd/mm/yyyy")
        Line Input #FreeFileNum, strText
        'שני is the name of the distributer
        If InStr(1, strText, MovStartDate & vbTab & "שני", vbTextCompare) > 0 Or InStr(1, strText, "שני" & vbTab & MovStartDate, vbTextCompare) > 0 Then 'Or intCinemas <> 0 'Also check for not the first time being here
            Do While InStr(1, strText, "עיר" & vbTab & "קולנוע" & vbTab & "סרט" & vbTab & "שעות", vbTextCompare) = 0 'Or Not intCinemas <> 0 'Also check (again) for not the first time being here
                Line Input #FreeFileNum, strText
            Loop
            Line Input #FreeFileNum, strText
            Do While Trim(strText) = ""
                Line Input #FreeFileNum, strText
            Loop
            Do
                 MoviePos = 1
                 .StatBar.Panels(1).Text = "Found matching dates; Determining cinema..."
                 .StatBar.Panels(2).Text = ""
                 DoEvents
                 If DontDoMore Then
                    Exit Do
                 End If
                 MoviePos = InStr(MoviePos, strText, vbTab, vbTextCompare) + 1
                 .StatBar.Panels(2).Tag = Trim(Mid(strText, MoviePos, InStr(MoviePos, strText, vbTab, vbTextCompare) - MoviePos))
                 intCinemas = GetCinema(Cinemas, NoQuotesString(.StatBar.Panels(2).Tag))
                 If intCinemas = -1 Then Exit Do 'User chose to skip in frmManual
                 If intCinemas = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(2).Text = Cinemas(intCinemas)
                 If SkipCinemas And intCinemas <> .txtStartAt.Text Then
                    Exit Do
                 End If
                 SkipCinemas = False
                 If .chkPositionEnd.Value And intCinemas = .txtEndAt.Text Then
                    DontDoMore = True
                 End If
            Do
                 .StatBar.Panels(1).Text = "Determining movie..."
                 .StatBar.Panels(3).Text = ""
                 DoEvents
                 MoviePos = InStr(MoviePos, strText, vbTab, vbTextCompare) + 1
                 .StatBar.Panels(3).Tag = Trim(Replace(Trim(Mid(strText, MoviePos, InStr(MoviePos, strText, vbTab, vbTextCompare) - MoviePos)) & " " & Trim(Replace(Right(strText, Len(strText) - FindLast(strText, vbTab)), "כן", "")), vbTab, "")) 'We should also remove the "כן" that we discover by mistake, they don't belong to the name of the film
                 intMovies = GetMovie(Movies, NoQuotesString(.StatBar.Panels(3).Tag))
                 If intMovies = -1 Then Exit Do 'User chose to skip in frmManual
                 If intMovies = -2 Then Err.Raise 5250, , "User-generated emergency stop"
                 .StatBar.Panels(3).Text = Movies(intMovies)
                 If DontFilmMore Then Exit Do
                 If SkipMovies Then
                    If intMovies <> frmWorkShop.cbxMovies(0).ItemData(frmWorkShop.cbxMovies(0).ListIndex) Then Exit Do
                 End If
                 SkipMovies = False
                 If frmWorkShop.chkEndMovie.Value = 1 Then
                    If intMovies = frmWorkShop.cbxMovies(1).ItemData(frmWorkShop.cbxMovies(1).ListIndex) Then DontDoMore = True
                 End If
                 .StatBar.Panels(1).Text = "Extracting days & times for movie #" & intMovies & "..."
                 DoEvents
                 ResetArray Times
                 ResetArray Days
                 ResetArray MarTimes
                 ResetVarArray Dates
                 intCount = 1
                 MoviePos = InStr(MoviePos, strText, vbTab, vbTextCompare) + 1
                 Do While Trim(strText) <> "" And HaveHour(Mid(strText, MoviePos)) And intCount <= MaxTimes
                    'Trimming down strText
                    strText = DashString(NoDoubleSpaceString(CommaString(NoQuotesString(Mid(strText, MoviePos)))))
                    'Get the times
                    Times(intCount) = Mid(strText, 1, InStr(1, strText, vbTab, vbTextCompare) - 1)
                    'Count the tabs and decide what days they represent
                    MoviePos = InStr(1, strText, vbTab, vbTextCompare) + 1
                    InsertDaysByString Mid(strText, MoviePos), vbTab, Days(intCount)
                    If EOF(FreeFileNum) Then
                        strText = ""
                    Else
                        Line Input #FreeFileNum, strText
                        MoviePos = InStr(1, strText, vbTab, vbTextCompare) + 1 'Getting MoviePos to point at the cinema name
                        If StrComp(.StatBar.Panels(2).Tag, Mid(strText, MoviePos, InStr(MoviePos, strText, vbTab, vbTextCompare) - MoviePos)) <> 0 Then Exit Do 'Not the same cinema
                        MoviePos = InStr(MoviePos, strText, vbTab, vbTextCompare) + 1 'Getting MoviePos to point at the movie name
                        If StrComp(.StatBar.Panels(3).Tag, Trim(Replace(Trim(Mid(strText, MoviePos, InStr(MoviePos, strText, vbTab, vbTextCompare) - MoviePos)) & " " & Trim(Replace(Right(strText, Len(strText) - FindLast(strText, vbTab)), "כן", "")), vbTab, ""))) <> 0 Then Exit Do 'Not the same movie
                        MoviePos = InStr(MoviePos, strText, vbTab, vbTextCompare) + 1 'Getting MoviePos to point at the movie times
                    End If
                    intCount = intCount + 1
                 Loop
                 .StatBar.Panels(1).Text = "Filling up the buffer..."
                 DoEvents
                 intCount = 1
                 intI = 1
                 intTimes = 1
                 Do While intCount <= MaxTimes And Days(intCount) <> ""
                   'No ranges! Whee-pee!
                   If InStr(1, Days(intCount), "ה", vbTextCompare) > 0 Or InStr(1, Days(intCount), "חמישי", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbThursday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbThursday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbThursday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "שישי", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbFriday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbFriday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbFriday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "שבת", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbSaturday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbSaturday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbSaturday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "א", vbTextCompare) > 0 Or InStr(1, Days(intCount), "ראשון", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbSunday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbSunday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbSunday, StartDate)
                            End If
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ב", vbTextCompare) > 0 Then 'And InStr(1, Days(intCount), "שבת", vbTextCompare) = 0) Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                        If (Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ש" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ת" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "י" And Mid(Days(intCount), InStr(1, Days(intCount), "ב", vbTextCompare) + 1, 1) <> "ר") Or InStr(1, Days(intCount), "שני", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbMonday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbMonday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbMonday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                        End If
                   End If
                   If InStr(1, Days(intCount), "ג", vbTextCompare) > 0 Or InStr(1, Days(intCount), "שלישי", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbTuesday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbTuesday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbTuesday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                   If InStr(1, Days(intCount), "ד", vbTextCompare) > 0 Or InStr(1, Days(intCount), "רביעי", vbTextCompare) > 0 Then
                            If InDates(Dates(), dhNextDOW(vbWednesday, StartDate)) Then 'for chronica lists that are more than 6 days
                                Dates(intI) = dhNextDOW(vbWednesday, StartDate + (EndDate - StartDate - 1))
                            Else
                                Dates(intI) = dhNextDOW(vbWednesday, StartDate)
                            End If
                            Dates(intI + 1) = Dates(intI)
                            intI = intI + 2
                            FixingUps intCinemas, intMovies, Dates, Times, intCount, intI
                            MarTimes(intTimes) = Times(intCount)
                            intTimes = intTimes + 1
                   End If
                 intCount = intCount + 1
                 Loop
            If Not Simulate Then
                POSTing strPrefix & .txtAddress, Trim(Str(intCinemas)), _
                        Dates, Trim(Str(intMovies)), MarTimes
            End If
            intNumMov = intNumMov + 1
            .StatBar.Panels(4).Text = ""
            .StatBar.Panels(1).Text = "Finished uploading data."
            DoEvents
            Do While Not EOF(FreeFileNum) And Trim(strText) = ""
                Line Input #FreeFileNum, strText
            Loop
            If EOF(FreeFileNum) Then Exit Do 'Check for end-of-file
            MoviePos = InStr(1, strText, vbTab, vbTextCompare) + 1 'Reset the variable
            If StrComp(.StatBar.Panels(2).Tag, Mid(strText, MoviePos, InStr(MoviePos, strText, vbTab, vbTextCompare) - MoviePos), vbTextCompare) <> 0 Or EOF(FreeFileNum) Then Exit Do 'Check for different cinema
            Loop 'of movies
        intNumCine = intNumCine + 1
        Loop Until EOF(FreeFileNum) 'of cinemas
        End If 'of date-checking
        End If
    Loop
    Reset 'Closes reading from file
    If .cmdMinimize.Tag = "1" Then
        .Show
        .cmdMinimize.Tag = ""
        RemoveNotifyIcon
    End If
    .StatBar.Panels(1).Text = ""
    .StatBar.Panels(4).Text = ""
    .StatBar.Panels(2).Text = ""
    .StatBar.Panels(3).Text = ""
    
    'Interface changes
    FormCosmetics False
    'fraStatus.Enabled = False
    'cmdExit.Enabled = True
    'cmdGo.Enabled = True
    MsgBox intNumMov & " movie(s) were entered in total of " & intNumCine & " cinema(s).", vbInformation + vbSystemModal, "Summary"
    If Simulate Then
        .Caption = .cmdSimulate.Tag
        .cmdSimulate.Tag = ""
    End If
    LetsDoIt = True
    Exit Function
Oops: 'Don't touch the "frmGlobus" references, they don't work with the "with" that was set high above
    ProcessErrorHandler Err.Number, FreeFileNum
    frmGlobus.Inet.Cancel
    Reset
    If frmGlobus.cmdMinimize.Tag = "1" Then
        frmGlobus.Show
        RemoveNotifyIcon
        frmGlobus.cmdMinimize.Tag = ""
    End If
    frmGlobus.StatBar.Panels(1).Text = ""
    frmGlobus.StatBar.Panels(4).Text = ""
    frmGlobus.StatBar.Panels(2).Text = ""
    frmGlobus.StatBar.Panels(3).Text = ""
    DoEvents
    
    If Not frmGlobus.cmdGo.Enabled Then 'if there should be interface changes, and
                              'cmdStop's procedure hasn't already done them
        'Interface changes
        FormCosmetics False
        'Special treat
        'cmdExit.Enabled = True
        'cmdGo.Enabled = True
        If Simulate Then
            frmGlobus.Caption = frmGlobus.cmdSimulate.Tag
            frmGlobus.cmdSimulate.Tag = ""
        End If
    End If
    
    End With 'End of frmGlobus references
End Function

Public Sub DetermineSourceType(FreeFileNum%)

    Dim strText$

    With frmGlobus
        Do While Not EOF(FreeFileNum)
            Line Input #FreeFileNum, strText
            If InStr(1, strText, "גלובוס גרופ", vbTextCompare) > 0 Then 'Globus Group
                .cbxType.ListIndex = GG
                Seek #FreeFileNum, 1
                Exit Do
            ElseIf InStr(1, strText, "פינת קינג ג'ורג'", vbTextCompare) > 0 Then 'Lev
                .cbxType.ListIndex = OldLev
                Seek #FreeFileNum, 1
                Exit Do
            ElseIf InStr(1, ReverseText(RemHenChar(HTMLessString(HTMLessString(strText)))), "תיאטראות ישראל", vbTextCompare) > 0 Then 'Rav-Hen
                .cbxType.ListIndex = RavHen
                Seek #FreeFileNum, 1
                Exit Do
            ElseIf InStr(1, strText, "בלונדון מיניסטור", vbTextCompare) > 0 Then 'London
                .cbxType.ListIndex = London
                Seek #FreeFileNum, 1
                Exit Do
            ElseIf InStr(1, strText, "סינמה קפה ""מוריה"" חיפה", vbTextCompare) > 0 Or InStr(1, strText, "פנורמה 1-2-3", vbTextCompare) > 0 Or InStr(1, strText, "סינמה קפה ""עממי""", vbTextCompare) > 0 Or InStr(1, strText, """בית גבריאל""", vbTextCompare) > 0 Or InStr(1, strText, "בתי קולנוע חיפה ""פנורמה""", vbTextCompare) > 0 Then 'Northen Cinemas
                .cbxType.ListIndex = Haifa
                Seek #FreeFileNum, 1
                Exit Do
            ElseIf InStr(1, strText, "תיאטרון ירושלים לאמנויות הבמה", vbTextCompare) > 0 Then 'Jerusalem Theater
                .cbxType.ListIndex = Jerusalem
                Seek #FreeFileNum, 1
                Exit Do
            ElseIf InStr(1, strText, "סרטי השבוע, החל מהיום, יום חמישי ", vbTextCompare) > 0 Then 'Specially-compiled
                .cbxType.ListIndex = Compiled
                Seek #FreeFileNum, 1
                Exit Do
            ElseIf InStr(1, strText, "מודעות לעיתונות", vbTextCompare) > 0 Then 'Dizengof Tel-Aviv
                .cbxType.ListIndex = Dizengof
                Seek #FreeFileNum, 1
                Exit Do
            ElseIf InStr(1, strText, "עיר" & vbTab & "קולנוע" & vbTab & "סרט" & vbTab & "שעות", vbTextCompare) > 0 Then 'Lev (Excel)
                .cbxType.ListIndex = LevExcel
                Seek #FreeFileNum, 1
                Exit Do
            End If
        Loop
    End With
End Sub

Public Sub POSTing(ByVal Dest$, ByVal Cine$, arDates(), ByVal Mov$, arTimes$())
'NOT taken from http://www.tagconsulting.com/Show.asp?Id=1025&S=1
    On Error GoTo Oops
    Dim strData, intI%, intMax%, intMaxInc%, intCount%, intI2%, intDummy%, intTime%
    
    frmGlobus.StatBar.Panels(1).Text = "Uploading data..."
    DoEvents
    intI = 1
    intI2 = 3
    intDummy = 3
    intMaxInc = 8
    intTime = 1
    If FindBiggest(arDates()) > intMaxInc Then
        intMax = intMaxInc
    Else
        intMax = FindBiggest(arDates())
    End If
    Do
        strData = Dest & "?cinema_id=" & Cine & "&from_date=" _
                  & arDates(intI) & "&to_date=" & arDates(intI + 1) & _
                  "&times[0]=" & arTimes(intTime)
        intI2 = 2
        intTime = intTime + 1
        For intI = intDummy To intMax Step 2
            strData = strData & "&from_date" & Mid(Str(intI2), 2) & "=" & _
                    arDates(intI)
            strData = strData & "&to_date" & Mid(Str(intI2), 2) & "=" & _
                    arDates(intI + 1)
            strData = strData & "&times" & Mid(Str(intI2), 2) & "[0]=" & _
                    Trim(arTimes(intTime))
            intI2 = intI2 + 1
            intTime = intTime + 1
        Next intI
        intDummy = intI + 2
        strData = strData & "&hall[0]=1&movie_id[0]=" & Mov & "&Action=save"
        'required because of a bug
        frmGlobus.Tag = strData
        strData = frmGlobus.Inet.OpenURL(frmGlobus.Tag)
        If InStr(1, strData, "עריכת לוח הזמנים", vbTextCompare) = 0 Then
            Err.Raise 5248, , "Could not reach Chronica, or it didn't gave back an expected response."
        ElseIf InStr(1, strData, "שעה שגויה:", vbTextCompare) > 0 Then
            Err.Raise 5249, , "Chronica returned an ""Invalid hour"" message."
        End If
        frmGlobus.StatBar.Panels(4).Text = ""
        intMaxInc = intMaxInc * 2
        If FindBiggest(arDates()) > intMaxInc Then
            intMax = intMaxInc
        Else
            intMax = FindBiggest(arDates())
        End If
    Loop Until intI > intMax
    Exit Sub
Oops:
    If Err.Number = 5248 Then
        MsgBox "Could not reach Chronica, or it didn't gave back an expected response." _
               & " Please check whether the target address and the password is correct." _
               & vbCrLf & "Please stop ChroniKey from further uploadings and consult t" _
               & "he main form." & vbCrLf & "Current cinema #: " & Cine & vbCrLf & "Cu" _
               & "rrent movie #: " & Mov, vbCritical + vbSystemModal, "Error while try" _
               & "ing to upload data"
    ElseIf Err.Number = 5249 Then
        MsgBox "Chronica returned an ""Invalid hour"" message, which is unexpected. Pl" _
               & "ease stop ChroniKey from further uploadings and scan the source file" _
               & " for mistakes or errors." & vbCrLf & "Current cinema #: " & Cine & _
               vbCrLf & "Current movie #: " & Mov, vbCritical + vbSystemModal, "Error " _
               & "while trying to upload data"
    Else
        MsgBox "ChroniKey encountered an error while trying to upload data. Please sto" _
               & "p ChroniKey from further uploadings." & vbCrLf & "Current cinema #: " _
               & Cine & vbCrLf & "Current movie #: " & Mov, vbCritical + vbSystemModal, _
               "Unexpected Error"
    End If
    ErrorLog Error, frmGlobus.StatBar.Panels(1).Text & " (" & Mid(frmGlobus.Tag, InStr(1, frmGlobus.Tag, "?", vbTextCompare)), CInt(Cine), CInt(Mov), 0
End Sub

Private Sub InsertDaysByString(ByVal strSource As String, ByVal strDelimit As String, ByRef strDest As String)
    Dim DaysRange() As String
    
    DaysRange = Split(strSource, strDelimit)
    
    If Not Trim(DaysRange(0)) = "" Then strDest = strDest & "חמישי "
    If Not Trim(DaysRange(1)) = "" Then strDest = strDest & "שישי "
    If Not Trim(DaysRange(2)) = "" Then strDest = strDest & "שבת "
    If Not Trim(DaysRange(3)) = "" Then strDest = strDest & "ראשון "
    If Not Trim(DaysRange(4)) = "" Then strDest = strDest & "שני (ב) " 'ב is required
    If Not Trim(DaysRange(5)) = "" Then strDest = strDest & "שלישי "
    If Not Trim(DaysRange(6)) = "" Then strDest = strDest & "רביעי "
End Sub

Public Sub GetUpdate()
    Dim SiteUpdate$, FileName$, Version$, TimeOut%
    
    On Error Resume Next
    Version = ""
    With frmGlobus
        .StatBarIdle.Panels(1).Text = "Checking for a newer version of " & App.Title & "..."
        DoEvents
        SiteUpdate = "http://" & frmGlobus.txtUserName.Text & ":" & frmGlobus.txtPassword.Text & "@" & ChroniKeySite & "/Update.ini"
        FileName = "http://" & frmGlobus.txtUserName.Text & ":" & frmGlobus.txtPassword.Text & "@" & ChroniKeySite & "/ChroniKey.zip"
        TimeOut = .Inet.RequestTimeout
        .Inet.RequestTimeout = 10 'Set a shorter time-out, so if the server fails, the user won't wait such a long time
        .Tag = SiteUpdate
        Version = .Inet.OpenURL(.Tag)
        '.txtINetStatus.Text = ""
        .Inet.RequestTimeout = TimeOut 'Restore
        If InStr(1, Version, "401", vbTextCompare) > 0 Then Exit Sub 'Wrong password
        If Version <= App.Major & "." & App.Minor & "." & App.Revision Or Trim(Version) = "" Then
            .StatBarIdle.Panels(1).Text = "No newer version was released."
            DoEvents
            Exit Sub
        Else
            If MsgBox("A newer version of " & App.Title & " was released. Download update?" & vbCrLf & "(If you choose ""yes"", " & App.Title & " will terminate so you can install the new update.)", vbYesNo + vbSystemModal + vbQuestion, "Update from version " & App.Major & "." & App.Minor & "." & App.Revision & " to version " & Version) = vbYes Then
                ShellExecute 0&, vbNullString, FileName, vbNullString, vbNullString, vbNormalFocus
                End
            End If
        End If
    End With
End Sub

Public Sub ProcessErrorHandler(Mis As Long, fileNum As Integer)
Dim strPrefix$

If BeingStopped Then Exit Sub 'Don't display or log error messages

With frmGlobus
Select Case Mis
    Case 76, 53 'If not such file or directory
        strPrefix = "No such file or directory. Couldn't " & _
               "find the file specified in the Source text box. " & _
               vbCrLf & "Please check whether the file exists an" & _
               "d then try again." & vbCrLf & vbCrLf & "Error #" & _
               Err.Number & ": " & Err.Description & vbCrLf & "L" & _
               "ast status message: " & .StatBar.Panels(1).Text
    Case 62
        strPrefix = "Can't retrieve information from the file specifi" & _
               "ed. The file is either empty or in a non-readabl" & _
               "e format." & vbCrLf & vbCrLf & "Error #" & _
               Err.Number & ": " & Err.Description & vbCrLf & "L" & _
               "ast status message: " & .StatBar.Panels(1).Text & vbCrLf & _
               "Current cinema: " & intCinemas & vbCrLf & "Curre" & _
               "nt movie: #" & intMovies
    Case 5252
        strPrefix = "Can't detect the source type. Please select one " & _
               "of the source types."
    Case 5111
        strPrefix = "No cinemas or no movies have been found while tried " & _
                    "to obtain them from the Internet. Please check your " & _
                    "password for typos and then try again."
    Case Else
    strPrefix = "The program encountered an error while trying to get " & _
            "information off the Internet. Please check your pass" & _
            "word for typos and if it's correct contact " & _
            AuthEMail & " for further instructions, and enclose " & _
            "the following details:" & vbCrLf & vbCrLf & "Error #" & _
            Err.Number & ": " & Err.Description & vbCrLf & "Last " & _
            "status message: " & .StatBar.Panels(1).Text & vbCrLf & _
            "Current cinema: #" & intCinemas & vbCrLf & "Current " & _
            "movie: #" & intMovies
End Select
End With
ErrorLog Err, frmGlobus.StatBar.Panels(1).Text, CInt(intCinemas), intMovies, fileNum
MsgBox strPrefix, vbSystemModal + vbCritical, "Error"
End Sub

Public Sub FixingUps(ByVal int_Cinemas%, ByVal int_Movies%, Da_tes(), Ti_mes$(), ByVal int_Count%, ByVal int_I%)
    Dim int__Start%
    
    int__Start = FixAvail(int_Cinemas, int_Movies, Da_tes(int_I - 1), Da_tes(int_I - 2))
    If int__Start >= 0 Then
        Ti_mes(int_Count) = NoDoubleSpaceString(Ti_mes(int_Count) & " " & CommaString(frmWorkShop.txtAdd(int__Start - 1).Text))
        RemoveHour Ti_mes(int_Count), NoDoubleSpaceString(CommaString(frmWorkShop.txtDelete(int__Start - 1).Text))
    End If
End Sub

Private Sub ErrorLog(Error As ErrObject, LastStatus As String, CurCine As Integer, CurMov As Integer, ByVal fileNum%)
    On Error Resume Next
    Dim LogLine As String
     
    LogLine = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
    LogLine = LogLine & "@ " & Now & ": (" & Error.Number & ") " & Error.Description & vbCrLf
    LogLine = LogLine & "Last status message: " & LastStatus & vbCrLf
    LogLine = LogLine & "Current cinema: #" & CurCine & vbCrLf
    LogLine = LogLine & "Current movie: #" & CurMov & vbCrLf
    LogLine = LogLine & "Progress in file: " & Seek(fileNum) & vbCrLf
    
    App.LogEvent LogLine, vbLogEventTypeError
End Sub
