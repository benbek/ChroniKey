VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGlobus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ChroniKey"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9180
   Icon            =   "fraGlobus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9180
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUpdateMovies 
      Height          =   285
      Left            =   2640
      TabIndex        =   18
      Text            =   ""
      Top             =   3600
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   4140
      Visible         =   0   'False
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   450
      SimpleText      =   "ChroniKey"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5512
            MinWidth        =   3528
            Text            =   "Ready."
            TextSave        =   "Ready."
            Object.ToolTipText     =   "Current status of work"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Current cinema that is being processed"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Current movie that is being proccessed"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Internet connection information"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialogFile 
      Left            =   6600
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   -2147483640
      DefaultExt      =   "*.htm;*.html;*.txt"
      DialogTitle     =   "Source Path"
      Filter          =   "Chronica Files (*.htm, *.html, *.txt)|*.htm;*.html;*.txt"
      Flags           =   4100
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   6600
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   1095
      Left            =   7200
      Picture         =   "fraGlobus.frx":35FA
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go!"
      Default         =   -1  'True
      Height          =   1095
      Left            =   4800
      Picture         =   "fraGlobus.frx":3AAA
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Frame fraGeneral 
      Caption         =   " General options "
      Height          =   1935
      Left            =   4680
      TabIndex        =   19
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdSimulate 
         Caption         =   "Simu&lation..."
         Height          =   375
         Left            =   2400
         Picture         =   "fraGlobus.frx":3EF8
         TabIndex        =   27
         ToolTipText     =   "Simulation options"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdWorkShop 
         Caption         =   "ChroniKey's &Workshop..."
         Height          =   375
         Left            =   120
         Picture         =   "fraGlobus.frx":3F91
         TabIndex        =   26
         ToolTipText     =   "More advanced options"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkPosition 
         Caption         =   "Sta&rt at cinema #:"
         Height          =   255
         Left            =   120
         Picture         =   "fraGlobus.frx":4035
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkSkipCC 
         Caption         =   "Orly wasn't nice this week, so skip Cinema City (#95)"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.TextBox txtStartAt 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   23
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox chkPositionEnd 
         Caption         =   "&End at cinema #:"
         Height          =   255
         Left            =   2280
         Picture         =   "fraGlobus.frx":4404
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtEndAt 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin VB.Frame fraType 
         Caption         =   " &Type of source "
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4215
         Begin VB.ComboBox cbxType 
            Height          =   315
            ItemData        =   "fraGlobus.frx":469B
            Left            =   120
            List            =   "fraGlobus.frx":46BD
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   280
            Width           =   3975
         End
      End
   End
   Begin VB.Frame fraInit 
      Caption         =   " Initialization "
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Frame fraLists 
         Caption         =   " Lists of cinemas and movies "
         Height          =   1335
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   4215
         Begin VB.TextBox txtObtain 
            Height          =   285
            Left            =   2400
            TabIndex        =   14
            Text            =   ""
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtUpdateCinemas 
            Height          =   285
            Left            =   2400
            TabIndex        =   16
            Text            =   ""
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblUpdateMovies 
            Caption         =   "Update the list of &movies by:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblObtain 
            Caption         =   "&Obtain the lists from:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   320
            Width           =   1455
         End
         Begin VB.Label lblUpdateCinemas 
            Caption         =   "Update the list of &cinemas by:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   630
            Width           =   2175
         End
      End
      Begin VB.Frame fraDest 
         Caption         =   " Destination:"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   4215
         Begin VB.TextBox txtAddress 
            Height          =   285
            Left            =   840
            TabIndex        =   11
            Text            =   ""
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label lblAddress 
            Caption         =   "&Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   275
            Width           =   615
         End
      End
      Begin VB.Frame fraSource 
         Caption         =   " &Source: "
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   4215
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&..."
            Height          =   255
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox txtSource 
            Height          =   285
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   7
            ToolTipText     =   "Enter the source files for the times. You can also drag & drop here .htm, .html and .txt files"
            Top             =   280
            Width           =   3615
         End
      End
      Begin VB.Frame fraLogin 
         Caption         =   " Authorization "
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2640
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtUserName 
            Height          =   285
            Left            =   960
            TabIndex        =   3
            Text            =   ""
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblPass 
            Caption         =   "&Password:"
            Height          =   255
            Left            =   1800
            TabIndex        =   4
            Top             =   280
            Width           =   855
         End
         Begin VB.Label lblUser 
            Caption         =   "&Username:"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   280
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "S&top!"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   4800
      Picture         =   "fraGlobus.frx":47AC
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Stops the current process"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "Minimi&ze"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   7200
      Picture         =   "fraGlobus.frx":481D
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Minimize ChroniKey to the notification area (taskbar tray)"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame frmMisgeret 
      Height          =   1455
      Left            =   4680
      TabIndex        =   29
      Top             =   2280
      Width           =   4455
   End
   Begin MSComctlLib.StatusBar StatBarIdle 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   32
      Top             =   3870
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   476
      SimpleText      =   "Chronikey is Ready"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8255
            Text            =   "ChroniKey ready."
            TextSave        =   "ChroniKey ready."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   0
            TextSave        =   "01/01/2006"
            Object.ToolTipText     =   "Today's date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "From 99/99/9999 to 00/00/0000"
            TextSave        =   "From 99/99/9999 to 00/00/0000"
            Object.ToolTipText     =   "Range of the movie week's dates (click to change)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGlobus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub chkPosition_Click()
    txtStartAt.Enabled = chkPosition.Value
    If txtStartAt.Enabled Then txtStartAt.SetFocus
End Sub

Private Sub chkPositionEnd_Click()
    txtEndAt.Enabled = chkPositionEnd.Value
    If txtEndAt.Enabled Then txtEndAt.SetFocus
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo ResetFile
    CommonDialogFile.ShowOpen
    txtSource.Text = CommonDialogFile.FileName
    Exit Sub
ResetFile:
End Sub

Private Sub cmdGo_Click()
    With frmSimulation
        If .Tag = 0 Then
            LetsDoIt True
        ElseIf .Tag = 1 Then
            If LetsDoIt(True) Then LetsDoIt (False)
        Else
            LetsDoIt False
        End If
    End With
    StatBarIdle.Panels(1).Text = App.Title & " " & App.Major & "." & App.Minor & " ready."
End Sub

Private Sub cmdSimulate_Click()
    frmSimulation.Show vbModal, Me
End Sub

Private Sub cmdWorkShop_Click()
    Dim intMoo%, intYaa%, strPrefix$
    Static BeenHere As Boolean
    
HereWeStart:
    If frmWorkShop.cmdUpdate.Tag <> "" Then
        BeenHere = False
        frmWorkShop.cmdUpdate.Tag = ""
    End If
 If Not BeenHere Then
    If Not ValidInput(1) Then Exit Sub
    
    Me.MousePointer = 11 'Waiting...
    StatBarIdle.Panels(1).Text = "Please wait..."
    strPrefix = "http://" + txtUserName.Text + ":" _
                + txtPassword.Text + "@"
    ResetArray Cinemas
    ResetArray Movies
    CinemasList Data, Cinemas, 1, 1, strPrefix
    'txtINetStatus.Text = ""
    MoviesList Data, Movies, 1, 1, strPrefix
    'txtINetStatus.Text = ""
    StatBarIdle.Panels(1).Text = "Populating ChroniKey's Workshop..."
    DoEvents
    With frmWorkShop
        For intYaa = 0 To .fraRepairs.Count + 1 'We started from 2
            .cbxCinemas(intYaa).Clear
            .cbxMovies(intYaa).Clear
        Next intYaa
        For intMoo = LBound(Cinemas) To UBound(Cinemas)
           If Not IsNumeric(Trim(Cinemas(intMoo))) And Trim(Cinemas(intMoo)) <> "" Then
                For intYaa = 0 To .fraRepairs.Count + 1
                    .cbxCinemas(intYaa).AddItem Cinemas(intMoo)
                    .cbxCinemas(intYaa).ItemData(.cbxCinemas(intYaa).NewIndex) = intMoo
                Next intYaa
           End If
        Next intMoo
        For intMoo = LBound(Movies) To UBound(Movies)
           If Not IsNumeric(Trim(Movies(intMoo))) And Trim(Movies(intMoo)) <> "" Then
                For intYaa = 0 To .fraRepairs.Count + 1
                    .cbxMovies(intYaa).AddItem Movies(intMoo)
                    .cbxMovies(intYaa).ItemData(.cbxMovies(intYaa).NewIndex) = intMoo
                Next intYaa
           End If
        Next intMoo
        .DTPicker(0).MinDate = StartDate
        .DTPicker(0).MaxDate = EndDate
        .DTPickerEnd(0).MinDate = StartDate
        .DTPickerEnd(0).MaxDate = EndDate
        BeenHere = True
    End With
  End If
    With frmWorkShop
        If Me.chkPosition.Value And Trim(txtStartAt.Text) <> "" Then
            .chkStartCinema.Value = 1
            FindInCombo txtStartAt.Text, .cbxCinemas(0)
        Else
            .chkStartCinema.Value = 0
        End If
        If Me.chkPositionEnd.Value And Trim(txtEndAt.Text) <> "" Then
            .chkEndCinema.Value = 1
            FindInCombo txtEndAt.Text, .cbxCinemas(1)
        Else
            .chkEndCinema.Value = 0
        End If
        Me.MousePointer = 0
        Me.StatBarIdle.Panels(1).Text = App.Title & " " & App.Major & "." & App.Minor & " ready."
        DoEvents
        .Show vbModal, Me
        If .chkStartCinema.Value And .cbxCinemas(0).ListIndex > -1 Then
            Me.chkPosition.Value = 1
            Me.txtStartAt.Text = .cbxCinemas(0).ItemData(.cbxCinemas(0).ListIndex)
        Else
            Me.chkPosition.Value = 0
            Me.txtStartAt.Text = ""
        End If
        If .chkEndCinema.Value And .cbxCinemas(1).ListIndex > -1 Then
            Me.chkPositionEnd.Value = 1
            Me.txtEndAt.Text = .cbxCinemas(1).ItemData(.cbxCinemas(1).ListIndex)
        Else
            Me.chkPositionEnd.Value = 0
            Me.txtEndAt.Text = ""
        End If
        If Not .chkStartMovie.Enabled Then 'If one is not enabled so is the other
            .chkStartMovie.Value = 0
            .chkEndMovie.Value = 0
        End If
        If .cmdUpdate.Tag <> "" Then GoTo HereWeStart
    End With
    On Error Resume Next 'So there won't be any problems when launching the window while ChroniKey is working
    cmdGo.SetFocus
End Sub

Private Sub Inet_StateChanged(ByVal State As Integer)
    StatBar.Panels(4).Text = GetState(State)
End Sub

Private Sub cmdExit_Click()
    Form_QueryUnload 0, 0
    End
End Sub

Private Sub cmdStop_Click()
    On Error Resume Next
    If MsgBox("Are you sure you want to stop?", vbApplicationModal + _
    vbYesNo + vbQuestion, "Stopping the process") = vbYes Then
        MsgBox "Please ignore any error messages that may come as Chro" & _
               "niKey is shutting down." & vbCrLf & vbCrLf & "For your" & _
               " information, the current cinema being uploaded is """ & _
               Cinemas(intCinemas) & """," & vbCrLf & "and the current" & _
               " movie is """ & Movies(intMovies) & """.", _
               vbApplicationModal + vbInformation, "Stopping the process"
        BeingStopped = True
        Inet.Cancel
        Reset
    If cmdMinimize.Tag = "1" Then 'Should not happen, because the user has to reach the Stop button by pressing the button
        Me.Show
        cmdMinimize.Tag = ""
        RemoveNotifyIcon
    End If
    StatBar.Panels(1).Text = ""
    StatBar.Panels(2).Text = ""
    StatBar.Panels(3).Text = ""
    StatBar.Panels(4).Text = ""
    StatBarIdle.Panels(1).Text = App.Title & " " & App.Major & "." & App.Minor & " ready."
    DoEvents
    'Interface changes
    FormCosmetics False
    'Special treat
    'fraStatus.Enabled = False
    'cmdExit.Enabled = True
    'cmdGo.Enabled = True
    If cmdSimulate.Tag <> "" Then 'for restoring the previous form title
        Me.Caption = cmdSimulate.Tag
        cmdSimulate.Tag = ""
    End If
    BeingStopped = False
    MsgBox intNumMov & " movie(s) were entered in total of " & intNumCine & " cinema(s).", vbInformation + vbSystemModal, "Summary"
    End If
End Sub

Private Sub Form_Load()
    SetMovieWeek StartDate, EndDate
    
    StatBarIdle.Panels(1).Text = App.Title & " " & App.Major & "." & App.Minor & " ready."
    StatBarIdle.Panels(3).Text = "From " & StartDate & " to " & EndDate
    
    cbxType.ListIndex = 0 'Choose detection
    
    'Birthdays and Fortunate Events!
    If DateDiff("d", Format(StartDate, "d/m"), "8/11") < 7 And DateDiff("d", Format(StartDate, "d/m"), "8/11") >= 0 Then
        Me.BackColor = &HFF&
        Me.Caption = "ChroniKey :: Happy Birthday to RedFish!"
        Me.cbxType.List(2) = "Golan-Globus"
    ElseIf DateDiff("d", Format(StartDate, "d/m"), "16/1") < 7 And DateDiff("d", Format(StartDate, "d/m"), "16/1") >= 0 Then
        Me.BackColor = &H80FF&
        Me.Caption = "ChroniKey :: Happy Birthday to Long John!"
        cmdWorkShop.Caption = "ChroniKey's S&weatshop..."
    ElseIf DateDiff("d", Format(StartDate, "d/m"), "28/2") < 7 And DateDiff("d", Format(StartDate, "d/m"), "28/2") >= 0 Then
        Me.Caption = "I had to chase Christian Slater down the street"
    ElseIf DateDiff("d", Format(StartDate, "d/m"), "1/1") < 7 And DateDiff("d", Format(StartDate, "d/m"), "1/1") >= 0 Then
        Me.Caption = "Wanna kiss me at midnight? I'm Batman."
        StatBarIdle.Panels(1).Text = StatBarIdle.Panels(1).Text & " Happy New Year!"
    ElseIf DateDiff("d", Format(StartDate, "d/m"), "13/10") < 7 And DateDiff("d", Format(StartDate, "d/m"), "13/10") >= 0 Then
        Me.BackColor = &HFF8080
        Me.Caption = "It's a fishy ChroniKey. Mazal-Tov!"
    ElseIf DateDiff("d", Format(StartDate, "d/m"), "21/2") < 7 And DateDiff("d", Format(StartDate, "d/m"), "21/2") >= 0 Then
        Me.BackColor = &HC0FFC0
        Me.Caption = "Wizard of ChroniKeyland :: Happy Birthday to Puddleglum!"
    ElseIf DateDiff("d", Format(StartDate, "d/m"), "6/7") < 7 And DateDiff("d", Format(StartDate, "d/m"), "6/7") >= 0 Then
        Me.Caption = Me.Caption & " :-)"
        MsgBox "I am ChroniKey! You will obey me! Resistance is futile!", vbApplicationModal, "And so has the mighty ChroniKey spoken"
    ElseIf DateDiff("d", Format(StartDate, "d/m"), "1/4") < 7 And DateDiff("d", Format(StartDate, "d/m"), "1/4") >= 0 Then
        cmdGo.Picture = chkPosition.Picture
        cmdExit.Picture = chkPositionEnd.Picture
        cmdStop.Picture = cmdWorkShop.Picture
        cmdMinimize.Picture = cmdSimulate.Picture
    End If
    
    'Enabling AutoComplete feature
    SHAutoComplete txtSource.hwnd, SHACF_DEFAULT Or SHACF_USETAB Or SHACF_FILEALL
    
    'Loading the supplemental forms
    Load frmWorkShop
    Load frmDates
End Sub

Private Sub StatBarIdle_PanelClick(ByVal Panel As MSComctlLib.Panel)
    On Error Resume Next 'Required for debug; could be removed

    If Panel.Index = 3 Then
    
        Dim str1, str2
        
        str1 = StartDate
        str2 = EndDate
        If App.LogPath = "" Then 'We're in a debug session! LogPath cannot be empty
            StartDate = DateAdd("yyyy", -10, Date$)
            EndDate = DateAdd("yyyy", 10, Date$)
        Else
            StartDate = DateAdd("yyyy", -10, Date$)
            EndDate = DateAdd("yyyy", 10, Date$)
        End If
        With frmDates
            .lblCannot.Visible = False
            .txtMissing.Visible = False
            .DTPicker(1).MinDate = StartDate
            .DTPicker(1).MaxDate = EndDate
            .DTPicker(2).MinDate = StartDate
            .DTPicker(2).MaxDate = EndDate
            .DTPicker(1).Value = str1
            .DTPicker(2).Value = str2
            .Show vbModal, frmGlobus
            StartDate = .DTPicker(1).Value
            EndDate = .DTPicker(2).Value
        End With
UpdateLabel:
        Panel.Text = "From " & StartDate & " to " & EndDate
        Unload frmDates
        Load frmDates 'Reset it
        
    End If
End Sub

Private Sub txtEndAt_Change()
    If Len(txtEndAt.Text) > 0 And Not IsNumeric(txtEndAt.Text) Then
        ShowTextboxBalloonTip txtEndAt.hwnd, "Please enter numbers only here.", "Invalid Value", 3
    End If
End Sub

Private Sub txtSource_Change()
    cbxType.ListIndex = 0 'Reset it
End Sub

Private Sub txtSource_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ErrDragDrop
    txtSource.Text = Data.Files(1)
    If LCase(Right(txtSource.Text, 4)) <> ".htm" And LCase(Right(txtSource.Text, 5)) <> ".html" And LCase(Right(txtSource.Text, 4)) <> ".txt" Then
        ShoutAtUser frmGlobus, txtSource, "Please only drag here .htm, .html or .txt files.", "Error in Drag & Drop"
        txtSource.Text = ""
    End If
    Exit Sub
ErrDragDrop:
    ShoutAtUser frmGlobus, txtSource, "Please only drag here .htm, .html or .txt files.", "Error in Drag & Drop"
    'MsgBox "Please only drag here .htm, .html or .txt files.", vbExclamation + vbSystemModal, "Error in Drag & Drop"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, _
    y As Single)
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim msg As Long
    If cmdMinimize.Tag = "" Then Exit Sub
    msg = x / Screen.TwipsPerPixelX
    Select Case msg
        Case WM_LBUTTONDBLCLK
            If cmdMinimize.Tag = "2" Then 'for the NotificationIcon subroutine
                cmdMinimize.Tag = "1"
            Else
                Me.Show
                cmdMinimize.Tag = ""
                RemoveNotifyIcon
            End If
    End Select
    'FlashWindow frmGlobus.hwnd, False
End Sub

Private Sub cmdMinimize_Click()
    'Minimize to the tray area.

    'Set the individual values of the NOTIFYICONDATA data type
    'nID.cbSize = Len(nID)
    'nID.hwnd = frmGlobus.hwnd
    'nID.uID = vbNull
    'nID.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'nID.uCallbackMessage = WM_MOUSEMOVE
    'nID.hIcon = Me.Icon.Handle
    'nID.szTip = App.Title & " (dbl-click to open)" & vbNullChar
    
    'nID.dwInfoFlags = 0
    'nID.szInfo = vbNullChar
    'nID.uTimeOutOrVersion = 0
    'nID.szInfoTitle = vbNullChar
 
    'Call the Shell_NotifyIcon function to add the icon to the taskbar
    'status area
    'MsgBox Shell_NotifyIcon(NIM_ADD, nID)
    
    AddNotifyIcon Me, Me.Icon.Handle, App.Title
    
    'Hide the form
    Me.Hide
    
    'Update the tag
    cmdMinimize.Tag = "1"
    
    'Stop the form flashing, if applicable
    'FlashWindow frmGlobus.hwnd, False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim x, y As Long

    On Error Resume Next

    'Check whether we're not in the middle of something
    If Not cmdGo.Enabled Then
        Cancel = 1 'Don't allow an exit
        Exit Sub
    End If
    
    For y = Me.Height To 320 Step -35
        DoEvents
        Me.Move Me.Left, Me.Top + (-35 \ 2), Me.Width, y
    Next y
    For x = Me.Width To 2400 Step -35
        DoEvents
        Me.Move Me.Left + (-35 \ 2), Me.Top, x, Me.Height
    Next x

    Dim Form As Form
    For Each Form In Forms
       Unload Form
       Set Form = Nothing
    Next Form

End Sub

Private Sub txtStartAt_Change()
    If Len(txtStartAt.Text) > 0 And Not IsNumeric(txtStartAt.Text) Then
        ShowTextboxBalloonTip txtStartAt.hwnd, "Please enter numbers only here.", "Invalid Value", 3
    End If
End Sub
