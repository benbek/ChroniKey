VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkShop 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ChroniKey's Workshop"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frmWorkshop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update lists"
      Height          =   495
      Left            =   2760
      TabIndex        =   27
      ToolTipText     =   "Updates the lists (if you've just added a movie)"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame fraStartEnd 
      Caption         =   " Start/Stop Options "
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox cbxMovies 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1680
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   2895
      End
      Begin VB.ComboBox cbxMovies 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1680
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox chkEndMovie 
         Caption         =   "Sto&p at movie:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1950
         Width           =   1335
      End
      Begin VB.CheckBox chkStartMovie 
         Caption         =   "Sta&rt at movie:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1590
         Width           =   1335
      End
      Begin VB.ComboBox cbxCinemas 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1680
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "List of available cinemas"
         Top             =   640
         Width           =   2895
      End
      Begin VB.CheckBox chkEndCinema 
         Caption         =   "S&top at cinema:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   650
         Width           =   1455
      End
      Begin VB.ComboBox cbxCinemas 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1680
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "List of available cinemas"
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkStartCinema 
         Caption         =   "&Start at cinema:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   280
         Width           =   1455
      End
      Begin VB.Label lblInstruction 
         Caption         =   "You can only set Start and Stop Movie when the Start and Stop cinemas are the same ones."
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   4455
      End
   End
   Begin VB.TextBox txtLMNum 
      Height          =   285
      Left            =   1640
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "1"
      Top             =   2715
      Width           =   270
   End
   Begin VB.Frame fraRepairs 
      Caption         =   " Last-Minute Adds/Deletes "
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   4695
      Begin VB.TextBox txtDelete 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   25
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtAdd 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   19
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   73793536
         CurrentDate     =   37935
         MaxDate         =   2958464
      End
      Begin VB.ComboBox cbxMovies 
         Height          =   315
         Index           =   2
         Left            =   1560
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "List of available movies"
         Top             =   600
         Width           =   3015
      End
      Begin VB.ComboBox cbxCinemas 
         Height          =   315
         Index           =   2
         Left            =   1560
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "List of available cinemas"
         Top             =   240
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPickerEnd 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   21
         Top             =   1440
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   73793536
         CurrentDate     =   37935
         MaxDate         =   2958464
      End
      Begin VB.Label lblToDate 
         Caption         =   "To date&:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   1480
         Width           =   975
      End
      Begin VB.Label lblDelete 
         Caption         =   "De&lete the following hours:"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   24
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblAdd 
         Caption         =   "&Add the following hours:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblDate 
         Caption         =   "Fix in &date:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblMovie 
         Caption         =   "Fix for &movie:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   640
         Width           =   975
      End
      Begin VB.Label lblCinema 
         Caption         =   "Fix for &cinema:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   280
         Width           =   1335
      End
   End
   Begin MSComCtl2.UpDown UpDown 
      Height          =   285
      Left            =   1905
      TabIndex        =   12
      Top             =   2715
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtLMNum"
      BuddyDispid     =   196618
      OrigLeft        =   720
      OrigTop         =   240
      OrigRight       =   960
      OrigBottom      =   495
      Max             =   1
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   26
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblNumber 
      Caption         =   "&Edit correction(s) no.              for the data in the source file:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   4695
   End
End
Attribute VB_Name = "frmWorkShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentFrame%
Const MaxFrame As Integer = 20

Private Sub cmdOkay_Click()
    If DTPicker(CurrentFrame - 1).Value > DTPickerEnd(CurrentFrame - 1).Value Then
        ShoutAtUser frmWorkShop, frmWorkShop.txtLMNum, _
                    "The end date cannot be earlier that the start date.", "Invalid Input", _
                    vbExclamation + vbApplicationModal
    Else
        Me.Hide
    End If
End Sub

Private Sub cmdUpdate_Click()
   If DTPicker(CurrentFrame - 1).Value > DTPickerEnd(CurrentFrame - 1).Value Then
        ShoutAtUser frmWorkShop, frmWorkShop.txtLMNum, _
                    "The end date cannot be earlier that the start date.", "Invalid Input", _
                    vbExclamation + vbApplicationModal
    Else
        cmdUpdate.Tag = "7"
        Me.Hide
    End If
End Sub

Private Sub Form_Load()

    UpDown.Max = MaxFrame
    For CurrentFrame = 2 To MaxFrame
        Load fraRepairs(CurrentFrame) 'Load a new frame
        fraRepairs(CurrentFrame).ZOrder 1 'Set it on bottom of the current frame (the first frame should be displayed first)
        fraRepairs(CurrentFrame).Move fraRepairs(CurrentFrame - 1).Left, fraRepairs(CurrentFrame - 1).Top
        'Load the new controls for the frame
        Load lblCinema(CurrentFrame - 1) 'lblCinema(0) for fraRepairs(1)
        With lblCinema(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move lblCinema(0).Left, lblCinema(0).Top
            .Visible = True
        End With
        Load cbxCinemas(CurrentFrame + 1) 'cbxCinemas(2) for fraRepairs(1)
        With cbxCinemas(CurrentFrame + 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move cbxCinemas(2).Left, cbxCinemas(2).Top, cbxCinemas(2).Width
            .Visible = True
            .Enabled = True
        End With
        Load lblMovie(CurrentFrame - 1) 'lblMovie(0) for fraRepairs(1)
        With lblMovie(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move lblMovie(0).Left, lblMovie(0).Top
            .Visible = True
        End With
        Load cbxMovies(CurrentFrame + 1) 'cbxMovies(2) for fraRepairs(1)
        With cbxMovies(CurrentFrame + 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move cbxMovies(2).Left, cbxMovies(2).Top, cbxMovies(2).Width
            .Visible = True
            .Enabled = True
        End With
        Load lblDate(CurrentFrame - 1) 'lblDate(0) for fraRepairs(1)
        With lblDate(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move lblDate(0).Left, lblDate(0).Top
            .Visible = True
        End With
        Load DTPicker(CurrentFrame - 1)
        With DTPicker(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move DTPicker(0).Left, DTPicker(0).Top
            .MinDate = StartDate
            .MaxDate = EndDate
            .Visible = True
        End With
        Load lblToDate(CurrentFrame - 1) 'lblToDate(0) for fraRepairs(1)
        With lblToDate(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move lblToDate(0).Left, lblToDate(0).Top
            .Visible = True
        End With
        Load DTPickerEnd(CurrentFrame - 1)
        With DTPickerEnd(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move DTPickerEnd(0).Left, DTPickerEnd(0).Top
            .MinDate = StartDate
            .MaxDate = EndDate
            .Visible = True
        End With
        Load lblAdd(CurrentFrame - 1) 'lblAdd(0) for fraRepairs(1)
        With lblAdd(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move lblAdd(0).Left, lblAdd(0).Top
            .Visible = True
        End With
        Load txtAdd(CurrentFrame - 1) 'txtAdd(0) for fraRepairs(1)
        With txtAdd(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move txtAdd(0).Left, txtAdd(0).Top
            .Visible = True
        End With
        Load lblDelete(CurrentFrame - 1) 'lblDelete(0) for fraRepairs(1)
        With lblDelete(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move lblDelete(0).Left, lblDelete(0).Top
            .Visible = True
        End With
        Load txtDelete(CurrentFrame - 1) 'txtDelete(0) for fraRepairs(1)
        With txtDelete(CurrentFrame - 1)
            Set .Container = fraRepairs(CurrentFrame)
            .Move txtDelete(0).Left, txtDelete(0).Top
            .Visible = True
        End With
    Next CurrentFrame
    CurrentFrame = 1
    fraRepairs(CurrentFrame).ZOrder 0
End Sub

Private Sub UpDown_Change()
    If txtLMNum.Text = CurrentFrame Then Exit Sub 'No need to change frame.
    If DTPicker(CurrentFrame - 1).Value > DTPickerEnd(CurrentFrame - 1).Value Then
        ShoutAtUser frmWorkShop, frmWorkShop.txtLMNum, _
                    "The end date cannot be earlier that the start date.", "Invalid Input", _
                    vbExclamation + vbApplicationModal
        Exit Sub
    End If
   ' Otherwise, hide old frame, show new.
    With fraRepairs(CurrentFrame)
        .Visible = False
        .ZOrder 1
    End With
    With fraRepairs(txtLMNum.Text)
        .Visible = True
        .ZOrder 0
   End With
   ' Set CurrentFrame to new value.
   CurrentFrame = txtLMNum.Text
End Sub

Private Sub cbxCinemas_Click(Index As Integer)
    On Error Resume Next
    
    If cbxCinemas(Index).ListIndex = -1 Then Exit Sub
    If cbxCinemas(0).ListIndex = -1 Or cbxCinemas(1).ListIndex = -1 Then Exit Sub
    If cbxCinemas(0).ItemData(cbxCinemas(0).ListIndex) = cbxCinemas(1).ItemData(cbxCinemas(1).ListIndex) And chkStartCinema.Value = 1 And chkEndCinema.Value = 1 Then
        chkStartMovie.Enabled = True
        chkEndMovie.Enabled = True
    Else
        chkStartMovie.Enabled = False
        chkEndMovie.Enabled = False
    End If
End Sub

Private Sub chkEndCinema_Click()
    On Error Resume Next
    
    cbxCinemas(1).Enabled = Not cbxCinemas(1).Enabled
    If cbxCinemas(1).Enabled Then
        If cbxCinemas(1).ListIndex = -1 Then cbxCinemas(1).ListIndex = 0
        If cbxCinemas(0).ListIndex = -1 Then cbxCinemas(0).ListIndex = 0
        If cbxCinemas(0).ItemData(cbxCinemas(0).ListIndex) = cbxCinemas(1).ItemData(cbxCinemas(1).ListIndex) And chkStartCinema.Value = 1 Then
            chkStartMovie.Enabled = True
            chkEndMovie.Enabled = True
        End If
    Else
        chkStartMovie.Enabled = False
        chkEndMovie.Enabled = False
    End If
End Sub

Private Sub chkEndMovie_Click()
    On Error Resume Next
    
    cbxMovies(1).Enabled = Not cbxMovies(1).Enabled
    If cbxMovies(1).Enabled And cbxMovies(1).ListIndex = -1 Then cbxMovies(1).ListIndex = 0
End Sub

Private Sub chkStartCinema_Click()
    On Error Resume Next
    
    cbxCinemas(0).Enabled = Not cbxCinemas(0).Enabled
    If cbxCinemas(0).Enabled Then
        If cbxCinemas(0).ListIndex = -1 Then cbxCinemas(0).ListIndex = 0
        If cbxCinemas(1).ListIndex = -1 Then cbxCinemas(1).ListIndex = 0
        If cbxCinemas(0).ItemData(cbxCinemas(0).ListIndex) = cbxCinemas(1).ItemData(cbxCinemas(1).ListIndex) And chkEndCinema.Value = 1 Then
            chkStartMovie.Enabled = True
            chkEndMovie.Enabled = True
        End If
    Else
        chkStartMovie.Enabled = False
        chkEndMovie.Enabled = False
        cbxMovies(0).Enabled = False
        cbxMovies(1).Enabled = False
    End If
End Sub

Private Sub chkStartMovie_Click()
    On Error Resume Next
    
    cbxMovies(0).Enabled = Not cbxMovies(0).Enabled
    If cbxMovies(0).Enabled And cbxMovies(0).ListIndex = -1 Then cbxMovies(0).ListIndex = 0
End Sub

