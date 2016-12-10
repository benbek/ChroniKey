VERSION 5.00
Begin VB.Form frmSkip 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Advanced Start/Stop Options"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOkay 
      Caption         =   "O&K"
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame fraStartEnd 
      Caption         =   " Start/Stop Options "
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CheckBox chkStartCinema 
         Caption         =   "&Start at cinema:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   280
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
         TabIndex        =   7
         ToolTipText     =   "List of available cinemas"
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkEndCinema 
         Caption         =   "S&top at cinema:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   650
         Width           =   1455
      End
      Begin VB.ComboBox cbxCinemas 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1680
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "List of available cinemas"
         Top             =   640
         Width           =   2895
      End
      Begin VB.CheckBox chkStartMovie 
         Caption         =   "Sta&rt at movie:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1590
         Width           =   1335
      End
      Begin VB.CheckBox chkEndMovie 
         Caption         =   "St&op at movie:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1950
         Width           =   1335
      End
      Begin VB.ComboBox cbxMovies 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1680
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   2895
      End
      Begin VB.ComboBox cbxMovies 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1680
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label lblInstruction 
         Caption         =   "You can only set Start and Stop Movie when the Start and Stop cinemas are the same ones."
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmSkip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()
    Me.Hide
End Sub

Private Sub cbxCinemas_Click(Index As Integer)
    On Error Resume Next
    
    If cbxCinemas(Index).ListIndex = -1 Then Exit Sub
    If cbxCinemas(0).ListIndex = -1 Or cbxCinemas(1).ListIndex = -1 Then Exit Sub
    If cbxCinemas(0).ItemData(cbxCinemas(0).ListIndex) = cbxCinemas(1).ItemData(cbxCinemas(1).ListIndex) Then
        chkStartMovie.Enabled = True
        chkEndMovie.Enabled = True
    Else
        chkStartMovie.Enabled = False
        chkEndMovie.Enabled = False
    End If
End Sub

Private Sub chkEndCinema_Click()
    cbxCinemas(1).Enabled = Not cbxCinemas(1).Enabled
    If cbxCinemas(1).Enabled Then
        If cbxCinemas(1).ListIndex = -1 Then cbxCinemas(1).ListIndex = 0
        If cbxCinemas(0).ItemData(cbxCinemas(0).ListIndex) = cbxCinemas(1).ItemData(cbxCinemas(1).ListIndex) And chkEndCinema.Enabled Then
            chkStartMovie.Enabled = True
            chkEndMovie.Enabled = True
        End If
    Else
        chkStartMovie.Enabled = False
        chkEndMovie.Enabled = False
    End If
End Sub

Private Sub chkEndMovie_Click()
    cbxMovies(1).Enabled = Not cbxMovies(1).Enabled
    If cbxMovies(1).Enabled And cbxMovies(1).ListIndex = -1 Then cbxMovies(1).ListIndex = 0
End Sub

Private Sub chkStartCinema_Click()
    cbxCinemas(0).Enabled = Not cbxCinemas(0).Enabled
    If cbxCinemas(0).Enabled Then
        If cbxCinemas(0).ListIndex = -1 Then cbxCinemas(0).ListIndex = 0
        If cbxCinemas(0).ItemData(cbxCinemas(0).ListIndex) = cbxCinemas(1).ItemData(cbxCinemas(1).ListIndex) And chkEndCinema.Enabled Then
            chkStartMovie.Enabled = True
            chkEndMovie.Enabled = True
        End If
    Else
        chkStartMovie.Enabled = False
        chkEndMovie.Enabled = False
    End If
End Sub

Private Sub chkStartMovie_Click()
    cbxMovies(0).Enabled = Not cbxMovies(0).Enabled
    If cbxMovies(0).Enabled And cbxMovies(0).ListIndex = -1 Then cbxMovies(0).ListIndex = 0
End Sub
