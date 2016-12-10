VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDates 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose Dates"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset Dates"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtMissing 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   200
      Width           =   2655
   End
   Begin VB.Frame fraTo 
      Caption         =   " &To date: "
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   2895
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24641536
         CurrentDate     =   38231
      End
   End
   Begin VB.Frame fraFrom 
      Caption         =   " &From date: "
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2895
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24641536
         CurrentDate     =   38231
      End
   End
   Begin VB.Label lblCannot 
      Caption         =   "Cannot parse the following date. Please select the appropriate date range for:"
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DoNotTouch As Boolean

Private Sub cmdOkay_Click()
    If DTPicker(1).Value > DTPicker(2).Value Then
        ShoutAtUser Me, Me.txtMissing, _
                    "The end date cannot be earlier that the start date.", "Invalid Input", _
                    vbExclamation + vbApplicationModal
    Else
        DoNotTouch = False
        Me.Hide
    End If
End Sub

Private Sub cmdReset_Click()
    Dim DateStart As Date, DateEnd As Date
    SetMovieWeek DateStart, DateEnd
    DTPicker(1).Value = DateStart
    DTPicker(2).Value = DateEnd
    DoNotTouch = False
End Sub

Private Sub DTPicker_Change(Index As Integer)
    If Index = 2 Then
        DoNotTouch = True
    Else
        If DoNotTouch = False Then
            If DateAdd("d", 6, DTPicker(1).Value) <= DTPicker(2).MaxDate Then
                DTPicker(2).Value = DateAdd("d", 6, DTPicker(1).Value)
            Else
                DTPicker(2).Value = DTPicker(2).MaxDate
            End If
        End If
    End If
End Sub

Private Sub DTPicker_GotFocus(Index As Integer)
    If Index = 2 Then
        DoNotTouch = True
    End If
End Sub

Private Sub Form_Load()
    'Dim intMaa
    '
    'For intMaa = 1 To 2
    '    DTPicker(intMaa).MinDate = StartDate
    '    DTPicker(intMaa).MaxDate = EndDate
    'Next intMaa
    
    DoNotTouch = False
End Sub
