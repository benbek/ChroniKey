VERSION 5.00
Begin VB.Form frmSimulation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simulation"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   Icon            =   "frmSimulation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame fraSimulationOptions 
      Caption         =   " Simulation options "
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
      Begin VB.OptionButton optSimulation 
         Caption         =   "&Go the whole nine yards"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   920
         Width           =   3375
      End
      Begin VB.OptionButton optSimulation 
         Caption         =   "Simulate and then go the &whole nine yards"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   """Going"" will only take place if the simulation ends successfully"
         Top             =   600
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton optSimulation 
         Caption         =   "&Simulation only"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Label lblSimulationDescription 
      Caption         =   $"frmSimulation.frx":000C
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmSimulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pick%

Private Sub cmdCancel_Click()
    optSimulation(pick).Value = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If optSimulation(0).Value Then
        pick = 0
    ElseIf optSimulation(1).Value Then
        pick = 1
    Else
        pick = 2
    End If
    Me.Tag = pick
    
    Me.Hide
End Sub

Private Sub Form_Load()
    If optSimulation(0).Value Then
        pick = 0
    ElseIf optSimulation(1).Value Then
        pick = 1
    Else
        pick = 2
    End If
    Me.Tag = pick
End Sub
