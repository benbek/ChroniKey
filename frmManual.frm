VERSION 5.00
Begin VB.Form frmManual 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manually enter a "
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3510
   ControlBox      =   0   'False
   Icon            =   "frmManual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDo 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1088
      TabIndex        =   8
      ToolTipText     =   "Process the selected action"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame fraActions 
      Caption         =   " What do you want to &do? "
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3375
      Begin VB.ComboBox cbxOptions 
         Height          =   315
         ItemData        =   "frmManual.frx":0ECA
         Left            =   120
         List            =   "frmManual.frx":0EE3
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   270
         Width           =   3135
      End
      Begin VB.Label lblUpdate 
         Caption         =   "Choose 'update' to update the list (if you've just inserted to the database this new item)."
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1440
      Top             =   480
   End
   Begin VB.TextBox txtQuick 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Type here the first words of the missing item for quicker selection"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.ComboBox cbxCombo 
      Height          =   315
      ItemData        =   "frmManual.frx":0F79
      Left            =   120
      List            =   "frmManual.frx":0F7B
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txtMissing 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblOrder 
      Caption         =   "Please select the appropriate item from the following &list:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lblCannot 
      Caption         =   "Cannot analyze the serial number for the following "
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TextGotFocus As Boolean, ChangeOptions As Boolean

Private Sub cbxCombo_Change()
    If ChangeOptions Then
        cbxOptions.ListIndex = 0 'Changes the option to "save" if right on the start we found an entry (for the user to do less clicking)
        ChangeOptions = False
    End If
End Sub

Private Sub cmdDo_Click()
    If cbxOptions.ListIndex = -1 Then Exit Sub
    With Me
        If .cbxCombo.ListIndex = -1 Then
            If .cbxOptions.ListIndex < 2 Then 'require selecting an item
                .cbxCombo.SetFocus
                'ShoutAtUser frmManual, .cbxCombo, .Caption, -- doesn't work with combo-boxes
                MsgBox _
                     "Please choose the appropriate item from the highlighted list.", vbExclamation + vbApplicationModal, .Caption
                Exit Sub
            End If
        End If
        Select Case .cbxOptions.ListIndex
            Case 0:
                .Tag = "Save"
            Case 1:
                .Tag = "Cont"
            Case 2:
                .Tag = "Update"
            Case 3:
                .Tag = "Make" 'Text in list:   Make a new entry
            Case 4:
                .Tag = "Revive" 'Text in list: Revive an entry
            Case 5:
                .Tag = "Skip"
            Case 6:
                .Tag = "Stop"
        End Select
        FlashWindow frmManual.hwnd, False
        tmrTimer.Enabled = False
        TextGotFocus = False
        ChangeOptions = True
        .Hide
    End With
End Sub

Private Sub Form_Activate()
    FlashWindow frmManual.hwnd, False
    tmrTimer.Enabled = False
End Sub

Private Sub Form_GotFocus()
    FlashWindow frmManual.hwnd, False
    tmrTimer.Enabled = False
End Sub

Private Sub Form_Load()
    'MsgBox "Please complete the following form.", vbSystemModal + vbExclamation + vbOKOnly, "Attention Drawer"
    tmrTimer.Enabled = True
    TextGotFocus = False
    ChangeOptions = True
End Sub

Private Sub tmrTimer_Timer()
    FlashWindow frmManual.hwnd, True
End Sub

Private Sub txtQuick_Change()
    Dim LastTyped As Integer, SearchRes%, i%, Index%, strArray() As String, tempName$
    On Error Resume Next
    If txtQuick.Text = "" Then Exit Sub
    If cbxCombo.ListIndex <> -1 Then 'avoid a never-ending recursive function
        If txtQuick.Text = cbxCombo.List(cbxCombo.ListIndex) Then Exit Sub
    End If
    LastTyped = Len(txtQuick.Text)
    SearchRes = SendMessage(cbxCombo.hwnd, CB_FINDSTRING, -1, _
       ByVal txtQuick.Text)
       'CStr(txtQuick.Text))
    If SearchRes <> -1 Then 'Found something
        cbxCombo.ListIndex = SearchRes
        If Not cbxCombo.List(cbxCombo.ListIndex) = "" Then txtQuick.Text = cbxCombo.List(cbxCombo.ListIndex)
        txtQuick.SetFocus
        txtQuick.SelStart = LastTyped
        txtQuick.SelLength = Len(txtQuick.Text)
        ChangeOptions = False
        cbxOptions.ListIndex = 0
    Else 'Let's try modifying the text a little bit, trunctuating words and see if it's good
        If TextGotFocus Then Exit Sub 'Don't bother if the user wants to mess around
        strArray = Split(txtQuick.Text)
        For Index = UBound(strArray) - 1 To 0 Step -1
            tempName = ""
            For i = 0 To Index
                tempName = tempName & strArray(i) & " "
            Next i
            tempName = Trim(tempName)
            SearchRes = SendMessage(cbxCombo.hwnd, CB_FINDSTRING, -1, _
                    ByVal tempName)
            If SearchRes <> -1 Then 'Found an entry!
                cbxCombo.ListIndex = SearchRes
                If Not cbxCombo.List(cbxCombo.ListIndex) = "" Then txtQuick.Text = cbxCombo.List(cbxCombo.ListIndex)
                'The user did not enter the text in this situation, so there's no need to store or retrive the typing's status
                'txtQuick.SetFocus
                'txtQuick.SelStart = LastTyped
                'txtQuick.SelLength = Len(txtQuick.Text)
                ChangeOptions = False
                cbxOptions.ListIndex = 0
                Exit For
            End If
        Next Index
    End If
    If cbxOptions.ListIndex = -1 Then cbxOptions.ListIndex = 2 'Set to update, the reasonable action user would want to do
End Sub

Private Sub txtQuick_GotFocus()
    TextGotFocus = True
End Sub
