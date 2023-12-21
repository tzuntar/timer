VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormMain 
   Caption         =   "Timer"
   ClientHeight    =   1680
   ClientLeft      =   7350
   ClientTop       =   5715
   ClientWidth     =   5295
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5295
   Begin VB.TextBox Text3 
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Text            =   "sec"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Text            =   "min"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "hours"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox HoursBox 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "0"
      Top             =   240
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar TimerProgressBar 
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   1000
   End
   Begin VB.Timer MainTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.CommandButton StartButton 
      Caption         =   "&Start"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox SecondsBox 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "30"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox MinutesBox 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label TimeLeftBox 
      Alignment       =   1  'Right Justify
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "remaining"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   270
      Width           =   495
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TimeLeftSeconds As Integer

Private Sub MainTimer_Timer()
    TimeLeftSeconds = TimeLeftSeconds - 1
    If TimeLeftSeconds <= 0 Then
        ResetTimer
        MsgBox "The time's up!", vbOKOnly, "Timer"
    Else
        TimeLeftBox.Caption = SecondsToTimeString(TimeLeftSeconds)
        TimerProgressBar.Value = TimerProgressBar.Value - 1
    End If
End Sub

Private Sub TimeLeftBox_Click()
    If MainTimer.Enabled And _
        MsgBox("Would you like to reset the timer?", vbYesNo) = vbYes Then
            ResetTimer
    End If
End Sub

Private Sub StartButton_Click()
    If Not MainTimer.Enabled Then
        If TimeLeftSeconds >= 1 Then
            StartButton.Caption = "&Pause"
            MainTimer.Enabled = True
        Else
            StartTimer
        End If
    Else
        MainTimer.Enabled = False
        StartButton.Caption = "&Resume"
    End If
End Sub

Private Sub StartTimer()
    TimeLeftSeconds = (HoursBox.Text * 60 * 60) _
        + (MinutesBox.Text * 60) + SecondsBox.Text
    TimerProgressBar.Max = TimeLeftSeconds
    TimerProgressBar.Value = TimeLeftSeconds
    TimeLeftBox.Caption = SecondsToTimeString(TimeLeftSeconds)
    StartButton.Caption = "&Pause"
    MainTimer.Enabled = True
End Sub

Private Sub ResetTimer()
    TimeLeftBox.Caption = "00:00:00"
    TimerProgressBar.Value = 0
    MainTimer.Enabled = False
    StartButton.Caption = "&Start"
    TimeLeftSeconds = 0
End Sub

Private Function SecondsToTimeString(Seconds As Integer)
    Dim Hours As Long, Minutes As Long, SecondsOut As Long
    Hours = Seconds \ 3600
    Minutes = (Seconds - (Hours * 3600)) \ 60
    SecondsOut = Seconds - ((Hours * 3600) + (Minutes * 60))
    SecondsToTimeString = Format(Hours, String(2, "0")) _
        & ":" & Format(Minutes, String(2, "0")) _
        & ":" & Format(SecondsOut, String(2, "0"))
End Function

