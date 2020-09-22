VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmAlarmTimerStopWatch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                           Alarm Clock - Stop Watch - Timer"
   ClientHeight    =   5640
   ClientLeft      =   3615
   ClientTop       =   285
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TimerApplication.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   175
      Left            =   2625
      Picture         =   "TimerApplication.frx":57E82
      ScaleHeight     =   150
      ScaleWidth      =   300
      TabIndex        =   25
      Top             =   4515
      Visible         =   0   'False
      Width           =   325
   End
   Begin VB.PictureBox picTest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   175
      Left            =   2625
      Picture         =   "TimerApplication.frx":581B0
      ScaleHeight     =   150
      ScaleWidth      =   300
      TabIndex        =   24
      Top             =   4550
      Visible         =   0   'False
      Width           =   325
   End
   Begin VB.PictureBox picTestSound 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   175
      Left            =   2615
      ScaleHeight     =   150
      ScaleWidth      =   300
      TabIndex        =   23
      Top             =   4550
      Visible         =   0   'False
      Width           =   325
   End
   Begin VB.PictureBox picSetBtnDn 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   2250
      Picture         =   "TimerApplication.frx":584DE
      ScaleHeight     =   150
      ScaleWidth      =   300
      TabIndex        =   20
      Top             =   4750
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picSetBtnUp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   2250
      Picture         =   "TimerApplication.frx":58778
      ScaleHeight     =   150
      ScaleWidth      =   300
      TabIndex        =   19
      Top             =   4750
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picSetBtn 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   2250
      Picture         =   "TimerApplication.frx":58A12
      ScaleHeight     =   150
      ScaleWidth      =   300
      TabIndex        =   18
      ToolTipText     =   "Change Settings"
      Top             =   4750
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picGrnBtnUp 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   930
      Left            =   3875
      Picture         =   "TimerApplication.frx":58CAC
      ScaleHeight     =   930
      ScaleWidth      =   900
      TabIndex        =   17
      Top             =   860
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picRedBtnUp 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   860
      Left            =   2140
      Picture         =   "TimerApplication.frx":5B886
      ScaleHeight     =   855
      ScaleWidth      =   615
      TabIndex        =   16
      Top             =   60
      Visible         =   0   'False
      Width           =   620
   End
   Begin VB.PictureBox picBlkBtnUp 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      Height          =   950
      Left            =   60
      Picture         =   "TimerApplication.frx":5D464
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   15
      Top             =   770
      Visible         =   0   'False
      Width           =   950
   End
   Begin VB.PictureBox picGrnBtnCocked 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   930
      Left            =   3875
      Picture         =   "TimerApplication.frx":603E6
      ScaleHeight     =   930
      ScaleWidth      =   900
      TabIndex        =   14
      Top             =   860
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picRedBtnCocked 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   860
      Left            =   2140
      Picture         =   "TimerApplication.frx":62FC0
      ScaleHeight     =   855
      ScaleWidth      =   615
      TabIndex        =   13
      Top             =   60
      Visible         =   0   'False
      Width           =   620
   End
   Begin VB.PictureBox picBlkBtnCocked 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      Height          =   950
      Left            =   60
      Picture         =   "TimerApplication.frx":64B9E
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   12
      Top             =   770
      Visible         =   0   'False
      Width           =   950
   End
   Begin VB.PictureBox picRedBtn 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   860
      Left            =   2140
      Picture         =   "TimerApplication.frx":67B20
      ScaleHeight     =   855
      ScaleWidth      =   615
      TabIndex        =   11
      ToolTipText     =   "Stop Watch Settings"
      Top             =   60
      Visible         =   0   'False
      Width           =   620
   End
   Begin VB.PictureBox picGrnBtn 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   930
      Left            =   3875
      Picture         =   "TimerApplication.frx":696FE
      ScaleHeight     =   930
      ScaleWidth      =   900
      TabIndex        =   10
      ToolTipText     =   "Timer Settings"
      Top             =   860
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picBlkBtn 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      Height          =   950
      Left            =   60
      Picture         =   "TimerApplication.frx":6C2D8
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   9
      ToolTipText     =   "Alarm Settings"
      Top             =   770
      Visible         =   0   'False
      Width           =   950
   End
   Begin VB.Timer tmrCountDown 
      Interval        =   100
      Left            =   4305
      Top             =   0
   End
   Begin VB.PictureBox picSoundOn 
      BorderStyle     =   0  'None
      Height          =   520
      Left            =   2175
      Picture         =   "TimerApplication.frx":6F25A
      ScaleHeight     =   525
      ScaleWidth      =   540
      TabIndex        =   6
      ToolTipText     =   "Turn off Watch Tick"
      Top             =   975
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picSoundOff 
      BorderStyle     =   0  'None
      Height          =   520
      Left            =   2175
      Picture         =   "TimerApplication.frx":70160
      ScaleHeight     =   525
      ScaleWidth      =   540
      TabIndex        =   5
      ToolTipText     =   "Turn on Watch Tick"
      Top             =   975
      Width           =   540
   End
   Begin VB.Timer tmrNormal 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picGrnBtnDn 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   930
      Left            =   3875
      Picture         =   "TimerApplication.frx":71066
      ScaleHeight     =   930
      ScaleWidth      =   900
      TabIndex        =   3
      Top             =   860
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picBlkBtnDn 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      Height          =   950
      Left            =   60
      Picture         =   "TimerApplication.frx":73C40
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   770
      Visible         =   0   'False
      Width           =   950
   End
   Begin VB.PictureBox picRedBtnDn 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   860
      Left            =   2140
      Picture         =   "TimerApplication.frx":76BC2
      ScaleHeight     =   855
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   620
   End
   Begin VB.PictureBox picAlarmOn 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1890
      Picture         =   "TimerApplication.frx":787A0
      ScaleHeight     =   225
      ScaleWidth      =   315
      TabIndex        =   22
      ToolTipText     =   "Current Picture Of Sound"
      Top             =   4515
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblStoppedTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   190
      Left            =   1557
      TabIndex        =   26
      Top             =   4035
      Width           =   1700
   End
   Begin VB.Line linTimeHands 
      BorderColor     =   &H00FF0000&
      Index           =   1
      X1              =   2400
      X2              =   2400
      Y1              =   3225
      Y2              =   1675
   End
   Begin VB.Shape shpCentre 
      FillStyle       =   0  'Solid
      Height          =   60
      Left            =   2375
      Shape           =   3  'Circle
      Top             =   3225
      Width           =   60
   End
   Begin VB.Shape shpfinial 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   2325
      Shape           =   3  'Circle
      Top             =   3175
      Width           =   150
   End
   Begin VB.Line linTimeHands 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   3225
      Y2              =   1675
   End
   Begin VB.Line linTimeHands 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   2
      X1              =   2400
      X2              =   2400
      Y1              =   3225
      Y2              =   1875
   End
   Begin VB.Line linTimeHands 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   10
      Index           =   3
      X1              =   2400
      X2              =   2400
      Y1              =   3225
      Y2              =   2075
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
      Height          =   330
      Left            =   4200
      TabIndex        =   21
      Top             =   5145
      Visible         =   0   'False
      Width           =   435
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -1240
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblSetInstruction 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   170
      Left            =   1607
      TabIndex        =   8
      Top             =   3825
      Width           =   1600
   End
   Begin VB.Label lblActionType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1717
      TabIndex        =   7
      Top             =   3360
      Width           =   1380
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   330
      Left            =   210
      TabIndex        =   4
      Top             =   5145
      Visible         =   0   'False
      Width           =   435
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -1240
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblTimeReadout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 PM"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1455
      TabIndex        =   0
      ToolTipText     =   "Analog Readout"
      Top             =   3570
      Width           =   1905
   End
End
Attribute VB_Name = "frmAlarmTimerStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnBlkBtnDn As Boolean            ' Was the Black (Alarm Settings) Button Clicked?
Dim blnRedBtnDn As Boolean            ' Was the Red (Stop Watch) Button Clicked?
Dim blnGrnBtnDn As Boolean            ' Was the Green (Timer) Button Clicked?

Dim intNumBlkClicks As Integer        ' How many Black Button clicks?
Dim intNumGrnClicks As Integer        ' How many Green Button clicks?
Dim intNumRedClicks As Integer        ' How many Red Button clicks?
Dim intblink As Integer               ' Blinking display toggle

Dim blnTimeDisplay As Boolean         ' Whether or not normal time is displayed
Dim blnAlarmOn As Boolean             ' Whether or not alarm is turned on
Dim blnStopTime As Boolean            ' Whether or not Stop Time is running
Dim blnTestOn As Boolean              ' Whether or not alarm sound is running
    
Dim intHandLength(4) As Integer       ' Lengths of the watch hands

Dim strSetReadout As String           ' Digital Timer readout (hh:mm:ss xm)
Dim strSetAlarm As String             ' Alarm "ON/OFF" readout
Dim strSetStop As String              ' Stop Watch sound setting "ON/OFF"
Dim strAlarmSound As String           ' Alarm .wav file sound name readout
Dim strStopSound As String            ' Stop Watch .wav file sound name readout
Dim strAlarmTime As String            ' Alarm time storage string (hh:mm:ss xm)
Dim strStopTime As String             ' Stop timer total time
Dim strLastSecond As String           ' Check against during Elapsed Time

Const PI As Double = 3.14159265 / 30  ' PI factor for calcuating positions of watch hands
  
Private Sub Form_Load()

  Dim I As Integer                    ' Loop counter
  Dim IntDeltaX As Integer            ' Difference between X'x in a watch hand
  Dim IntDeltaY As Integer            ' Difference between X'x in a watch hand
  
  ' Calculate the length of the watch hands
  '
  For I = 0 To 3
    IntDeltaX = linTimeHands(I).X1 - linTimeHands(I).X2
    IntDeltaY = linTimeHands(I).Y1 - linTimeHands(I).Y2
    intHandLength(I) = Sqr(IntDeltaX ^ 2 + IntDeltaY ^ 2)
  Next I
  
  ' Centre the watch form on the screen
  '
  frmAlarmTimerStopWatch.Left = (Screen.Width - frmAlarmTimerStopWatch.Width) / 2
  frmAlarmTimerStopWatch.Top = (Screen.Height - frmAlarmTimerStopWatch.Height) / 2
  
  ' Start the timer and load the ticker sound
  '
  blnTimeDisplay = True
  tmrNormal.Enabled = True
  MediaPlayer1.Open (App.Path & "\" & "tick.wav")

  ' Initialize the settings for alarm/stopwatch/timer
  '
  MediaPlayer2.Open (App.Path & "\" & "Alarm_Clock.wav")
  strAlarmTime = "00:00:00 AM"
  strStopTime = "00:00:00 ST"
  strAlarmSound = "Alarm_Clock"
  strStopSound = "Alarm_Clock"
  strSetAlarm = "OFF"
  strSetStop = "OFF"
  intblink = 1
  
  ' Make the Function buttons visible, enabled and in up position
  '
  picBlkBtn.Enabled = True
  picBlkBtn.Visible = True
  picBlkBtn.Picture = picBlkBtnUp.Picture
  
  picGrnBtn.Enabled = True
  picGrnBtn.Visible = True
  picGrnBtn.Picture = picGrnBtnUp.Picture
  
  picRedBtn.Enabled = True
  picRedBtn.Visible = True
  picRedBtn.Picture = picRedBtnUp.Picture
  
End Sub


Private Sub picBlkBtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' As soon as the Black button is clicked, flag the turning off of normal time display and
  ' show the Black button depressing. Flag that it is locked in down position
  '
  blnTimeDisplay = False
  picBlkBtn.Picture = picBlkBtnDn.Picture
  blnBlkBtnDn = True
  
  ' Any other locked down button returns back to the up postion
  '
  picRedBtn.Picture = picRedBtnUp.Picture
  intNumRedClicks = 0
  picGrnBtn.Picture = picGrnBtnUp.Picture
  intNumGrnClicks = 0

End Sub

Private Sub picBlkBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  intNumBlkClicks = intNumBlkClicks + 1               ' Track number of Black button clicks
                                                      ' and process accordingliy
  Select Case intNumBlkClicks
    Case 1                                            '======= FIRST CLICK ON BLACK ============
      picBlkBtn.Picture = picBlkBtnCocked.Picture     ' Black Button is in locked down position
      DisplayHands (False)                            ' Hide the watch hands for now
      lblActionType.Caption = "Set Alarm"             ' Indicate that Alarm is being set
      picSetBtn.Visible = True                        ' Show the option (red) setting button
      lblSetInstruction.Caption = "Hour"              ' Setting the Alarm HOUR at this point
      strSetReadout = strAlarmTime                    ' Use last Alarm setting if one exists
      picRedBtn.Enabled = False                       ' Disable the Red button
      picGrnBtn.Enabled = False                       ' Disable the Green button
    Case 2                                            '======= SECOND CLICK ON BLACK ===========
      picBlkBtn.Picture = picBlkBtnCocked.Picture     ' Black Button is in locked down position
      lblSetInstruction.Caption = "Minute"            ' Setting the Alarm MINUTE at this point
    Case 3                                            '======= THIRD CLICK ON BLACK ============
      picBlkBtn.Picture = picBlkBtnCocked.Picture     ' Black Button is in locked down position
      lblSetInstruction.Caption = "Second"            ' Setting the Alarm SECOND at this point
    Case 4                                            '======= FORTH CLICK ON BLACK ============
      picBlkBtn.Picture = picBlkBtnCocked.Picture     ' Black Button is in locked down position
      lblSetInstruction.Caption = "AM/PM"             ' Indicate setting of AM/PM at this point
    Case 5                                            '======= FIFTH CLICK ON BLACK ============
      picBlkBtn.Picture = picBlkBtnCocked.Picture     ' Black Button is in locked down position
      If (Left$(strSetReadout, 1) = "0") Then
        strAlarmTime = Right$(strSetReadout, 10)      ' Store the set Alarm time string
      Else
        strAlarmTime = strSetReadout
      End If
      strSetReadout = strSetAlarm                     ' Display the current ON/OFF setting
      lblSetInstruction.Caption = "ON/OFF"            ' Indicate setting of ON/OFF at this point
    Case 6                                            '======= SIXTH CLICK ON BLACK ============
      picBlkBtn.Picture = picBlkBtnCocked.Picture     ' Black Button is in locked down position
      strSetAlarm = strSetReadout                     ' Store the Alarm ON/OFF setting
      If (strSetAlarm = "OFF") Then
        picSetBtn.Visible = False                     ' Hide the option setting button
        strSetReadout = "Alarm is OFF"                ' Can't set sound with alarm set to OFF
      Else
        lblSetInstruction.Caption = "Sound"           ' Indicate setting sound at this point
        strSetReadout = strAlarmSound                 ' Display the current Alarm sound setting
        picTestSound.Picture = picTest.Picture        ' Indicate ready to "test" sound
        picTestSound.ToolTipText = "Test The Sound"   ' Set the tool tip text
        picTestSound.Visible = True                   ' Show the sound test button
      End If
    Case 7                                            '======== LAST CLICK ON BLACK ============
      picBlkBtn.Picture = picBlkBtnUp.Picture         ' Black Button returns to up position
      If (strSetAlarm = "ON") Then
        strAlarmSound = strSetReadout                 ' Store the newly set Alarm sound
      End If
      DisplayHands (True)                             ' Show the watch hands again
      picTestSound.Visible = False                    ' Hide the sound test button
      MediaPlayer2.Stop                               ' Stop (if still) playing alarm sound
      picSetBtn.Visible = False                       ' Hide the option setting button
      lblActionType.Caption = ""                      ' Clear the Indicator labels
      lblSetInstruction.Caption = ""                  '   ""          ""       ""
      blnBlkBtnDn = False                             ' Flag that the Black Button is up
      intNumBlkClicks = 0                             ' Number of Black clicks returns to zero
      blnTimeDisplay = True                           ' Flag Return to normal time display
      picRedBtn.Enabled = True                        ' Enable the Red button again
      picGrnBtn.Enabled = True                        ' Enable the Green button again
  End Select

End Sub

Private Sub picGrnBtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' As soon as the Green button is clicked, flag the turning off of normal time display and
  ' show the Green button depressing. Flag that it is locked in down position
  '
  blnTimeDisplay = False
  picGrnBtn.Picture = picGrnBtnDn.Picture
  blnGrnBtnDn = True
  
  ' Any other locked down button returns back to the up postion
  '
  picBlkBtn.Picture = picBlkBtnUp.Picture
  intNumBlkClicks = 0
  picRedBtn.Picture = picRedBtnUp.Picture
  intNumRedClicks = 0

End Sub

Private Sub picGrnBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  intNumGrnClicks = intNumGrnClicks + 1               ' Track number of Green button clicks
                                                      ' and process accordingliy
  Select Case intNumGrnClicks
    Case 1                                            '======= FIRST CLICK ON GREEN ============
      picGrnBtn.Picture = picGrnBtnCocked.Picture     ' Green Button is in locked down position
      lblActionType.Caption = "Elasped Time"          ' Indicate that Elapsed Timer has started
      lblSetInstruction.Caption = "Started at " & Time ' Show the starting time
      strSetReadout = "00:00:00"                      ' Initialize the elapsed timer readout
      strLastSecond = Left$(Right$(Time, 4), 2)       ' Capture the current second
      picBlkBtn.Enabled = False                       ' Disable the Black button
      picRedBtn.Enabled = False                       ' Disable the Red button
    Case 2                                            '======= SECOND CLICK ON GREEN ===========
      picGrnBtn.Picture = picGrnBtnCocked.Picture     ' Green Button is in locked down position
      lblStoppedTime.Caption = " Stopped at " & Time  ' Show the ending time
    Case 3                                            '======== LAST CLICK ON GREEN ============
      picGrnBtn.Picture = picGrnBtnUp.Picture         ' Green Button returns to up position
      lblActionType.Caption = ""                      ' Clear the Indicator labels
      lblSetInstruction.Caption = ""                  '    ""       ""        ""
      lblStoppedTime.Caption = ""                     '    ""       ""        ""
      blnGrnBtnDn = False                             ' Flag that the Green Button is up
      intNumGrnClicks = 0                             ' Number of Black clicks returns to zero
      blnTimeDisplay = True                           ' Flag Return to normal time display
      picBlkBtn.Enabled = True                        ' Enable the Black button
      picRedBtn.Enabled = True                        ' Enable the Red button
  End Select

End Sub

Private Sub picRedBtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' As soon as the Red button is clicked, flag the turning off of normal time display and
  ' show the Red button depressing. Flag that it is locked in down position
  
  blnTimeDisplay = False
  picRedBtn.Picture = picRedBtnDn.Picture
  blnRedBtnDn = True
  
  ' Any other locked down button returns back to the up postion
  '
  picBlkBtn.Picture = picBlkBtnUp.Picture
  intNumBlkClicks = 0
  picGrnBtn.Picture = picGrnBtnUp.Picture
  intNumGrnClicks = 0

End Sub

Private Sub picRedBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  intNumRedClicks = intNumRedClicks + 1               ' Track number of Red button clicks
                                                      ' and process accordingliy
  Select Case intNumRedClicks
    Case 1                                            '========= FIRST CLICK ON RED ============
      picRedBtn.Picture = picRedBtnCocked.Picture     ' Red Button is in locked down position
      DisplayHands (False)                            ' Hide the watch hands
      lblActionType.Caption = "Set Stop Watch"        ' Indicate that Stop Watch is being set
      picSetBtn.Visible = True                        ' Show the option (red) setting button
      lblSetInstruction.Caption = "Hour"              ' Set the Stop Watch HOUR at this point
      strSetReadout = strStopTime                     ' Use last Stop Watch setting if it exists
      picBlkBtn.Enabled = False                       ' Disable the Black button
      picGrnBtn.Enabled = False                       ' Disable the Green button
    Case 2                                            '========= SECOND CLICK ON RED ===========
      picRedBtn.Picture = picRedBtnCocked.Picture     ' Red Button is in locked down position
      lblSetInstruction.Caption = "Minute"            ' Set the Stop Watch MINUTE at this point
    Case 3                                            '========= THIRD CLICK ON RED ============
      picRedBtn.Picture = picRedBtnCocked.Picture     ' Red Button is in locked down position
      lblSetInstruction.Caption = "Second"            ' Set the Stop Watch SECOND at this point
    Case 4                                            '========= FORTH CLICK ON RED ============
      strStopTime = strSetReadout                     ' Store the set Stop time string
      picRedBtn.Picture = picRedBtnCocked.Picture     ' Red Button is in locked down position
      strSetReadout = strSetStop                      ' Display stop sound ON/OFF setting
      lblSetInstruction.Caption = "Sound ON/OFF"      ' Indicate sound ON/OFF at this point
    Case 5                                            '========= FIFTH CLICK ON RED ============
      picRedBtn.Picture = picRedBtnCocked.Picture     ' Red Button is in locked down position
      strSetStop = strSetReadout                      ' Store the Stop Watch ON/OFF setting
      If (strSetStop = "OFF") Then
        picSetBtn.Visible = False                     ' Hide the option setting button
        strSetReadout = "Stop Sound is OFF"           ' Can't set sound with alarm set to OFF
      Else
        strSetReadout = strStopSound                  ' Display the current Stop Watch sound
        lblSetInstruction.Caption = "Sound"           ' Indicate setting sound at this point
        picTestSound.Picture = picTest.Picture        ' Indicate ready to "test" sound
        picTestSound.Visible = True                   ' Show the sound test button
      End If
    Case 6                                            '========= SIXTH CLICK ON RED ============
      picRedBtn.Picture = picRedBtnCocked.Picture     ' Red Button is in locked down position
      If (strSetStop = "ON") Then
        strStopSound = strSetReadout                  ' Store the newly set Stop sound
      End If
      picTestSound.Visible = False                    ' Hide the sound test button
      MediaPlayer2.Stop                               ' Stop (if still) playing alarm sound
      picSetBtn.Visible = False                       ' Hide the option setting button
      lblActionType.Caption = ""                      ' Clear the Indicator labels
      lblSetInstruction.Caption = ""                  '   ""          ""       ""
      strSetReadout = "READY"                         ' Display "READY" for count down
    Case 7                                            '======== SEVENTH CLICK ON RED ===========
      picRedBtn.Picture = picRedBtnCocked.Picture     ' Red Button is in locked down position
      strSetReadout = "SET"                           ' Display "SET" for count down
    Case 8                                            '========== LAST CLICK ON RED ===========
      picRedBtn.Picture = picRedBtnCocked.Picture     ' Red Button is in locked down position
      DisplayHands (True)                             ' Show the watch hands again
      strSetReadout = Left$(strStopTime, 8)           ' Display Stop Time readout
      strLastSecond = Left$(Right$(Time, 4), 2)       ' Capture the current second
      blnStopTime = True                              ' Begin Stop Timer
      picRedBtn.Enabled = False                       ' Disable the Red button until time is up
  
  End Select

End Sub

Private Sub picSetBtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Show the option setting button depressing while the mouse is down
  picSetBtn.Picture = picSetBtnDn.Picture
End Sub

Private Sub picSetBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Show the option setting button return to the up position when mouse click finishes
  picSetBtn.Picture = picSetBtnUp.Picture
  Select Case True
    Case blnBlkBtnDn          ' If the BLACK Button is locked in the down position
      SetAlarmTime            '   then we are setting the Alarm Options
    Case blnRedBtnDn          ' If the RED Button is Locked in the down position
      SetStopTime             '   then we are setting the Stop Timer
  End Select
End Sub

Private Sub picSoundOff_Click()

  ' When Sound Off image is clicked change to Sound On image and turn on ticker
  '
  picSoundOff.Visible = False
  picSoundOn.Visible = True
  MediaPlayer1.Play
  
End Sub

Private Sub picSoundOn_Click()
  
  ' When Sound On image is clicked change to Sound Off image and turn off ticker
  '
  picSoundOn.Visible = False
  picSoundOff.Visible = True
  MediaPlayer1.Stop
  MediaPlayer1.FastReverse

End Sub

Private Sub picTestSound_Click()
  If Not (blnTestOn) Then
    blnTestOn = True
    MediaPlayer2.Play
    picTestSound.Picture = picDone.Picture
    picTestSound.ToolTipText = "Turn off Sound Test"
    picSetBtn.Enabled = False
    picBlkBtn.Enabled = False
  Else
    blnTestOn = False
    MediaPlayer2.Stop
    MediaPlayer2.FastReverse
    picTestSound.ToolTipText = "Test The Sound"
    picTestSound.Visible = False
    picSetBtn.Enabled = True
    picBlkBtn.Enabled = True
    If (blnTimeDisplay) Then
      blnAlarmOn = False
      picAlarmOn.Visible = False
      strSetAlarm = "OFF"
    End If
  End If
End Sub

Private Sub tmrNormal_Timer()
  
  ' For each timer interval show digital/analog times for either the normal time or ..
  '
  If (blnTimeDisplay) Then
    lblTimeReadout.Caption = Time
    ClockNormalTime
    If (strAlarmTime = lblTimeReadout.Caption) Then
      If (blnAlarmOn) Then
        MediaPlayer2.Play
        picTestSound.Picture = picDone.Picture
        picTestSound.Visible = True
        blnTestOn = True
      End If
    End If
  Else      ' ... or show the Stop Watch time or Timer time or don't display time
    Select Case True
      Case blnBlkBtnDn
        lblTimeReadout.Caption = BlinkingAlarmReadout(strSetReadout, intNumBlkClicks)
      Case blnGrnBtnDn
        If ((intNumGrnClicks = 1) Or (intNumGrnClicks = 2)) Then
          If (intNumGrnClicks = 1) Then
            MakeElapsedTime
            ClockOtherTime
          End If
          lblTimeReadout.Caption = strSetReadout & " TE"
        End If
      Case blnRedBtnDn
        If (blnStopTime) Then
          If (strSetReadout = "00:00:00") Then
            picRedBtn.Picture = picRedBtnUp.Picture         ' Red Button returns to up position
            DisplayHands (True)                             ' Show the watch hands again
            blnRedBtnDn = False                             ' Flag that the Red Button is up
            intNumGrnClicks = 0                             ' Number of Red clicks returns to zero
            blnTimeDisplay = True                           ' Flag Return to normal time display
            picRedBtn.Enabled = True                        ' Enable the Red button again
            picBlkBtn.Enabled = True                        ' Enable the Black button
            picGrnBtn.Enabled = True                        ' Enable the Green button
            MediaPlayer2.Play
            picTestSound.Picture = picDone.Picture
            picTestSound.Visible = True
            blnTestOn = True
          Else
            MakeStopTime
            ClockOtherTime
            lblTimeReadout.Caption = strSetReadout & " ST"
          End If
        Else
          lblTimeReadout.Caption = BlinkingStopReadout(strSetReadout, intNumRedClicks)
        End If
    End Select
  End If

End Sub

Private Function BlinkingAlarmReadout(strTimeString As String, intNumClicks As Integer) As String
  intblink = intblink * (-1)
  Select Case intNumClicks
    Case 1
      If (intblink < 0) Then
        BlinkingAlarmReadout = "__" & Right$(strTimeString, 9)
      Else
        BlinkingAlarmReadout = strTimeString
      End If
    Case 2
      If (intblink < 0) Then
        BlinkingAlarmReadout = Left$(strTimeString, 3) & "__" & Right$(strTimeString, 6)
      Else
        BlinkingAlarmReadout = strTimeString
      End If
    Case 3
      If (intblink < 0) Then
        BlinkingAlarmReadout = Left$(strTimeString, 6) & "__" & Right$(strTimeString, 3)
      Else
        BlinkingAlarmReadout = strTimeString
      End If
    Case 4
      If (intblink < 0) Then
        BlinkingAlarmReadout = Left$(strTimeString, 9) & "__"
      Else
        BlinkingAlarmReadout = strTimeString
      End If
    Case Else
      If (intblink < 0) Then
        BlinkingAlarmReadout = ""
      Else
        BlinkingAlarmReadout = strTimeString
      End If
  End Select
End Function

Private Function BlinkingStopReadout(strTimeString As String, intNumClicks As Integer) As String
  intblink = intblink * (-1)
  Select Case intNumClicks
    Case 1
      If (intblink < 0) Then
        BlinkingStopReadout = "__" & Right$(strTimeString, 9)
      Else
        BlinkingStopReadout = strTimeString
      End If
    Case 2
      If (intblink < 0) Then
        BlinkingStopReadout = Left$(strTimeString, 3) & "__" & Right$(strTimeString, 6)
      Else
        BlinkingStopReadout = strTimeString
      End If
    Case 3
      If (intblink < 0) Then
        BlinkingStopReadout = Left$(strTimeString, 6) & "__" & Right$(strTimeString, 3)
      Else
        BlinkingStopReadout = strTimeString
      End If
    Case Else
      If (intblink < 0) Then
        BlinkingStopReadout = ""
      Else
        BlinkingStopReadout = strTimeString
      End If
  End Select
End Function

Private Sub ClockNormalTime()
  
  Dim dblHandAngle As Double, Seconds As Double, Minutes As Double
  
  ' Calculate the position of the second hand
  '
  dblHandAngle = PI * (Second(Time) - 15)
  linTimeHands(2).X2 = linTimeHands(2).X1 + Cos(dblHandAngle) * intHandLength(2)
  linTimeHands(2).Y2 = linTimeHands(2).Y1 + Sin(dblHandAngle) * intHandLength(2)

  ' Calculate the position of the minute hand
  '
  Seconds = Second(Time) / 60
  dblHandAngle = PI * (Minute(Time) + Seconds - 15)
  linTimeHands(0).X2 = linTimeHands(0).X1 + Cos(dblHandAngle) * intHandLength(0)
  linTimeHands(0).Y2 = linTimeHands(0).Y1 + Sin(dblHandAngle) * intHandLength(0)
  linTimeHands(1).X2 = linTimeHands(1).X1 + Cos(dblHandAngle) * intHandLength(1)
  linTimeHands(1).Y2 = linTimeHands(1).Y1 + Sin(dblHandAngle) * intHandLength(1)

  ' Calculate the position of the hour hand
  '
  Minutes = Minute(Time) / 60
  dblHandAngle = (PI * 30) / 6 * (Hour(Time) + Minutes - 15)
  linTimeHands(3).X2 = linTimeHands(3).X1 + Cos(dblHandAngle) * intHandLength(3)
  linTimeHands(3).Y2 = linTimeHands(3).Y1 + Sin(dblHandAngle) * intHandLength(3)

End Sub

Private Function SetStopTime()
  
  Dim strtemp As String
  Select Case intNumRedClicks
    Case 1
      strtemp = Left$(strSetReadout, 2)
      strSetReadout = Increment(strtemp, 13) & Right$(strSetReadout, 9)
    Case 2
      strtemp = Mid$(strSetReadout, 4, 2)
      strSetReadout = Left$(strSetReadout, 3) & Increment(strtemp, 61) & Right$(strSetReadout, 6)
    Case 3
      strtemp = Mid$(strSetReadout, 7, 2)
      strSetReadout = Left$(strSetReadout, 6) & Increment(strtemp, 61) & Right$(strSetReadout, 3)
    Case 4
      If (blnAlarmOn) Then
        picAlarmOn.Visible = False
        blnAlarmOn = False
        strSetReadout = "OFF"
      Else
        picAlarmOn.Visible = True
        blnAlarmOn = True
        strSetReadout = "ON"
      End If
    Case 5
      If Not (blnTestOn) Then
        picTestSound.Visible = True
        picTestSound.Picture = picTest.Picture
      End If
      If (strSetReadout = "Alarm_Clock") Then
        strSetReadout = "Rooster"
      ElseIf (strSetReadout = "Rooster") Then
        strSetReadout = "Siren"
      ElseIf (strSetReadout = "Siren") Then
        strSetReadout = "Explosion"
      ElseIf (strSetReadout = "Explosion") Then
        strSetReadout = "Air_Horn"
      ElseIf (strSetReadout = "Air_Horn") Then
        strSetReadout = "Warning"
      ElseIf (strSetReadout = "Warning") Then
        strSetReadout = "Alarm_Clock"
      End If
      picAlarmOn.Picture = LoadPicture(App.Path & "\" & strSetReadout & ".bmp")
      MediaPlayer2.Open (App.Path & "\" & strSetReadout & ".wav")
  End Select

End Function

Private Sub SetAlarmTime()
  Dim strtemp As String
  Select Case intNumBlkClicks
    Case 1
      strtemp = Left$(strSetReadout, 2)
      strSetReadout = Increment(strtemp, 13) & Right$(strSetReadout, 9)
    Case 2
      strtemp = Mid$(strSetReadout, 4, 2)
      strSetReadout = Left$(strSetReadout, 3) & Increment(strtemp, 61) & Right$(strSetReadout, 6)
    Case 3
      strtemp = Mid$(strSetReadout, 7, 2)
      strSetReadout = Left$(strSetReadout, 6) & Increment(strtemp, 61) & Right$(strSetReadout, 3)
    Case 4
      If (Right$(strSetReadout, 2) = "AM") Then
        strSetReadout = Left$(strSetReadout, 9) & "PM"
      Else
        strSetReadout = Left$(strSetReadout, 9) & "AM"
      End If
    Case 5
      If (blnAlarmOn) Then
        picAlarmOn.Visible = False
        blnAlarmOn = False
        strSetReadout = "OFF"
      Else
        picAlarmOn.Visible = True
        blnAlarmOn = True
        strSetReadout = "ON"
      End If
    Case 6
      If Not (blnTestOn) Then
        picTestSound.Visible = True
        picTestSound.Picture = picTest.Picture
      End If
      If (strSetReadout = "Alarm_Clock") Then
        strSetReadout = "Rooster"
      ElseIf (strSetReadout = "Rooster") Then
        strSetReadout = "Siren"
      ElseIf (strSetReadout = "Siren") Then
        strSetReadout = "Explosion"
      ElseIf (strSetReadout = "Explosion") Then
        strSetReadout = "Air_Horn"
      ElseIf (strSetReadout = "Air_Horn") Then
        strSetReadout = "Warning"
      ElseIf (strSetReadout = "Warning") Then
        strSetReadout = "Alarm_Clock"
      End If
      picAlarmOn.Picture = LoadPicture(App.Path & "\" & strSetReadout & ".bmp")
      MediaPlayer2.Open (App.Path & "\" & strSetReadout & ".wav")
  End Select
End Sub

Private Sub DisplayHands(blnMode As Boolean)
  Dim I As Integer
  For I = 0 To 3
    linTimeHands(I).Visible = blnMode
  Next I
End Sub

Private Sub ClockOtherTime()
  Dim dblHandAngle As Double, Seconds As Double, Minutes As Double
  
  ' Calculate the position of the second hand
  '
  dblHandAngle = PI * (Second(strSetReadout) - 15)
  linTimeHands(2).X2 = linTimeHands(2).X1 + Cos(dblHandAngle) * intHandLength(2)
  linTimeHands(2).Y2 = linTimeHands(2).Y1 + Sin(dblHandAngle) * intHandLength(2)

  ' Calculate the position of the minute hand
  '
  Seconds = Second(strSetReadout) / 60
  dblHandAngle = PI * (Minute(strSetReadout) + Seconds - 15)
  linTimeHands(0).X2 = linTimeHands(0).X1 + Cos(dblHandAngle) * intHandLength(0)
  linTimeHands(0).Y2 = linTimeHands(0).Y1 + Sin(dblHandAngle) * intHandLength(0)
  linTimeHands(1).X2 = linTimeHands(1).X1 + Cos(dblHandAngle) * intHandLength(1)
  linTimeHands(1).Y2 = linTimeHands(1).Y1 + Sin(dblHandAngle) * intHandLength(1)

  ' Calculate the position of the hour hand
  '
  Minutes = Minute(strSetReadout) / 60
  dblHandAngle = (PI * 30) / 6 * (Hour(strSetReadout) + Minutes - 15)
  linTimeHands(3).X2 = linTimeHands(3).X1 + Cos(dblHandAngle) * intHandLength(3)
  linTimeHands(3).Y2 = linTimeHands(3).Y1 + Sin(dblHandAngle) * intHandLength(3)

End Sub

Private Sub MakeElapsedTime()
  Static intSecs As Integer
  Static intMins As Integer
  Dim strtemp As String, strCurrent As String
  strtemp = Left$(Right$(Time, 4), 2)
  If (strtemp <> strLastSecond) Then
    strCurrent = Right$(strSetReadout, 2)
    strSetReadout = Left$(strSetReadout, 6) & Increment(strCurrent, 60)
    strLastSecond = strtemp
    intSecs = intSecs + 1
    If (intSecs = 60) Then
      strCurrent = Mid$(strSetReadout, 4, 2)
      strSetReadout = Left$(strSetReadout, 3) & Increment(strCurrent, 60) & Right$(strSetReadout, 3)
      intSecs = 0
      intMins = intMins + 1
      If (intMins = 60) Then
        strCurrent = Left$(strSetReadout, 2)
        strSetReadout = Increment(strCurrent, 12) & Right$(strSetReadout, 6)
        intMins = 0
      End If
    End If
  End If
End Sub


Private Function Increment(strSet As String, Max As Integer) As String
  If ((Val(Mid$(strSet, 1, 2)) + 1) < Max) Then
    If ((Val(Mid$(strSet, 1, 2)) + 1) < 10) Then
      Increment = "0" & Val(Mid$(strSet, 2, 1) + 1)
    Else
      Increment = Val(Mid$(strSet, 1, 2) + 1)
    End If
  ElseIf ((Val(Mid$(strSet, 1, 2)) + 1) = Max) Then
    Increment = "00"
  Else
    Increment = "01"
  End If
End Function

Private Sub MakeStopTime()
  Static intNumSecs As Integer
  Static intNumMins As Integer
  Dim strtemp As String, strCurrent As String
  strtemp = Left$(Right$(Time, 4), 2)
  If (strtemp <> strLastSecond) Then
    strCurrent = Right$(strSetReadout, 2)
    If (strCurrent = "00") Then intNumSecs = 59
    strSetReadout = Left$(strSetReadout, 6) & Decrement(strCurrent, 60)
    strLastSecond = strtemp
    intNumSecs = intNumSecs + 1
    If (intNumSecs = 60) Then
      strCurrent = Mid$(strSetReadout, 4, 2)
      If (strCurrent = "00") Then intNumMins = 59
      strSetReadout = Left$(strSetReadout, 3) & Decrement(strCurrent, 60) & Right$(strSetReadout, 3)
      intNumSecs = 0
      intNumMins = intNumMins + 1
      If (intNumMins = 60) Then
        strCurrent = Left$(strSetReadout, 2)
        strSetReadout = Decrement(strCurrent, 12) & Right$(strSetReadout, 6)
        intNumMins = 0
      End If
    End If
  End If
End Sub

Private Function Decrement(strSet As String, Max As Integer) As String
  If ((Val(Mid$(strSet, 1, 2)) - 1) > 0) Then
    If ((Val(Mid$(strSet, 1, 2)) - 1) < 9) Then
      Decrement = "0" & Val(Mid$(strSet, 2, 1) - 1)
    ElseIf ((Val(Mid$(strSet, 1, 2)) - 1) = 9) Then
      Decrement = "0" & Val(Mid$(strSet, 1, 2) - 1)
    Else
      Decrement = Val(Mid$(strSet, 1, 2) - 1)
    End If
  ElseIf ((Val(Mid$(strSet, 1, 2)) - 1) = 0) Then
    Decrement = "00"
  Else
    Decrement = Right$(Str$(Max - 1), 2)
  End If
End Function

