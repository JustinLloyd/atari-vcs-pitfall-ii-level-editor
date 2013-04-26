VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRoom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2145
      ScaleWidth      =   2025
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
   End
   Begin VB.PictureBox picBackbuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   4920
      ScaleHeight     =   1995
      ScaleWidth      =   2085
      TabIndex        =   1
      Top             =   120
      Width           =   2115
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2145
      ScaleWidth      =   2025
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   9465
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "MapInfo"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Redraw"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const k_SCREEN_WIDTH = 160
Private Const k_SCREEN_HEIGHT = 144
Private Const k_SCREEN_SCALE = 2

Private Const k_MSG_GAME_PAUSED = "Game Paused"
Private Const k_MSG_GAME_RUNNING = "Game In Progress"

Private Const vbPlayerKeyLeft = vbKeyLeft
Private Const vbPlayerKeyRight = vbKeyRight
Private Const vbPlayerKeyUp = vbKeyUp
Private Const vbPlayerKeyDown = vbKeyDown
Private Const vbPlayerKeyJump = vbKeyA

Private Const k_KEY_UP = 1
Private Const k_KEY_DOWN = 2
Private Const k_KEY_LEFT = 4
Private Const k_KEY_RIGHT = 8
Private Const k_KEY_JUMP = 16


Private m_drawTime As Long
Private m_isPlaying As Boolean
Private m_frameCount As Long

Private Sub Form_Load()
    Initialise
End Sub

Private Sub InitRoomBuffer()
    picRoom.ScaleMode = vbTwips
'    picRoom.Visible = False
    picRoom.Height = Screen.TwipsPerPixelY * k_ROOM_HEIGHT_PIX
    picRoom.Width = Screen.TwipsPerPixelX * k_ROOM_WIDTH_PIX
    picRoom.ScaleMode = vbPixels
'    picRoom.BorderStyle = vbBSNone
End Sub

Private Sub Initialise()
    InitRoomBuffer
    ' set the on-screen buffer to the correct size
    picScreen.Move 0, 0, LCDScreenWidth(), LCDScreenHeight()
    picScreen.BorderStyle = vbBSNone
    ' set up the back buffer
'    picBackbuffer.Visible = False
    picBackbuffer.Width = LCDScreenWidth()
    ' make the back buffer tall enough to have 4 rooms drawn to it
    picBackbuffer.Height = Screen.TwipsPerPixelY * k_SCREEN_SCALE * k_ROOM_HEIGHT_PIX * 4
'    picBackbuffer.BorderStyle = vbBSNone
    
    m_isPlaying = False
    m_frameCount = 0
    UpdateRedrawTime statusBar, "Redraw", 0
'    InsizeForm Me, picScreen.Width, picScreen.Height + statusBar.Height
    CenterForm Me
End Sub

Private Function LCDScreenWidth()
    LCDScreenWidth = k_SCREEN_WIDTH * k_SCREEN_SCALE * Screen.TwipsPerPixelX
End Function

Private Function LCDScreenHeight()
    LCDScreenHeight = k_SCREEN_HEIGHT * k_SCREEN_SCALE * Screen.TwipsPerPixelY
End Function

Private Sub GetPlayerInput()
    If GetAsyncKeyState(vbPlayerKeyLeft) Then
        m_keyPress = m_keyPress Or k_KEY_LEFT
    End If
    If GetAsyncKeyState(vbPlayerKeyRight) Then
        m_keyPress = m_keyPress Or k_KEY_RIGHT
    End If
    If GetAsyncKeyState(vbPlayerKeyUp) Then
        m_keyPress = m_keyPress Or k_KEY_UP
    End If
    If GetAsyncKeyState(vbPlayerKeyDown) Then
        m_keyPress = m_keyPress Or k_KEY_DOWN
    End If
    If GetAsyncKeyState(vbPlayerKeyJump) Then
        m_keyPress = m_keyPress Or k_KEY_JUMP
    End If
End Sub

Private Sub DrawScreen()
    DrawMapView
    DrawScore
    DrawDeveloperLogo
End Sub

Private Sub DrawMapView()
    Dim proom As PictureBox
    Dim room As CRoom
    Dim col As Integer
    Dim row As Integer
    Dim startTick As Long
    Dim endTick As Long
    Dim hdcBackBuffer As Long
    
    startTick = GetTickCount
'    For col = 0 To k_VIEW_WIDTH - 1
'        For row = 0 To k_VIEW_HEIGHT - 1
'            Set proom = picRoom(row + col * k_VIEW_HEIGHT)
            Set room = g_map.GetRoom(0, 0)
            ' to do
            ' blit the room to a picture box
            ' blit the picturebox to the screen picturebox
            ' repeat for four rooms
            DrawRoom picRoom, room, m_frameCount, 0, 0
'            picRoom.Refresh
'            picRoom.Picture = picRoom.Image
'            hdcBackBuffer = CreateCompatibleDC(picRoom.hdc)
'            Debug.Assert hdcBackBuffer <> 0
            ret = BitBlt(picBackbuffer.hdc, 0, 0, k_SCREEN_WIDTH, k_SCREEN_HEIGHT, picRoom.hdc, 0&, 0&, SRCCOPY)
'            picBackbuffer.Refresh
'            picBackbuffer.Refresh
'            picBackbuffer.Picture = picBackbuffer.Image
'            picScreen.Picture = picBackbuffer.Picture
            ret = BitBlt(picScreen.hdc, 0, 0, k_SCREEN_WIDTH, k_SCREEN_HEIGHT, picBackbuffer.hdc, 0&, 0&, SRCCOPY)
            picScreen.Refresh
'            picScreen.Picture = picScreen.Image
'        Next
'    Next
    endTick = GetTickCount
    m_drawTime = endTick - startTick
    UpdateRedrawTime statusBar, "Redraw", m_drawTime

End Sub

Private Sub DrawScore()

End Sub

Private Sub DrawDeveloperLogo()

End Sub

Private Sub UpdateCreatures()

End Sub

Private Sub UpdatePlayer()

End Sub

Private Sub PlayGame()
    Do
        m_frameCount = m_frameCount + 1
        DrawScreen
        GetPlayerInput
        UpdateCreatures
        UpdatePlayer
        ' This is here to stop the animation
        ' getting too fast to see:
        Sleep 33 - m_drawTime
        ' Ensure we can still click buttons etc
        DoEvents
    Loop While m_isPlaying
End Sub

Private Sub SetWindowTitle(ByVal title As String)
        Me.Caption = title
End Sub

Private Sub Action_TogglePlay()
    If m_isPlaying Then
        m_isPlaying = False
        SetWindowTitle k_MSG_GAME_PAUSED
    Else
        m_isPlaying = True
        SetWindowTitle k_MSG_GAME_RUNNING
        PlayGame
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_isPlaying = False
    DoEvents
    Sleep 50
    DoEvents
End Sub

Private Sub picScreen_Click()
    Action_TogglePlay
End Sub

Public Sub PauseGame()
    m_isPlaying = False
    SetWindowTitle k_MSG_GAME_PAUSED
End Sub
