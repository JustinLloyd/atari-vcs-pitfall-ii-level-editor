VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pitfall 2+ -- Level Editor"
   ClientHeight    =   8205
   ClientLeft      =   5895
   ClientTop       =   3810
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13395
   Begin VB.Frame frameVCSOptions 
      Caption         =   "Room Options"
      Height          =   6015
      Left            =   7800
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
      Begin VB.CheckBox checkExportCreatures 
         Caption         =   "Export Creatures"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   4800
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox checkExit 
         Caption         =   "Blocked Exit (F4)"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox comboLowNibble 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox comboHighNibble 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label labelSavePoints 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label labelFrogs 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label labelGoldBars 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label labelCondors 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label labelBats 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label labelScorpions 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Low Nibble (F3)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "High Nibble (F2)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
   End
   Begin ComCtl3.CoolBar CoolBar 
      Height          =   660
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   1164
      BandCount       =   1
      _CBWidth        =   9615
      _CBHeight       =   660
      _Version        =   "6.7.8862"
      Child1          =   "toolbarMain"
      MinWidth1       =   3195
      MinHeight1      =   600
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin MSComctlLib.Toolbar toolbarMain 
         Height          =   600
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   1058
         ButtonWidth     =   1217
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               Key             =   "New"
               Object.ToolTipText     =   "New Project"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Open"
               Key             =   "Open"
               Object.ToolTipText     =   "Open Project"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Save"
               Key             =   "Save"
               Object.ToolTipText     =   "Save Project"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cut"
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Copy"
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Paste"
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Undo"
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo the last action"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Redo"
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo the last action"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Animate"
               Key             =   "Animate"
               Object.ToolTipText     =   "Animate the map"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Play"
               Key             =   "Play"
               Object.ToolTipText     =   "Play the level"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imagelistToolbar 
      Left            =   3720
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0120
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0234
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0348
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":04AC
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":060C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":077C
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0890
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":09A4
            Key             =   "Animate"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0B04
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0C64
            Key             =   "Play"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   7905
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "StartRoomInfo"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "MapInfo"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "MemoryInfo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Redraw"
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar hscrollMap 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   4455
   End
   Begin VB.VScrollBar vscrollMap 
      Height          =   4935
      Left            =   4440
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.Timer timerAnimation 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2640
      Top             =   4680
   End
   Begin VB.PictureBox picRoom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Index           =   0
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Shape shapeHighlight 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   1695
      Left            =   720
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu menuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu menuEmpty0 
         Caption         =   "-"
      End
      Begin VB.Menu menuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu menuSaveAs 
         Caption         =   "Sa&ve As..."
      End
      Begin VB.Menu menuEmpty2 
         Caption         =   "-"
      End
      Begin VB.Menu menuExportToVCS 
         Caption         =   "Export to &Atari VCS"
         Shortcut        =   ^A
      End
      Begin VB.Menu menuMRU 
         Caption         =   "MRU"
         Index           =   0
      End
      Begin VB.Menu menuEmpty1 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu menuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu menuRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu menuEmpty4 
         Caption         =   "-"
      End
      Begin VB.Menu menuCut 
         Caption         =   "Cu&t Room"
         Shortcut        =   ^X
      End
      Begin VB.Menu menuCopy 
         Caption         =   "C&opy Room"
         Shortcut        =   ^C
      End
      Begin VB.Menu menuPaste 
         Caption         =   "&Paste Room"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu menuProject 
      Caption         =   "&Project"
      Begin VB.Menu menuProjectSetSize 
         Caption         =   "Prop&erties..."
      End
   End
   Begin VB.Menu menuRoom 
      Caption         =   "&Room"
      Begin VB.Menu menuResetRoom 
         Caption         =   "&Reset Room"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu menuTricks 
      Caption         =   "&Tricks"
      Begin VB.Menu menuLoadBitmaps 
         Caption         =   "&Load Bitmaps"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu menuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_WINDOW_TITLE = k_TITLE & " -- Level Editor"
Private Const k_KEYCODE_HIGH_NIBBLE = 113   ' keycode used to switch high nibble room codes
Private Const k_KEYCODE_LOW_NIBBLE = 114    ' keycode used to switch low nibble room codes
Private Const k_KEYCODE_EXIT = 115          ' keycode used to toggle room exit

Private Const vbmaPause = 0
Private Const vbmaPlay = 1

Private m_frameCount As Long
Private m_updatingProperties As Boolean
Private m_worldCol As Integer
Private m_worldRow As Integer
Private m_currentRoomCol As Integer
Private m_currentRoomRow As Integer
Private m_currentRoom As CRoom
Private m_currentPicture As PictureBox
Private m_copyRoom As CRoom
Private m_mruList As CMRUList
Private m_animateMap As Integer

Private Sub checkExit_Click()
    If Not m_updatingProperties Then
        Action_ModifyExitFlag checkExit.value
        DrawRoom m_currentPicture, m_currentRoom, m_frameCount, m_currentRoomCol, m_currentRoomRow
    End If
End Sub

Private Sub comboHighNibble_Click()
    If Not m_updatingProperties Then
        Action_ModifyHighNibble comboHighNibble.ListIndex
        DrawRoom m_currentPicture, m_currentRoom, m_frameCount, m_currentRoomCol, m_currentRoomRow
    End If
End Sub

Private Sub Action_ModifyHighNibble(ByVal nibble As Integer)
    m_currentRoom.HighNibble = nibble
End Sub

Private Sub comboLowNibble_Click()
    If Not m_updatingProperties Then
        Action_ModifyLowNibble comboLowNibble.ListIndex
        DrawRoom m_currentPicture, m_currentRoom, m_frameCount, m_currentRoomCol, m_currentRoomRow
    End If
End Sub

Private Sub Action_ModifyExitFlag(ByVal val As Boolean)
    m_currentRoom.ExitFlag = val
End Sub

Private Sub Action_ModifyLowNibble(ByVal nibble As Integer)
    m_currentRoom.LowNibble = nibble
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    Debug.Print KeyCode
    If KeyCode = k_KEYCODE_HIGH_NIBBLE Then
        If Shift Then
            PreviousHighNibble
        Else
            NextHighNibble
        End If
    ElseIf KeyCode = k_KEYCODE_LOW_NIBBLE Then
        If Shift Then
            PreviousLowNibble
        Else
            NextLowNibble
        End If
    ElseIf KeyCode = k_KEYCODE_EXIT Then
        ToggleExit
    End If
End Sub

Private Sub SetWindowTitle()
    Dim projName As String
    
    If g_projectFilename = "" Then
        projName = "Untitled"
    Else
        projName = g_projectFilename
    End If
    
    Me.Caption = k_WINDOW_TITLE & " (" & Trim(projName) & ")"
End Sub

Private Sub ToggleExit()
    If checkExit.value = vbChecked Then
        checkExit.value = vbUnchecked
    Else
        checkExit.value = vbChecked
    End If
End Sub

Private Sub PreviousLowNibble()
    If comboLowNibble.ListIndex - 1 >= 0 Then
        comboLowNibble.ListIndex = comboLowNibble.ListIndex - 1
    Else
        comboLowNibble.ListIndex = comboLowNibble.ListCount - 1
    End If
End Sub

Private Sub NextLowNibble()
    If comboLowNibble.ListIndex + 1 < comboLowNibble.ListCount Then
        comboLowNibble.ListIndex = comboLowNibble.ListIndex + 1
    Else
        comboLowNibble.ListIndex = 0
    End If
End Sub

Private Sub PreviousHighNibble()
    If comboHighNibble.ListIndex - 1 >= 0 Then
        comboHighNibble.ListIndex = comboHighNibble.ListIndex - 1
    Else
        comboHighNibble.ListIndex = comboHighNibble.ListCount - 1
    End If
End Sub

Private Sub NextHighNibble()
    If comboHighNibble.ListIndex + 1 < comboHighNibble.ListCount Then
        comboHighNibble.ListIndex = comboHighNibble.ListIndex + 1
    Else
        comboHighNibble.ListIndex = 0
    End If
End Sub

Private Sub Form_Load()
    Initialise
End Sub

Private Sub Initialise()
    m_frameCount = 0
    m_updatingProperties = False
    m_animateMap = vbmaPlay
    Set m_mruList = New CMRUList
    m_mruList.Init menuMRU
    InitToolbars
'    InitMenuIcons
    InitFlightPaths
    LoadBitmaps
    DisableProjectCommands
    DisablePaste
    InitOptions
    InitRooms
    CreateNewLevel
    InitFormComponents
    InitMapPosition
    Me.KeyPreview = True
End Sub

Private Sub InitFormComponents()
    Dim boundingBox As New CBoundingBox
    
    vscrollMap.Left = picRoom(k_NUM_VIEWABLE_ROOMS - 1).Left + picRoom(k_NUM_VIEWABLE_ROOMS - 1).Width + Screen.TwipsPerPixelX * 4
    vscrollMap.Top = picRoom(0).Top
    vscrollMap.Height = picRoom(k_NUM_VIEWABLE_ROOMS - 1).Top + picRoom(k_NUM_VIEWABLE_ROOMS - 1).Height - 140
    hscrollMap.Left = picRoom(0).Left
    hscrollMap.Top = picRoom(k_NUM_VIEWABLE_ROOMS - 1).Top + picRoom(k_NUM_VIEWABLE_ROOMS - 1).Height + Screen.TwipsPerPixelY * 4
    hscrollMap.Width = picRoom(k_NUM_VIEWABLE_ROOMS - 1).Left + picRoom(k_NUM_VIEWABLE_ROOMS - 1).Width - 140
    frameVCSOptions.Top = picRoom(0).Top
    frameVCSOptions.Left = vscrollMap.Left + vscrollMap.Width + 100
'    Frame2.Left = frameVCSOptions.Left
'    labelStartRoom.Left = frameGBOptions.Left
'    slideAnimSpeed.Left = frameGBOptions.Left
    CoolBar.Top = 0
    CoolBar.Left = 0
    CoolBar.Height = 600
    
'    frameVCSOptions.Left frameVCSOptions.Width, frameVCSOptions.Top + frameVCSOptions.Height + hscrollMap.Height + statusBar.Height
    
    GetControlSize Me, boundingBox
    
    InsizeForm Me, boundingBox.Width, boundingBox.Height
    CenterForm Me
    CoolBar.Width = Me.Width - 100
    CoolBar.Height = 600
End Sub

Private Sub DisableProjectCommands()
    menuSave.Enabled = False
    menuSaveAs.Enabled = False
    menuExportToVCS.Enabled = False
    toolbarMain.Buttons.Item("Save").Enabled = False
End Sub

Private Sub EnableProjectCommands()
    menuSave.Enabled = True
    menuSaveAs.Enabled = True
    menuExportToVCS.Enabled = True
    toolbarMain.Buttons.Item("Save").Enabled = True
End Sub

Private Function VerifyExitWithoutSaving() As Boolean
    Dim s As String
    
    s = "The current level has been altered but not saved." & vbCrLf
    s = s + "Are you sure you want to exit?"
    If MsgBox(s, vbYesNo, "Warning: Save Level") = vbYes Then
        VerifyExitWithoutSaving = True
    Else
        VerifyExitWithoutSaving = False
    End If
End Function

Private Function VerifyNewWithoutSaving() As Boolean
    Dim s As String
    
    s = "The current level has been altered but not saved." & vbCrLf
    s = s + "Are you sure you want to create a new level?"
    If MsgBox(s, vbYesNo, "Warning: Save Level") = vbYes Then
        VerifyNewWithoutSaving = True
    Else
        VerifyNewWithoutSaving = False
    End If
End Function

Private Function VerifyOpenWithoutSaving() As Boolean
    Dim s As String
    
    s = "The current level has been altered but not saved." & vbCrLf
    s = s + "Are you sure you want to open another level?"
    If MsgBox(s, vbYesNo, "Warning: Save Level") = vbYes Then
        VerifyOpenWithoutSaving = True
    Else
        VerifyOpenWithoutSaving = False
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    ' if current project has changed
    If g_map.HasChanged Then
        ' if user didn't really want to exit
        If Not VerifyExitWithoutSaving Then
            ' cancel exit
            Cancel = True
            ' exit function
            Exit Sub
        End If
    End If
    
    ' clear up current project
    Set g_map = Nothing
'    frmGame.PauseGame
'    DoEvents
'    Sleep 50
'    DoEvents
    Unload frmGame
End Sub

Private Sub InitScrollBars()
    vscrollMap.SmallChange = 1
    vscrollMap.LargeChange = k_VIEW_HEIGHT - 1
    vscrollMap.Min = 0
    vscrollMap.Max = g_map.Height - k_VIEW_HEIGHT
    vscrollMap.value = 0
    
    hscrollMap.SmallChange = 1
    hscrollMap.LargeChange = k_VIEW_WIDTH - 1
    hscrollMap.Min = 0
    hscrollMap.Max = g_map.Width - k_VIEW_WIDTH
    hscrollMap.value = 0
End Sub

Private Sub InitMapPosition()
    m_worldCol = 0
    m_worldRow = 0
    m_currentRoomCol = 0
    m_currentRoomRow = 0
    InitScrollBars
    Set m_currentPicture = picRoom(0)
    Set m_currentRoom = g_map.GetRoom(0, 0)
    HighlightRoom m_currentPicture
    UpdateRoomProperties m_currentRoom
    UpdateMapInfo
    DrawMapView
End Sub

Private Sub InitOptions()
    InitVCSOptions
End Sub

Private Sub InitVCSOptions()
    ' initialise low nibble options
    With comboLowNibble
        .Clear
        .AddItem k_STR_OPT_LOW_NONE, k_LOW_NONE
        .AddItem k_STR_OPT_LOW_WATER, k_LOW_WATER
        .AddItem k_STR_OPT_LOW_EARTH, k_LOW_EARTH
        .AddItem k_STR_OPT_LOW_TREETOPS_1, k_LOW_TREE_TOPS_1
        .AddItem k_STR_OPT_LOW_TREES_1, k_LOW_TREES_1
        .AddItem k_STR_OPT_LOW_FLOOR_TWO_HOLES_AND_LADDER, k_LOW_FLOOR_TWO_HOLES_AND_LADDER
        .AddItem k_STR_OPT_LOW_CORRUPT, k_LOW_CORRUPT_1
        .AddItem k_STR_OPT_LOW_CORRUPT, k_LOW_CORRUPT_2
        .AddItem k_STR_OPT_LOW_EARTH_FLAT_FLOOR, k_LOW_EARTH_FLAT_FLOOR
        .AddItem k_STR_OPT_LOW_WALKWAY, k_LOW_WALKWAY
        .AddItem k_STR_OPT_LOW_SINGLE_HOLE, k_LOW_SINGLE_HOLE
        .AddItem k_STR_OPT_LOW_SINGLE_HOLE_AND_LADDER, k_LOW_SINGLE_HOLE_AND_LADDER
        .AddItem k_STR_OPT_LOW_RIVER, k_LOW_RIVER
        .AddItem k_STR_OPT_LOW_TREE_TOPS_2, k_LOW_TREE_TOPS_2
        .AddItem k_STR_OPT_LOW_TREES_2, k_LOW_TREES_2
        .AddItem k_STR_OPT_LOW_CORRUPT, k_LOW_CORRUPT_3
    End With
    
    ' initialise high nibble options
    With comboHighNibble
        .Clear
        .AddItem k_STR_OPT_HIGH_NONE, k_HIGH_NONE
        .AddItem k_STR_OPT_HIGH_SAVE_POINT, k_HIGH_SAVE_POINT
        .AddItem k_STR_OPT_HIGH_PLATFORM_LEFT, k_HIGH_PLATFORM_LEFT
        .AddItem k_STR_OPT_HIGH_QUICKCLAW, k_HIGH_QUICKCLAW
        .AddItem k_STR_OPT_HIGH_SCORPION, k_HIGH_SCORPION
        .AddItem k_STR_OPT_HIGH_BAT, k_HIGH_BAT
        .AddItem k_STR_OPT_HIGH_CONDOR, k_HIGH_CONDOR
        .AddItem k_STR_OPT_HIGH_GOLD_BAR_LEFT, k_HIGH_GOLD_BAR_LEFT
        .AddItem k_STR_OPT_HIGH_STONE_RAT, k_HIGH_STONE_RAT
        .AddItem k_STR_OPT_HIGH_WATERFALL, k_HIGH_WATERFALL
        .AddItem k_STR_OPT_HIGH_PLATFORM_RIGHT, k_HIGH_PLATFORM_RIGHT
        .AddItem k_STR_OPT_HIGH_RHONDA, k_HIGH_RHONDA
        .AddItem k_STR_OPT_HIGH_DIAMOND_RING, k_HIGH_DIAMOND_RING
        .AddItem k_STR_OPT_HIGH_BALLOON, k_HIGH_BALLOON
        .AddItem k_STR_OPT_HIGH_FROG, k_HIGH_FROG
        .AddItem k_STR_OPT_HIGH_GOLD_BAR_RIGHT, k_HIGH_GOLD_BAR_RIGHT

    End With
    
End Sub

Private Sub InitRooms()
    Dim room As PictureBox
    Dim index As Integer
    Dim col As Integer
    Dim row As Integer
    Dim x As Long
    Dim y As Long
    Dim startX As Long
    Dim startY As Long
    
    startX = picRoom(0).Left
    startY = picRoom(0).Top
    Set m_copyRoom = Nothing
    For index = 0 To k_NUM_VIEWABLE_ROOMS - 1
        If index <> 0 Then
            Load picRoom(index)
        End If
        
        Set room = picRoom(index)
        room.ScaleMode = vbTwips
        row = index Mod k_VIEW_HEIGHT
        col = index \ k_VIEW_HEIGHT
        x = col * (k_ROOM_WIDTH_PIX + 2)
        x = startX + x * Screen.TwipsPerPixelX
        y = row * (k_ROOM_HEIGHT_PIX + 2)
        y = startY + y * Screen.TwipsPerPixelY
        room.Visible = True
        room.Left = x
        room.Top = y
        room.Height = Screen.TwipsPerPixelY * k_ROOM_HEIGHT_PIX
        room.Width = Screen.TwipsPerPixelX * k_ROOM_WIDTH_PIX
        room.ScaleMode = vbPixels
    Next
    
End Sub

Private Sub hscrollMap_Change()
    m_worldCol = hscrollMap.value
    DrawMapView
End Sub

Private Sub hscrollMap_Scroll()
    m_worldCol = hscrollMap.value
    DrawMapView
End Sub

Private Sub menuCopy_Click()
    Action_CopyRoom
End Sub

Private Sub menuCut_Click()
    Action_CutRoom
End Sub

Private Sub menuExit_Click()
    Unload Me
End Sub

Private Sub menuExportToVCS_Click()
    ExportLevel
End Sub

Private Sub menuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub menuLoadBitmaps_Click()
    LoadBitmaps
End Sub

Private Sub menuMRU_Click(index As Integer)
    VerifiedLoadLevelFromMRU index
End Sub

Private Sub menuNew_Click()
    VerifiedNewLevel
End Sub

Private Sub VerifiedNewLevel()
    ' if current project has changed
    If g_map.HasChanged Then
            ' if user didn't really want to exit
            If Not VerifyNewWithoutSaving Then
                ' exit function
                Exit Sub
            End If
    End If
    
    CreateNewLevel
End Sub
Private Sub menuOpen_Click()
    VerifiedLoadLevel
End Sub

Private Sub menuPaste_Click()
    Action_PasteRoom
End Sub

Private Sub menuProjectSetSize_Click()
    formProjectProperties.Show vbModal
    InitMapPosition
    UpdateMapInfo
End Sub

Private Sub menuResetRoom_Click()
    Action_ResetRoom
End Sub

Private Sub menuSave_Click()
    SaveLevel
End Sub

Private Sub SaveLevelToFile(ByVal filename As String)
    Dim proj As CProject
    
    Set proj = New CProject
    g_map.SerialOut proj
    If Not proj.WriteContents(filename) Then
        MsgBox "Failed to save level", , "File Error"
    Else
        g_map.ClearChange
    End If
    
End Sub

Private Function LoadLevelFromFile(ByVal filename As String) As Boolean
    Dim proj As CProject
        
    Set proj = New CProject
    If Not proj.ReadContents(filename) Then
        MsgBox "Failed to load level", , "File Error"
        LoadLevelFromFile = False
    Else
        Set g_map = New CMap
        g_map.SerialIn proj
        g_map.ClearChange
        LoadLevelFromFile = True
    End If

End Function

Private Sub menuSaveAs_Click()
    SaveLevelAs
End Sub

Private Sub picRoom_Click(index As Integer)
    Dim col As Integer
    Dim row As Integer
    
    col = index \ k_VIEW_HEIGHT
    row = index Mod k_VIEW_HEIGHT
    m_currentRoomCol = m_worldCol + col
    m_currentRoomRow = m_worldRow + row
    
    Set m_currentRoom = g_map.GetRoom(m_currentRoomCol, m_currentRoomRow)
    Set m_currentPicture = picRoom(index)
    UpdateMapInfo
    HighlightRoom m_currentPicture
    UpdateRoomProperties m_currentRoom
End Sub

Private Sub UpdateRoomNumber()
    Dim s As String
    
    s = "Room #" & Trim(m_currentRoomCol * g_map.Height + m_currentRoomRow)
    s = s & "  (" & Trim(m_currentRoomCol) & ", " & Trim(m_currentRoomRow)
    s = s & ") of " & Trim(g_map.Width * g_map.Height)
    s = s & "  (" & Trim(g_map.Width) & ", " & Trim(g_map.Height) & ")"
    statusBar.Panels.Item("MapInfo").Text = s
End Sub

Private Sub UpdateStartRoom()
    Dim s As String
    s = "Start Room #" & Trim(g_map.StartRoom) & "  (" & Trim(g_map.StartRoom \ g_map.Height) & ", " & Trim(g_map.StartRoom Mod g_map.Height) & ")"
    statusBar.Panels.Item("StartRoomInfo").Text = s
'    labelStartRoom.Caption = s
End Sub

Private Sub UpdateRoomProperties(ByRef room As CRoom)
    m_updatingProperties = True
    comboLowNibble.ListIndex = room.LowNibble
    comboHighNibble.ListIndex = room.HighNibble
    If room.ExitFlag Then
        checkExit.value = vbChecked
    Else
        checkExit.value = vbUnchecked
    End If
    
    m_updatingProperties = False
End Sub

Private Sub HighlightRoom(ByRef proom As PictureBox)
    shapeHighlight.Visible = True
    shapeHighlight.Left = proom.Left - (2 * Screen.TwipsPerPixelX)
    shapeHighlight.Width = proom.Width + (4 * Screen.TwipsPerPixelX)
    shapeHighlight.Top = proom.Top - (2 * Screen.TwipsPerPixelY)
    shapeHighlight.Height = proom.Height + (4 * Screen.TwipsPerPixelY)
End Sub


Private Sub DrawMapView()
    Dim proom As PictureBox
    Dim room As CRoom
    Dim col As Integer
    Dim row As Integer
    Dim startTick As Long
    Dim endTick As Long
    
    startTick = GetTickCount
    For col = 0 To k_VIEW_WIDTH - 1
        For row = 0 To k_VIEW_HEIGHT - 1
            Set proom = picRoom(row + col * k_VIEW_HEIGHT)
            Set room = g_map.GetRoom(m_worldCol + col, m_worldRow + row)
            DrawRoom proom, room, m_frameCount, m_worldCol + col, m_worldRow + row
        Next
    Next
    endTick = GetTickCount
    UpdateRedrawTime statusBar, "Redraw", endTick - startTick
End Sub

Private Sub timerAnimation_Timer()
    m_frameCount = m_frameCount + 1
    DrawMapView
End Sub

Private Sub vscrollMap_Change()
    m_worldRow = vscrollMap.value
    DrawMapView
End Sub

Private Function GetMapMemory() As Long
    GetMapMemory = g_map.Width * g_map.Height * k_ROOM_DATA_SIZE
End Function

Private Sub UpdateMapInfo()
    UpdateRoomNumber
    UpdateStartRoom
    UpdateMemoryStats
    UpdateRoomStats
End Sub

Private Sub UpdateRoomStats()
    Dim col As Integer
    Dim row As Integer
    Dim room As CRoom
    Dim batCount As Integer
    Dim condorCount As Integer
    Dim scorpionCount As Integer
    Dim goldBarCount As Integer
    Dim frogCount As Integer
    Dim savePointCount As Integer
    
    batCount = 0
    condorCount = 0
    scorpionCount = 0
    goldBarCount = 0
    frogCount = 0
    savePointCount = 0
    For col = 0 To g_map.Width - 1
        For row = 0 To g_map.Height - 1
            Set room = g_map.GetRoom(col, row)
            With room
                Select Case .HighNibble
                    Case k_HIGH_BAT
                        batCount = batCount + 1
                    Case k_HIGH_SCORPION
                        scorpionCount = scorpionCount + 1
                    Case k_HIGH_CONDOR
                        condorCount = condorCount + 1
                    Case k_HIGH_FROG
                        frogCount = frogCount + 1
                    Case k_HIGH_GOLD_BAR_LEFT
                        goldBarCount = goldBarCount + 1
                    Case k_HIGH_GOLD_BAR_RIGHT
                        goldBarCount = goldBarCount + 1
                    Case k_HIGH_SAVE_POINT
                        savePointCount = savePointCount + 1
                End Select
            End With
        Next
    Next
    
    labelBats.Caption = "Bats: " & Trim(batCount)
    labelCondors.Caption = "Condors: " & Trim(condorCount)
    labelScorpions.Caption = "Scorpions: " & Trim(scorpionCount)
    labelFrogs.Caption = "Frogs: " & Trim(frogCount)
    labelSavePoints.Caption = "Save Points: " & Trim(savePointCount)
    labelGoldBars.Caption = "Gold Bars: " & Trim(goldBarCount)
End Sub

Private Sub UpdateMemoryStats()
    Dim s As String
    
    s = "  Memory: " & Trim(GetMapMemory / 1024) & "KB"
    statusBar.Panels.Item("MemoryInfo").Text = s
End Sub

Private Function InputBinaryFilenameToExport(ByRef filename As String) As Boolean
    filename = InputFilename(".BIN", "Pitfall 2 Export (*.BIN)|*.bin|All (*.*)|*.*", "Export Pitfall 2 Level As", vbffSaveDialog)
    If filename = "" Then
        InputBinaryFilenameToExport = False
    Else
        InputBinaryFilenameToExport = True
    End If

End Function

Private Function InputProjectFilenameToSave(ByRef filename As String) As Boolean
    filename = InputFilename(".PF2", "Pitfall 2 Level (*.PF2)|*.pf2|All (*.*)|*.*", "Save Pitfall 2 Level As", vbffSaveDialog)
    If filename = "" Then
        InputProjectFilenameToSave = False
    Else
        InputProjectFilenameToSave = True
    End If
    
End Function

Private Function InputProjectFilenameToOpen(ByRef selectedFilename As String) As Boolean
    selectedFilename = InputFilename(".PF2", "Pitfall 2 Level (*.PF2)|*.pf2|All (*.*)|*.*", "Open Pitfall 2 Level File", vbffLoadDialog)
    If selectedFilename = "" Then
        InputProjectFilenameToOpen = False
    Else
        InputProjectFilenameToOpen = True
    End If
    
End Function

Private Sub CreateNewLevel()
    g_projectFilename = ""
    ' create a new level map
    Set g_map = Nothing
    Set g_map = New CMap
    ' enable save project
    EnableProjectCommands
    InitMapPosition
    UpdateMapInfo
    SetWindowTitle
End Sub

Private Sub LoadLevelFromMRU(ByVal index As Integer)
    Dim ret As Boolean
    Dim filename As String
    
    ' get file from mru list
    g_projectFilename = m_mruList.Item(index)
    ' load level from supplied file
    If LoadLevelFromFile(g_projectFilename) Then
        InitMapPosition
        UpdateMapInfo
        SetWindowTitle
        m_mruList.AddItem g_projectFilename
    End If
    
End Sub

Private Sub LoadLevel()
    Dim ret As Boolean
    Dim filename As String
    
    ' get file from user
    ret = InputProjectFilenameToOpen(filename)
    If Not ret Then
        Exit Sub
    End If
    
    g_projectFilename = filename
    ' load level from supplied file
    If LoadLevelFromFile(g_projectFilename) Then
        InitMapPosition
        UpdateMapInfo
        SetWindowTitle
        m_mruList.AddItem g_projectFilename
    End If
    
End Sub

Private Sub ExportLevel()
    Dim ret As Boolean
    Dim filename As String
    
    ret = InputBinaryFilenameToExport(filename)
    
    If Not ret Then
        Exit Sub
    End If

    If checkExportCreatures.value = vbChecked Then
        ExportLevelToVCS filename, True
    Else
        ExportLevelToVCS filename, False
    End If
    
End Sub

Private Sub SaveLevel()
    Dim ret As Boolean
    Dim filename As String
    
    If g_projectFilename = "" Then
        ret = InputProjectFilenameToSave(filename)
        If Not ret Then
            Exit Sub
        End If
        
        g_projectFilename = filename
    End If
    
    SaveLevelToFile g_projectFilename
    SetWindowTitle
    m_mruList.AddItem g_projectFilename
End Sub

Private Sub SaveLevelAs()
    Dim ret As Boolean
    Dim filename As String
    
    ret = InputProjectFilenameToSave(filename)
    If Not ret Then
        Exit Sub
    End If
        
    g_projectFilename = filename
    SaveLevel
End Sub

Private Sub ExportLevelToVCS(ByVal baseFilename As String, ByVal exportCreatures As Boolean)
    ExportBinaryFile baseFilename, exportCreatures
End Sub

Private Sub ExportBinaryFile(ByVal baseFilename As String, ByVal exportCreatures As Boolean)
    Dim F As Integer
    Dim label As String
    Dim headerFilename As String
    Dim exportData As CVCSImage
    
    If g_map.Width <> 8 Then
        MsgBox "The map must be 8 rooms wide"
        Exit Sub
    End If
    
    If g_map.Height <> 32 Then
        MsgBox "The map must be 32 rooms high"
        Exit Sub
    End If
    
    On Local Error GoTo ErrorHandler
    ' create atari binary
    Set exportData = New CVCSImage
    g_map.ExportBinary exportData, exportCreatures
    
    ' write source file
    F = FreeFile
    Open baseFilename For Binary Access Write As F
    Put #F, , exportData.GetData
    Close F
    Exit Sub
ErrorHandler:
    MsgBox "Failed to export source file"
End Sub

Private Sub Action_PlayGame()
    frmGame.Show
End Sub

Private Sub Action_CutRoom()
    Action_CopyRoom
    Action_ResetRoom
End Sub

Private Sub Action_CopyRoom()
    Set m_copyRoom = New CRoom
    m_copyRoom.Copy m_currentRoom
    EnablePaste
End Sub

Private Sub Action_PasteRoom()
    m_currentRoom.Copy m_copyRoom
    UpdateRoomProperties m_currentRoom
    DrawRoom m_currentPicture, m_currentRoom, m_frameCount, m_currentRoomCol, m_currentRoomRow
End Sub

Private Sub DisablePaste()
    menuPaste.Enabled = False
    toolbarMain.Buttons.Item("Paste").Enabled = False
End Sub

Private Sub EnablePaste()
    menuPaste.Enabled = True
    toolbarMain.Buttons.Item("Paste").Enabled = True
End Sub

Private Sub Action_ResetRoom()
    m_currentRoom.Reset
    UpdateRoomProperties m_currentRoom
    DrawRoom m_currentPicture, m_currentRoom, m_frameCount, m_currentRoomCol, m_currentRoomRow
End Sub

Private Sub Action_AnimateMap()
    Debug.Assert m_animateMap = vbmaPause Or m_animateMap = vbmaPlay
    If m_animateMap = vbmaPlay Then
        m_animateMap = vbmaPause
        toolbarMain.Buttons.Item("Animate").Image = "Pause"
        timerAnimation.Enabled = True
    Else
        m_animateMap = vbmaPlay
        toolbarMain.Buttons.Item("Animate").Image = "Animate"
        timerAnimation.Enabled = False
    End If
        
End Sub

Private Sub vscrollMap_Scroll()
    m_worldRow = vscrollMap.value
    DrawMapView
End Sub

Private Sub InitToolbars()
    CoolBar.Height = 600
    InitToolbarMain
End Sub

Private Sub InitToolbarMain()
    Dim buttonItem As Button
    
    toolbarMain.ImageList = imagelistToolbar
    For Each buttonItem In toolbarMain.Buttons
        If buttonItem.Style <> tbrSeparator Then
            buttonItem.Image = buttonItem.Key
            buttonItem.Caption = ""
        End If
    Next
    
    DisableToolbar
End Sub

Private Sub EnableToolbar()
    Dim buttonItem As Button
    
    ' enable all buttons
    For Each buttonItem In toolbarMain.Buttons
        buttonItem.Enabled = True
    Next

End Sub

Private Sub DisableToolbar()
    Dim buttonItem As Button
    
    ' disable all buttons
    For Each buttonItem In toolbarMain.Buttons
        buttonItem.Enabled = False
    Next
    
    ' enable those buttons that are always enabled
    toolbarMain.Buttons.Item("New").Enabled = True
    toolbarMain.Buttons.Item("Open").Enabled = True
    toolbarMain.Buttons.Item("Copy").Enabled = True
    toolbarMain.Buttons.Item("Cut").Enabled = True
    toolbarMain.Buttons.Item("Animate").Enabled = True
    toolbarMain.Buttons.Item("Play").Enabled = True
'    toolbarMain.Buttons.Item("Undo").Enabled = True
'    toolbarMain.Buttons.Item("Redo").Enabled = True
'    toolbarMain.Buttons.Item("Toolbox").Enabled = True
'    toolbarMain.Buttons.Item("Spawnpoints").Enabled = True
End Sub

Private Sub toolbarMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "New" Then
        VerifiedNewLevel
    ElseIf Button.Key = "Save" Then
        SaveLevel
    ElseIf Button.Key = "Open" Then
        VerifiedLoadLevel
    ElseIf Button.Key = "Cut" Then
        Action_CutRoom
    ElseIf Button.Key = "Copy" Then
        Action_CopyRoom
    ElseIf Button.Key = "Paste" Then
        Action_PasteRoom
    ElseIf Button.Key = "Undo" Then
        Debug.Assert False
    ElseIf Button.Key = "Redo" Then
        Debug.Assert False
    ElseIf Button.Key = "Animate" Then
        Action_AnimateMap
    ElseIf Button.Key = "Play" Then
        Action_PlayGame
    End If
   
End Sub

Private Sub VerifiedLoadLevel()
    ' if current project has changed
    If g_map.HasChanged Then
            ' if user didn't really want to exit
            If Not VerifyOpenWithoutSaving Then
                ' cancel exit
                ' exit function
                Exit Sub
            End If
    End If
    
    DisablePaste
    LoadLevel
End Sub

Private Sub VerifiedLoadLevelFromMRU(ByVal index As Integer)
    ' if current project has changed
    If g_map.HasChanged Then
            ' if user didn't really want to exit
            If Not VerifyOpenWithoutSaving Then
                ' cancel exit
                ' exit function
                Exit Sub
            End If
    End If
    
    DisablePaste
    LoadLevelFromMRU index
End Sub

