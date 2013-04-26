Attribute VB_Name = "General"
Option Explicit

Public Const k_TITLE = "Pitfall 2+"
Public Const k_PLATFORM = "Atari VCS 2600"

Public Const k_FLIGHT_PATH_LEN = 32
Public Const k_MAX_ANIM_FRAMES = 16

Public Const k_MIN_MAP_WIDTH = 8
Public Const k_MAX_MAP_WIDTH = 8
Public Const k_MIN_MAP_HEIGHT = 32
Public Const k_MAX_MAP_HEIGHT = 32

Public Const k_ROOM_DATA_SIZE = 1
Public Const k_ROOM_HEIGHT_PIX = 48&
Public Const k_ROOM_WIDTH_PIX = 160&

Public Const k_FLOOR_X = 0&
Public Const k_FLOOR_Y = 0&

Public Const k_BACKGROUND_X = 0&
Public Const k_BACKGROUND_Y = 0&

Public Const k_FEATURE_LADDER_X = 76&
Public Const k_FEATURE_LADDER_Y = 0&
Public Const k_FEATURE_SAVE_POINT_X = 40&
Public Const k_FEATURE_SAVE_POINT_Y = 40&
Public Const k_FEATURE_WATERFALL_X = 0&
Public Const k_FEATURE_WATERFALL_Y = 0&


Public Const k_ITEM_STONE_RAT_X = 40&
Public Const k_ITEM_STONE_RAT_Y = 40&
Public Const k_ITEM_QUICKCLAW_CAT_X = 40&
Public Const k_ITEM_QUICKCLAW_CAT_Y = 24&
Public Const k_ITEM_DIAMOND_RING_X = 40&
Public Const k_ITEM_DIAMOND_RING_Y = 40&
Public Const k_ITEM_RHONDA_GIRL_X = 40&
Public Const k_ITEM_RHONDA_GIRL_Y = 0&
Public Const k_ITEM_GOLD_BAR_LEFT_X = 20&
Public Const k_ITEM_GOLD_BAR_LEFT_Y = 32&
Public Const k_ITEM_GOLD_BAR_RIGHT_X = 120&
Public Const k_ITEM_GOLD_BAR_RIGHT_Y = 32&

Public Const k_CREATURE_BAT_Y = 24&
Public Const k_CREATURE_SCORPION_Y = 32&
Public Const k_CREATURE_CONDOR_Y = 24&
Public Const k_CREATURE_EEL_Y = 16&
Public Const k_CREATURE_FROG_Y = 40&

Public Const k_EXIT_LEFT_X = 0&
Public Const k_EXIT_LEFT_Y = 0&

Public Const k_EXIT_RIGHT_X = 144&
Public Const k_EXIT_RIGHT_Y = 0&

Public Const k_VIEW_WIDTH = 4
Public Const k_VIEW_HEIGHT = 9
Public Const k_NUM_VIEWABLE_ROOMS = k_VIEW_WIDTH * k_VIEW_HEIGHT


Public Type Animation
    m_frameCount As Integer
    m_width As Long
    m_height As Long
    m_frame() As Picture
    m_mask() As Picture
End Type

Public g_picBackground(0 To k_BACKGROUND_LAST - 1) As Animation
Public g_picFloor(0 To k_FLOOR_LAST - 1) As Animation
Public g_picFeature(0 To k_FEATURE_LAST - 1) As Animation
Public g_picCreature(0 To k_CREATURE_LAST - 1) As Animation
Public g_picItem(0 To k_ITEM_LAST - 1) As Animation
Public g_picExitLeft(0 To k_EXIT_LEFT_LAST - 1) As Animation
Public g_picExitRight(0 To k_EXIT_RIGHT_LAST - 1) As Animation
Public g_flightPath(0 To k_FLIGHT_PATH_LEN - 1) As Integer
Public g_projectFilename As String
Public g_map As CMap

Public Sub LoadBitmaps()
    LoadBackgroundBitmaps
    LoadFloorBitmaps
    LoadFeatureBitmaps
    LoadCreatureBitmaps
    LoadItemBitmaps
    LoadExitLeftBitmaps
    LoadExitRightBitmaps
End Sub

Private Sub LoadComponentBitmaps(ByRef component As Animation, ByVal baseName As String)
    Dim bitmapFilepath As String
    Dim bitmapDir As String
    Dim index As Integer
    
    Debug.Assert Len(baseName) > 0
    component.m_frameCount = 0
    bitmapDir = App.path & "\Bitmaps\"
    bitmapFilepath = bitmapDir & baseName & ".bmp"
    If DoesFileExist(bitmapFilepath) Then
        bitmapFilepath = bitmapDir & baseName & ".bmp"
        component.m_frameCount = 1
        LoadBitmap bitmapFilepath, component, 0
    Else
        For index = 0 To 15
            bitmapFilepath = bitmapDir & baseName & " Frame " & Trim(index) & ".bmp"
            If DoesFileExist(bitmapFilepath) Then
                component.m_frameCount = index + 1
                LoadBitmap bitmapFilepath, component, index
            Else
                Exit For
            End If
        Next
        
    End If
    
    Debug.Assert component.m_frameCount > 0
    component.m_width = formMain.ScaleX(component.m_frame(0).Width, vbHimetric, vbPixels)
    component.m_height = formMain.ScaleY(component.m_frame(0).Height, vbHimetric, vbPixels)
End Sub

Private Sub LoadBitmap(ByVal bitmapFilepath As String, ByRef component As Animation, ByVal index As Integer)
    ReDim Preserve component.m_frame(0 To index)
    Set component.m_frame(index) = LoadPicture(bitmapFilepath)
    ReDim Preserve component.m_mask(0 To index)
    Set component.m_mask(index) = CreateMask(formMain, component.m_frame(index))
End Sub

Public Sub LoadFloorBitmaps()
    LoadComponentBitmaps g_picFloor(k_FLOOR_PLAT_LEFT), "Floor 1 Platform Left"
    LoadComponentBitmaps g_picFloor(k_FLOOR_PLAT_RIGHT), "Floor 2 Platform Right"
    LoadComponentBitmaps g_picFloor(k_FLOOR_PLAT_BOTH), "Floor 3 Platform Both"
    LoadComponentBitmaps g_picFloor(k_FLOOR_PLAT_LEFT_WATER), "Floor 4 Platform Left Water"
    LoadComponentBitmaps g_picFloor(k_FLOOR_PLAT_RIGHT_WATER), "Floor 5 Platform Right Water"
    LoadComponentBitmaps g_picFloor(k_FLOOR_PLAT_BOTH_WATER), "Floor 6 Platform Both Water"
    LoadComponentBitmaps g_picFloor(k_FLOOR_WATER), "Floor 7 Water"
    LoadComponentBitmaps g_picFloor(k_FLOOR_SOLID), "Floor 8 Solid"
    LoadComponentBitmaps g_picFloor(k_FLOOR_WALKWAY), "Floor 9 Walkway"
    LoadComponentBitmaps g_picFloor(k_FLOOR_WALKWAY_SINGLE_HOLE), "Floor 10 Walkway Single Hole"
    LoadComponentBitmaps g_picFloor(k_FLOOR_WALKWAY_THREE_HOLES), "Floor 11 Walkway Three Holes"
    LoadComponentBitmaps g_picFloor(k_FLOOR_RIVER_BED), "Floor 12 River Bed"
    LoadComponentBitmaps g_picFloor(k_FLOOR_WALKWAY_HOLE_WITH_LADDER), "Floor 13 Walkway Hole With Ladder"
    LoadComponentBitmaps g_picFloor(k_FLOOR_WALKWAY_SIX_HOLES), "Floor 14 Walkway Six Holes"
End Sub

Public Sub LoadFeatureBitmaps()
    LoadComponentBitmaps g_picFeature(k_FEATURE_SAVE_POINT), "Feature 1 Save Point"
    LoadComponentBitmaps g_picFeature(k_FEATURE_LADDER), "Feature 2 Ladder"
    LoadComponentBitmaps g_picFeature(k_FEATURE_WATERFALL), "Feature 4 Waterfall"
End Sub

Public Sub LoadExitLeftBitmaps()
    LoadComponentBitmaps g_picExitLeft(k_EXIT_LEFT_DARK_ROCK_BLUE), "Exit Left 1 Dark Rock Blue"
    LoadComponentBitmaps g_picExitLeft(k_EXIT_LEFT_LIGHT_ROCK_BLACK), "Exit Left 2 Light Rock Black"
    LoadComponentBitmaps g_picExitLeft(k_EXIT_LEFT_DARK_ROCK_BLACK), "Exit Left 3 Dark Rock Black"
    LoadComponentBitmaps g_picExitLeft(k_EXIT_LEFT_LIGHT_ROCK_GREEN), "Exit Left 4 Light Rock Green"
    LoadComponentBitmaps g_picExitLeft(k_EXIT_LEFT_DARK_ROCK_GREEN), "Exit Left 5 Dark Rock Green"
    LoadComponentBitmaps g_picExitLeft(k_EXIT_LEFT_PATTERN_ROCK_BLACK), "Exit Left 6 Pattern Rock Black"
End Sub

Public Sub LoadExitRightBitmaps()
    LoadComponentBitmaps g_picExitRight(k_EXIT_RIGHT_DARK_ROCK_BLUE), "Exit Right 1 Dark Rock Blue"
    LoadComponentBitmaps g_picExitRight(k_EXIT_RIGHT_LIGHT_ROCK_BLACK), "Exit Right 2 Light Rock Black"
    LoadComponentBitmaps g_picExitRight(k_EXIT_RIGHT_DARK_ROCK_BLACK), "Exit Right 3 Dark Rock Black"
    LoadComponentBitmaps g_picExitRight(k_EXIT_RIGHT_LIGHT_ROCK_GREEN), "Exit Right 4 Light Rock Green"
    LoadComponentBitmaps g_picExitRight(k_EXIT_RIGHT_DARK_ROCK_GREEN), "Exit Right 5 Dark Rock Green"
    LoadComponentBitmaps g_picExitRight(k_EXIT_RIGHT_PATTERN_ROCK_BLACK), "Exit Right 6 Pattern Rock Black"
End Sub

Public Sub LoadCreatureBitmaps()
    LoadComponentBitmaps g_picCreature(k_CREATURE_BAT), "Creature 1 Bat"
    LoadComponentBitmaps g_picCreature(k_CREATURE_CONDOR), "Creature 2 Condor"
    LoadComponentBitmaps g_picCreature(k_CREATURE_EEL), "Creature 3 Electric Eel"
    LoadComponentBitmaps g_picCreature(k_CREATURE_FROG), "Creature 4 Frog"
    LoadComponentBitmaps g_picCreature(k_CREATURE_SCORPION), "Creature 5 Scorpion"
End Sub

Public Sub LoadItemBitmaps()
    LoadComponentBitmaps g_picItem(k_ITEM_STONE_RAT), "Item 1 Stone Rat"
    LoadComponentBitmaps g_picItem(k_ITEM_QUICKCLAW_CAT), "Item 2 Quickclaw Cat"
    LoadComponentBitmaps g_picItem(k_ITEM_DIAMOND_RING), "Item 3 Diamond"
    LoadComponentBitmaps g_picItem(k_ITEM_RHONDA_GIRL), "Item 4 Rhonda"
    LoadComponentBitmaps g_picItem(k_ITEM_GOLD_BAR_LEFT), "Item 5 Gold Bar"
    LoadComponentBitmaps g_picItem(k_ITEM_GOLD_BAR_RIGHT), "Item 5 Gold Bar"
End Sub

Public Sub LoadBackgroundBitmaps()
    LoadComponentBitmaps g_picBackground(k_BACKGROUND_NONE), "Background 0 None"
    LoadComponentBitmaps g_picBackground(k_BACKGROUND_TREES), "Background 1 Trees"
    LoadComponentBitmaps g_picBackground(k_BACKGROUND_TREE_TOPS), "Background 2 Tree Tops"
    LoadComponentBitmaps g_picBackground(k_BACKGROUND_WATER), "Background 3 Water"
    LoadComponentBitmaps g_picBackground(k_BACKGROUND_EARTH), "Background 4 Earth"
End Sub

Public Sub InitFlightPaths()
    g_flightPath(0) = 0
    g_flightPath(1) = 0
    g_flightPath(2) = 1
    g_flightPath(3) = 1
    g_flightPath(4) = 2
    g_flightPath(5) = 2
    g_flightPath(6) = 3
    g_flightPath(7) = 3
    g_flightPath(8) = 4
    g_flightPath(9) = 4
    g_flightPath(10) = 5
    g_flightPath(11) = 5
    g_flightPath(12) = 6
    g_flightPath(13) = 6
    g_flightPath(14) = 7
    g_flightPath(15) = 7
    g_flightPath(16) = 7
    g_flightPath(17) = 7
    g_flightPath(18) = 6
    g_flightPath(19) = 6
    g_flightPath(20) = 5
    g_flightPath(21) = 5
    g_flightPath(22) = 4
    g_flightPath(23) = 4
    g_flightPath(24) = 3
    g_flightPath(25) = 3
    g_flightPath(26) = 2
    g_flightPath(27) = 2
    g_flightPath(28) = 1
    g_flightPath(29) = 1
    g_flightPath(30) = 0
    g_flightPath(31) = 0
End Sub


Public Sub UpdateRedrawTime(ByRef status As statusBar, ByVal panelID As String, ByVal redrawTime As Long)
    Dim s As String
    
    s = "  Redraw: " & redrawTime & " ms"
    status.Panels.Item(panelID).Text = s
End Sub

Public Sub Main()
    frmSplash.Show vbModal
    formMain.Show
End Sub

Public Sub DrawRoom(ByRef proom As PictureBox, ByRef room As CRoom, ByVal frameCount As Long, ByVal col As Integer, ByVal row As Integer)
    DrawRoomBackground proom.hdc, room, frameCount
    DrawRoomFloor proom.hdc, room, frameCount
    DrawRoomExit proom.hdc, g_map.GetRoom(col, row), frameCount
    DrawRoomItem proom.hdc, room, frameCount
    DrawRoomFeature proom.hdc, room, frameCount
    DrawRoomCreature proom.hdc, room, frameCount
    DrawPath proom, col, row
    proom.Refresh
    
End Sub

Private Sub DrawPath(ByRef proom As PictureBox, ByVal col As Integer, ByVal row As Integer)
    Dim roomAbove As CRoom
    Dim roomBelow As CRoom
    Dim roomToLeft As CRoom
    Dim roomToRight As CRoom
    Dim thisRoom As CRoom
    
    Dim exitAbove As Boolean
    Dim exitBelow As Boolean
    Dim exitToLeft As Boolean
    Dim exitToRight As Boolean
    
    exitAbove = False
    exitBelow = False
    exitToLeft = False
    exitToRight = False
    ' this room
    Set thisRoom = g_map.GetRoom(col, row)
    
    ' room above
    If row > 0 Then
        Set roomAbove = g_map.GetRoom(col, row - 1)
    Else
        Set roomAbove = g_map.GetRoom(col, k_MAX_MAP_HEIGHT - 1)
    End If
    
    ' room below
    If row < k_MAX_MAP_HEIGHT - 1 Then
        Set roomBelow = g_map.GetRoom(col, row + 1)
    Else
        Set roomBelow = g_map.GetRoom(col, 0)
    End If
    
    ' room to left
    If col > 0 Then
        Set roomToLeft = g_map.GetRoom(col - 1, row)
    Else
        Set roomToLeft = g_map.GetRoom(k_MAX_MAP_WIDTH - 1, row)
    End If
    
    ' room to right
    If col < k_MAX_MAP_WIDTH - 1 Then
        Set roomToRight = g_map.GetRoom(col + 1, row)
    Else
        Set roomToRight = g_map.GetRoom(0, row)
    End If
    
    ' analyse exit above
    Select Case thisRoom.LowNibble
        Case k_LOW_NONE
            exitAbove = True
        Case k_LOW_WATER
            exitAbove = True
        Case k_LOW_RIVER
            exitAbove = True
        Case k_LOW_FLOOR_TWO_HOLES_AND_LADDER
            exitAbove = True
        Case k_LOW_SINGLE_HOLE
            exitAbove = True
        Case k_LOW_SINGLE_HOLE_AND_LADDER
            exitAbove = True
    End Select
    
    Select Case roomAbove.LowNibble
        Case k_LOW_EARTH
            exitAbove = False
        Case k_LOW_EARTH_FLAT_FLOOR
            exitAbove = False
    End Select
    
    ' analyse exit below
    Select Case roomBelow.LowNibble
        Case k_LOW_NONE
            exitBelow = True
        Case k_LOW_FLOOR_TWO_HOLES_AND_LADDER
            exitBelow = True
        Case k_LOW_RIVER
            exitBelow = True
'        Case k_LOW_SINGLE_HOLE
'            exitBelow = True
        Case k_LOW_SINGLE_HOLE_AND_LADDER
            exitBelow = True
        Case k_LOW_WATER
            exitBelow = True
    End Select
    
    Select Case thisRoom.LowNibble
        Case k_LOW_EARTH
            exitBelow = False
        Case k_LOW_EARTH_FLAT_FLOOR
            exitBelow = False
    End Select
    
    ' analyse exit to right
    If thisRoom.ExitFlag = False Then
        exitToRight = True
    End If
    
    If thisRoom.LowNibble = k_LOW_EARTH Then
        exitToRight = False
    ElseIf thisRoom.LowNibble = k_LOW_EARTH_FLAT_FLOOR Then
        exitToRight = False
    End If
    
    If roomBelow.LowNibble = k_LOW_RIVER Or roomBelow.LowNibble = k_LOW_NONE Then
        exitToRight = False
    End If
    
    ' analyse exit to left
    If roomToLeft.ExitFlag = False Then
        exitToLeft = True
    End If
    
    Select Case thisRoom.LowNibble
        Case k_LOW_EARTH
            exitToLeft = False
        Case k_LOW_EARTH_FLAT_FLOOR
            exitToLeft = False
    End Select
    
    If roomBelow.LowNibble = k_LOW_RIVER Or roomBelow.LowNibble = k_LOW_NONE Then
        exitToLeft = False
    End If
    
    
    ' draw path
    If exitAbove Then
        proom.Line (k_ROOM_WIDTH_PIX / 2, 0)-(k_ROOM_WIDTH_PIX / 2, k_ROOM_HEIGHT_PIX / 2), RGB(255, 255, 0)
    End If
    
    If exitBelow Then
        proom.Line (k_ROOM_WIDTH_PIX / 2, k_ROOM_HEIGHT_PIX / 2)-(k_ROOM_WIDTH_PIX / 2, k_ROOM_HEIGHT_PIX), RGB(255, 255, 0)
    End If
    
    If exitToLeft Then
        proom.Line (0, k_ROOM_HEIGHT_PIX / 2)-(k_ROOM_WIDTH_PIX / 2, k_ROOM_HEIGHT_PIX / 2), RGB(255, 255, 0)
    End If
    
    If exitToRight Then
        proom.Line (k_ROOM_WIDTH_PIX / 2, k_ROOM_HEIGHT_PIX / 2)-(k_ROOM_WIDTH_PIX, k_ROOM_HEIGHT_PIX / 2), RGB(255, 255, 0)
    End If
    
    If thisRoom.HighNibble = k_HIGH_PLATFORM_LEFT Then
        proom.Line (0, k_ROOM_HEIGHT_PIX / 2)-(k_ROOM_WIDTH_PIX / 4, k_ROOM_HEIGHT_PIX / 2), RGB(255, 255, 0)
    ElseIf thisRoom.HighNibble = k_HIGH_PLATFORM_RIGHT Then
        proom.Line (k_ROOM_WIDTH_PIX, k_ROOM_HEIGHT_PIX / 2)-(k_ROOM_WIDTH_PIX / 4 * 3, k_ROOM_HEIGHT_PIX / 2), RGB(255, 255, 0)
    End If
End Sub

Private Sub DrawRoomFeature(ByVal hdcRoom As Long, ByRef room As CRoom, ByVal frameCount As Long)
    Dim animFrame As Long
    Dim hdcFeature As Long
    Dim hdcOld As Long
    Dim ret As Long
    Dim featureType As Integer

    Select Case room.HighNibble
        Case k_HIGH_SAVE_POINT
            featureType = k_FEATURE_SAVE_POINT
        Case k_HIGH_WATERFALL
            featureType = k_FEATURE_WATERFALL
        Case Else
            featureType = k_FEATURE_NONE
    End Select
    
'    With g_picBackground(room.Background)
'        hdcOld = SelectObject(hdcBackground, .m_frame(currFrame))
'        ret = BitBlt(proom.hdc, k_BACKGROUND_X, k_BACKGROUND_Y, .m_width, .m_height, hdcBackground, 0&, 0&, SRCCOPY)
'    End With
'    ret = SelectObject(hdcBackground, hdcOld)

    If Not CreateComponentDC(hdcRoom, hdcFeature) Then Exit Sub
    If featureType = k_FEATURE_SAVE_POINT Then
        With g_picFeature(k_FEATURE_SAVE_POINT)
            hdcOld = SelectObject(hdcFeature, .m_frame(0))
            ret = BitBlt(hdcRoom, k_FEATURE_SAVE_POINT_X, k_FEATURE_SAVE_POINT_Y, .m_width, .m_height, hdcFeature, 0&, 0&, SRCCOPY)
            ret = SelectObject(hdcFeature, hdcOld)
        End With
    End If

'    If room.balloon = True Then
'        With g_picFeature(k_FEATURE_LADDER)
'        hdcold = SelectObject(hDC, g_picFeature(k_FEATURE_BALLOON).m_frame(0))
'        ret = BitBlt(hdcRoom, k_FEATURE_BALLOON_X, k_FEATURE_BALLOON_Y, g_picFeature(k_FEATURE_BALLOON).m_Width, g_picFeature(k_FEATURE_BALLOON).m_Height, hDC, 0&, 0&, SRCCOPY)
'        ret = SelectObject(hdcFeature, hdcOld)
'        end with
'    End If

    If featureType = k_FEATURE_WATERFALL Then
        animFrame = frameCount Mod 2&
        With g_picFeature(k_FEATURE_WATERFALL)
            hdcOld = SelectObject(hdcFeature, g_picFeature(k_FEATURE_WATERFALL).m_frame(animFrame))
            ret = BitBlt(hdcRoom, k_FEATURE_WATERFALL_X, k_FEATURE_WATERFALL_Y, .m_width, .m_height, hdcFeature, 0&, 0&, SRCCOPY)
            ret = SelectObject(hdcFeature, hdcOld)
        End With
    End If

'    If room.lara = True Then
'        With g_picFeature(k_FEATURE_LADDER)
'            hdcold = SelectObject(hdc, g_picFeature(k_FEATURE_LARA).m_frame(0))
'            ret = BitBlt(hdcRoom, k_FEATURE_LARA_X, k_FEATURE_LARA_Y, g_picFeature(k_FEATURE_LARA).m_width, g_picFeature(k_FEATURE_LARA).m_height, hdc, 0&, 0&, SRCCOPY)
'            ret = SelectObject(hdcFeature, hdcOld)
'        End With
'    End If

'    If room.vine = True Then
'        With g_picFeature(k_FEATURE_LADDER)
'            hdcold = SelectObject(hdc, g_picFeature(k_FEATURE_VINE).m_frame(0))
'            ret = BitBlt(hdcRoom, k_FEATURE_VINE_X, k_FEATURE_VINE_Y, g_picFeature(k_FEATURE_VINE).m_width, g_picFeature(k_FEATURE_VINE).m_height, hdc, 0&, 0&, SRCCOPY)
'            ret = SelectObject(hdcFeature, hdcOld)
'        End With
'    End If

    ReleaseComponentDC hdcFeature
End Sub


Private Sub DrawRoomItem(ByVal hdcRoom As Long, ByRef room As CRoom, ByVal frameCount As Long)
    Dim animFrame As Integer
    Dim hdcItem As Long
    Dim hdcOld As Long
    Dim ret As Long
    Dim dx As Long
    Dim dy As Long
    Dim itemType As Integer
    
    Select Case room.HighNibble
        Case k_HIGH_QUICKCLAW
            itemType = k_ITEM_QUICKCLAW_CAT
        Case k_HIGH_GOLD_BAR_LEFT
            itemType = k_ITEM_GOLD_BAR_LEFT
        Case k_HIGH_STONE_RAT
            itemType = k_ITEM_STONE_RAT
        Case k_HIGH_RHONDA
            itemType = k_ITEM_RHONDA_GIRL
        Case k_HIGH_DIAMOND_RING
            itemType = k_ITEM_DIAMOND_RING
        Case k_HIGH_GOLD_BAR_RIGHT
            itemType = k_ITEM_GOLD_BAR_RIGHT
        Case Else
            itemType = k_ITEM_NONE
    End Select
    

    If Not CreateComponentDC(hdcRoom, hdcItem) Then Exit Sub
    If itemType = k_ITEM_NONE Then
    ElseIf itemType = k_ITEM_STONE_RAT Then
        animFrame = 0
        dx = k_ITEM_STONE_RAT_X
        dy = k_ITEM_STONE_RAT_Y
    ElseIf itemType = k_ITEM_QUICKCLAW_CAT Then
        animFrame = frameCount Mod 2&
        dx = k_ITEM_QUICKCLAW_CAT_X
        dy = k_ITEM_QUICKCLAW_CAT_Y
    ElseIf itemType = k_ITEM_DIAMOND_RING Then
        animFrame = 0
        dx = k_ITEM_DIAMOND_RING_X
        dy = k_ITEM_DIAMOND_RING_Y
    ElseIf itemType = k_ITEM_RHONDA_GIRL Then
        animFrame = 0
        dx = k_ITEM_RHONDA_GIRL_X
        dy = k_ITEM_RHONDA_GIRL_Y
    ElseIf itemType = k_ITEM_GOLD_BAR_LEFT Then
        animFrame = frameCount Mod 2&
        dx = k_ITEM_GOLD_BAR_LEFT_X
        dy = k_ITEM_GOLD_BAR_LEFT_Y
    ElseIf itemType = k_ITEM_GOLD_BAR_RIGHT Then
        animFrame = frameCount Mod 2&
        dx = k_ITEM_GOLD_BAR_RIGHT_X
        dy = k_ITEM_GOLD_BAR_RIGHT_Y
    Else
        Debug.Assert False
    End If

    If itemType <> k_ITEM_NONE Then
        With g_picItem(itemType)
            hdcOld = SelectObject(hdcItem, .m_mask(animFrame))
            ret = BitBlt(hdcRoom, dx, dy, .m_width, .m_height, hdcItem, 0&, 0&, SRCAND)

            ret = SelectObject(hdcItem, .m_frame(animFrame))
            ret = BitBlt(hdcRoom, dx, dy, .m_width, .m_height, hdcItem, 0&, 0&, SRCPAINT)
        End With
        ret = SelectObject(hdcItem, hdcOld)
    End If

    ReleaseComponentDC hdcItem
End Sub

Private Sub DrawRoomBackground(ByVal hdcRoom As Long, ByRef room As CRoom, ByVal frameCount As Long)
    Dim hdcBackground As Long
    Dim hdcOld As Long
    Dim ret As Long
    Dim backgroundType As Long
    

    If Not CreateComponentDC(hdcRoom, hdcBackground) Then Exit Sub
    Select Case room.LowNibble
        Case k_LOW_NONE
            backgroundType = k_BACKGROUND_NONE
        Case k_LOW_WATER
            backgroundType = k_BACKGROUND_WATER
        Case k_LOW_EARTH
            backgroundType = k_BACKGROUND_EARTH
        Case k_LOW_TREE_TOPS_1
            backgroundType = k_BACKGROUND_TREE_TOPS
        Case k_LOW_TREES_1
            backgroundType = k_BACKGROUND_TREES
        Case k_LOW_FLOOR_TWO_HOLES_AND_LADDER
            backgroundType = k_BACKGROUND_NONE
        Case k_LOW_CORRUPT_1
            backgroundType = k_BACKGROUND_NONE
        Case k_LOW_CORRUPT_2
            backgroundType = k_BACKGROUND_NONE
        Case k_LOW_EARTH_FLAT_FLOOR
            backgroundType = k_BACKGROUND_EARTH
        Case k_LOW_WALKWAY
            backgroundType = k_BACKGROUND_NONE
        Case k_LOW_SINGLE_HOLE
            backgroundType = k_BACKGROUND_NONE
        Case k_LOW_SINGLE_HOLE_AND_LADDER
            backgroundType = k_BACKGROUND_NONE
        Case k_LOW_RIVER
            backgroundType = k_BACKGROUND_WATER
        Case k_LOW_TREE_TOPS_2
            backgroundType = k_BACKGROUND_TREE_TOPS
        Case k_LOW_TREES_2
            backgroundType = k_BACKGROUND_TREES
        Case k_LOW_CORRUPT_3
            backgroundType = k_BACKGROUND_NONE
        Case Else
            Debug.Assert False
    End Select
    
    With g_picBackground(backgroundType)
        hdcOld = SelectObject(hdcBackground, .m_frame(0))
        ret = BitBlt(hdcRoom, k_BACKGROUND_X, k_BACKGROUND_Y, .m_width, .m_height, hdcBackground, 0&, 0&, SRCCOPY)
    End With
    ret = SelectObject(hdcBackground, hdcOld)
    If room.LowNibble = k_LOW_FLOOR_TWO_HOLES_AND_LADDER Or room.LowNibble = k_LOW_SINGLE_HOLE_AND_LADDER Then
        With g_picFeature(k_FEATURE_LADDER)
            hdcOld = SelectObject(hdcBackground, .m_frame(0))
            ret = BitBlt(hdcRoom, k_FEATURE_LADDER_X, k_FEATURE_LADDER_Y, .m_width, .m_height, hdcBackground, 0&, 0&, SRCCOPY)
            ret = SelectObject(hdcBackground, hdcOld)
        End With

    End If

    ret = DeleteDC(hdcBackground)
End Sub

Private Sub DrawRoomFloor(ByVal hdcRoom As Long, ByRef room As CRoom, ByVal frameCount As Long)
    Dim currFrame As Integer
    Dim floorType As Integer
    Dim hdcFloor As Long
    Dim hdcOld As Long
    Dim ret As Long
    Dim y As Long

    Select Case room.LowNibble
        Case k_LOW_NONE
            floorType = k_FLOOR_NONE
        Case k_LOW_WATER
            floorType = k_FLOOR_NONE
        Case k_LOW_EARTH
            floorType = k_FLOOR_NONE
        Case k_LOW_TREE_TOPS_1
            floorType = k_FLOOR_NONE
        Case k_LOW_TREES_1
            floorType = k_FLOOR_NONE
        Case k_LOW_FLOOR_TWO_HOLES_AND_LADDER
            floorType = k_FLOOR_WALKWAY_THREE_HOLES
        Case k_LOW_CORRUPT_1
            floorType = k_FLOOR_NONE
        Case k_LOW_CORRUPT_2
            floorType = k_FLOOR_NONE
        Case k_LOW_EARTH_FLAT_FLOOR
            floorType = k_FLOOR_SOLID
        Case k_LOW_WALKWAY
            floorType = k_FLOOR_WALKWAY
        Case k_LOW_SINGLE_HOLE
            floorType = k_FLOOR_WALKWAY_SINGLE_HOLE
        Case k_LOW_SINGLE_HOLE_AND_LADDER
            floorType = k_FLOOR_WALKWAY_HOLE_WITH_LADDER
        Case k_LOW_RIVER
            If room.HighNibble = k_HIGH_PLATFORM_LEFT Then
                floorType = k_FLOOR_PLAT_LEFT_WATER
            ElseIf room.HighNibble = k_HIGH_PLATFORM_RIGHT Then
                floorType = k_FLOOR_PLAT_RIGHT_WATER
            Else
                floorType = k_FLOOR_WATER
            End If
        Case k_LOW_TREE_TOPS_2
            floorType = k_FLOOR_NONE
        Case k_LOW_TREES_2
            floorType = k_FLOOR_NONE
        Case k_LOW_CORRUPT_3
            floorType = k_FLOOR_NONE
        Case Else
            Debug.Assert False
    End Select
    
    y = k_FLOOR_Y
    If floorType <> k_FLOOR_NONE Then
        If floorType = k_FLOOR_WATER Then
'        If floorType = k_FLOOR_WATER Or _
'        floorType = k_FLOOR_PLAT_LEFT_WATER Or _
'        floorType = k_FLOOR_PLAT_RIGHT_WATER Or _
'        floorType = k_FLOOR_PLAT_BOTH_WATER Then
            currFrame = frameCount Mod 4&
            If currFrame = 3 Then currFrame = 1
        Else
            currFrame = 0
        End If

        If Not CreateComponentDC(hdcRoom, hdcFloor) Then Exit Sub
        With g_picFloor(floorType)
            hdcOld = SelectObject(hdcFloor, .m_frame(currFrame))
            ret = BitBlt(hdcRoom, k_FLOOR_X, y, .m_width, .m_height, hdcFloor, 0&, 0&, SRCCOPY)
        End With
        ret = SelectObject(hdcFloor, hdcOld)
        ReleaseComponentDC hdcFloor
    End If

    y = k_FLOOR_Y + 32
    If room.HighNibble = k_HIGH_PLATFORM_LEFT Then
        floorType = k_FLOOR_PLAT_LEFT
    ElseIf room.HighNibble = k_HIGH_PLATFORM_RIGHT Then
        floorType = k_FLOOR_PLAT_RIGHT
    Else
        floorType = k_FLOOR_NONE
    End If
    If floorType <> k_FLOOR_NONE Then
        currFrame = 0
        If Not CreateComponentDC(hdcRoom, hdcFloor) Then Exit Sub
        With g_picFloor(floorType)
            hdcOld = SelectObject(hdcFloor, .m_frame(currFrame))
            ret = BitBlt(hdcRoom, k_FLOOR_X, y, .m_width, .m_height, hdcFloor, 0&, 0&, SRCCOPY)
        End With
        ret = SelectObject(hdcFloor, hdcOld)
        ReleaseComponentDC hdcFloor
    End If

End Sub

Private Sub DrawRoomExit(ByVal hdcRoom As Long, ByRef room As CRoom, ByVal frameCount As Long)
    Dim hdcExit As Long
    Dim hdcOld As Long
    Dim ret As Long

'    If room.ExitLeft <> k_EXIT_LEFT_OPEN Then
'        If Not CreateComponentDC(hdcRoom, hdcExit) Then Exit Sub
'        With g_picExitLeft(room.ExitLeft)
'            hdcOld = SelectObject(hdcExit, .m_frame(0))
'            ret = BitBlt(hdcRoom, k_EXIT_LEFT_X, k_EXIT_LEFT_Y, .m_width, .m_height, hdcExit, 0&, 0&, SRCCOPY)
'        End With
'        ret = SelectObject(hdcExit, hdcOld)
'        ReleaseComponentDC hdcExit
'    End If

    If room.ExitFlag Then
        If Not CreateComponentDC(hdcRoom, hdcExit) Then Exit Sub
        With g_picExitRight(k_EXIT_RIGHT_DARK_ROCK_BLACK)
            hdcOld = SelectObject(hdcExit, .m_frame(0))
            ret = BitBlt(hdcRoom, k_EXIT_RIGHT_X, k_EXIT_RIGHT_Y, .m_width, .m_height, hdcExit, 0&, 0&, SRCCOPY)
        End With
        ret = SelectObject(hdcExit, hdcOld)
        ReleaseComponentDC hdcExit
    End If
End Sub

Private Sub DrawRoomCreature(ByVal hdcRoom As Long, ByRef room As CRoom, ByVal frameCount As Long)
    Dim animFrame As Integer
    Dim hdcCreature As Long
    Dim hdcOld As Long
    Dim ret As Long
    Dim dx As Long
    Dim dy As Long
    Dim flightPathPos As Integer
    Dim creatureType As Integer
    
    Select Case room.HighNibble
        Case k_HIGH_SCORPION
            creatureType = k_CREATURE_SCORPION
        Case k_HIGH_BAT
            If room.LowNibble = k_LOW_WATER Then
                creatureType = k_CREATURE_EEL
            Else
                creatureType = k_CREATURE_BAT
            End If
        Case k_HIGH_CONDOR
            creatureType = k_CREATURE_CONDOR
        Case k_HIGH_FROG
            creatureType = k_CREATURE_FROG
        Case Else
            creatureType = k_CREATURE_NONE
    End Select

    If creatureType = k_CREATURE_NONE Then
    ElseIf creatureType = k_CREATURE_BAT Then
        animFrame = frameCount Mod 2&
        dx = frameCount Mod 160&
        flightPathPos = frameCount Mod k_FLIGHT_PATH_LEN
        dy = k_CREATURE_BAT_Y + g_flightPath(flightPathPos)
    ElseIf creatureType = k_CREATURE_SCORPION Then
        animFrame = frameCount Mod 2&
        dx = frameCount * 2& Mod 160&
        dy = k_CREATURE_SCORPION_Y
    ElseIf creatureType = k_CREATURE_EEL Then
        animFrame = frameCount Mod 2&
        If Int(Rnd * 100) > 50 Then animFrame = animFrame + 2
        dx = frameCount * 2& Mod 160&
        dy = k_CREATURE_EEL_Y
    ElseIf creatureType = k_CREATURE_CONDOR Then
        animFrame = frameCount Mod 2&
        dx = 160 - (frameCount Mod 160&)
        flightPathPos = (frameCount \ 2&) Mod k_FLIGHT_PATH_LEN
        dy = k_CREATURE_CONDOR_Y + g_flightPath(flightPathPos)
    ElseIf creatureType = k_CREATURE_FROG Then
        animFrame = frameCount Mod 2&
        dx = 80
        dy = k_CREATURE_FROG_Y
    Else
        animFrame = 0
        dx = 0
        dy = 0
    End If

    If creatureType <> k_CREATURE_NONE Then
        If Not CreateComponentDC(hdcRoom, hdcCreature) Then Exit Sub
        With g_picCreature(creatureType)
            hdcOld = SelectObject(hdcCreature, .m_mask(animFrame))
            ret = BitBlt(hdcRoom, dx, dy, .m_width, .m_height, hdcCreature, 0&, 0&, SRCAND)

            ret = SelectObject(hdcCreature, .m_frame(animFrame))
            ret = BitBlt(hdcRoom, dx, dy, .m_width, .m_height, hdcCreature, 0&, 0&, SRCPAINT)
        End With

        ret = SelectObject(hdcCreature, hdcOld)
        ReleaseComponentDC hdcCreature
    End If

End Sub

