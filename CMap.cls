VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_undoCollection As Collection
Private m_mapWidth As Integer
Private m_mapHeight As Integer
Private m_startRoom As Integer              ' first room that player starts in on this map
Private m_hasChanged As Boolean

Private m_room(0 To k_MAX_MAP_WIDTH - 1, 0 To k_MAX_MAP_HEIGHT - 1) As CRoom

Public Sub ClearChange()
    Dim col As Integer
    Dim row As Integer
    
    m_hasChanged = False
    For col = 0 To k_MAX_MAP_WIDTH - 1
        For row = 0 To k_MAX_MAP_HEIGHT - 1
            m_room(col, row).ClearChange
        Next
    Next
End Sub

Public Function HasChanged() As Boolean
    Dim col As Integer
    Dim row As Integer
    
    If m_hasChanged Then
        HasChanged = True
        Exit Function
    End If
    
    For col = 0 To k_MAX_MAP_WIDTH - 1
        For row = 0 To k_MAX_MAP_HEIGHT - 1
            If m_room(col, row).HasChanged Then
                HasChanged = True
                Exit Function
            End If
        Next
    Next
    
    HasChanged = False
End Function

Private Sub FlagMapChange()
    m_hasChanged = True
End Sub

Public Sub Resize(ByVal numCol As Integer, ByVal numRow As Integer)
    Debug.Assert numCol >= k_MIN_MAP_WIDTH And numCol <= k_MAX_MAP_WIDTH
    Debug.Assert numRow >= k_MIN_MAP_HEIGHT And numRow <= k_MAX_MAP_HEIGHT
    
    If GetCol(m_startRoom) > numCol Or GetRow(m_startRoom) > numRow Then
        m_startRoom = 0
    End If
    
    m_mapWidth = numCol
    m_mapHeight = numRow
'    CreateMap numCol, numRow
    FlagMapChange
End Sub

Public Property Let StartRoom(ByVal roomNum As Integer)
    m_startRoom = roomNum
    FlagMapChange
End Property

Public Property Get StartRoom() As Integer
    StartRoom = m_startRoom
End Property

Public Function Height() As Integer
    Height = m_mapHeight
End Function

Public Function Width() As Integer
    Width = m_mapWidth
End Function

Public Function GetRoom(ByVal col As Integer, ByVal row As Integer) As CRoom
    Debug.Assert col >= 0 And col < m_mapWidth
    Debug.Assert row >= 0 And row < m_mapHeight
    If col < 0 Or col >= m_mapWidth Or row < 0 Or row >= m_mapHeight Then
        Set GetRoom = Null
    Else
        Set GetRoom = m_room(col, row)
    End If
    
End Function

Private Function GetCol(ByVal roomNum As Integer) As Integer
    GetCol = roomNum \ m_mapHeight
End Function

Private Function GetRow(ByVal roomNum As Integer) As Integer
    GetRow = roomNum Mod m_mapHeight
End Function

Public Function GetRoomNum(ByVal col As Integer, ByVal row As Integer) As Integer
    GetRoomNum = col * Height + row
End Function

Private Sub Class_Initialize()
    m_hasChanged = False
    CreateMap k_MAX_MAP_WIDTH, k_MAX_MAP_HEIGHT
End Sub

Private Sub CreateMap(ByVal w As Integer, ByVal h As Integer)
    Dim col As Integer
    Dim row As Integer
    
    m_startRoom = 0
    m_mapWidth = w
    m_mapHeight = h
    For row = 0 To m_mapHeight - 1
        For col = 0 To m_mapWidth - 1
            Set m_room(col, row) = New CRoom
        Next
    Next
End Sub

Public Sub SerialOut(ByRef proj As CProject)
    Dim prop As PropertyBag
    Dim col As Integer
    Dim row As Integer
    Dim result As Boolean
    Dim id As String
    
'    formProgress.StartProgress "Compressing Level", "Compressing Level Data", m_rows
    Set prop = New PropertyBag
    prop.WriteProperty "Cols", m_mapWidth
    prop.WriteProperty "Rows", m_mapHeight
    prop.WriteProperty "StartRoom", m_startRoom
    result = proj.AddContent("Map", prop)
    Debug.Assert result <> False
    
    For row = 0 To m_mapHeight - 1
        For col = 0 To m_mapWidth - 1
            id = "Cell-" & Trim(col) & "-" & Trim(row)
            m_room(col, row).SerialOut id, proj
        Next
'        formProgress.UpdateProgress row
    Next
    
'    formProgress.EndProgress
End Sub

Public Sub SerialIn(ByRef proj As CProject)
    Dim col As Integer
    Dim numCols As Integer
    Dim row As Integer
    Dim numRows As Integer
    Dim prop As PropertyBag
    Dim result As Boolean
    Dim id As String
    Dim StartRoom As Integer
    
    Set prop = New PropertyBag
    result = proj.GetContent("Map", prop)
    Debug.Assert result <> False
    numCols = prop.ReadProperty("Cols", k_MAX_MAP_WIDTH)
    numRows = prop.ReadProperty("Rows", k_MAX_MAP_HEIGHT)
    StartRoom = prop.ReadProperty("StartRoom", 0)
'    formProgress.StartProgress "Decompressing Level", "Decompressing Level Data", numRows
    CreateMap numCols, numRows
    m_startRoom = StartRoom
    For col = 0 To m_mapWidth - 1
        For row = 0 To m_mapHeight - 1
            id = "Cell-" & Trim(col) & "-" & Trim(row)
            m_room(col, row).SerialIn id, proj
        Next
'        formProgress.UpdateProgress row
    Next
    
'    formProgress.EndProgress
End Sub

Private Sub ExportVCSExits(ByRef export As CVCSImage)
    Dim col As Integer
    Dim row As Integer
    Dim addr As Long
    Dim i As Long
    Dim bitShift As Integer
    Dim exitBits As Byte
    
    addr = 3174 ' $0C66
    For row = 0 To m_mapHeight - 1
        bitShift = 128
        exitBits = 0
'        exitBits = export.GetByte(addr)
        For col = 0 To m_mapWidth - 1
            If m_room(col, row).ExitFlag Then
                exitBits = exitBits Or bitShift
            End If
            bitShift = bitShift / 2
        Next
            
        export.SetByte addr, exitBits
        addr = addr + 1
    Next
    
End Sub

Private Sub ExportVCSRooms(ByRef export As CVCSImage, ByVal exportCreatures As Boolean)
    Dim col As Integer
    Dim row As Integer
    Dim addr As Long
    Dim HighNibble As Byte
    Dim LowNibble As Byte
    Dim room As CRoom
    
    addr = 3271 ' $0CC7
    For row = 0 To m_mapHeight - 1
        For col = 0 To m_mapWidth - 1
            If exportCreatures Then
                    export.SetByte addr, m_room(col, row).NibblePair
            Else
                HighNibble = m_room(col, row).HighNibble
                If HighNibble = k_HIGH_BAT Or _
                    HighNibble = k_HIGH_CONDOR Or _
                    HighNibble = k_HIGH_FROG Or _
                    HighNibble = k_HIGH_SCORPION Then
                    export.SetByte addr, m_room(col, row).LowNibble
                Else
                    export.SetByte addr, m_room(col, row).NibblePair
                End If
            End If
            addr = addr + 1
        Next
    Next

End Sub

Public Sub ExportBinary(ByRef export As CVCSImage, ByVal exportCreatures As Boolean)
    ExportVCSExits export
    ExportVCSRooms export, exportCreatures
    export.DumpROM
End Sub

'Public Sub SetVCSLowNibble(ByVal col As Integer, ByVal row As Integer, ByVal nibble As Byte)
'    m_room(col, row).Background = k_BACKGROUND_NONE
'    m_room(col, row).Floor = k_FLOOR_NONE
'    m_room(col, row).ExitLeft = k_EXIT_LEFT_OPEN
'    m_room(col, row).ExitRight = k_EXIT_RIGHT_OPEN
'    If row > 1 Then
'        m_room(col, row - 1).Ladder = False
'    End If
'    Select Case nibble
'        Case 0
'            ' nothing
'        Case 1
'            m_room(col, row).Background = k_BACKGROUND_WATER
'        Case 2
'            m_room(col, row).Background = k_BACKGROUND_EARTH
'        Case 3
'            m_room(col, row).Background = k_BACKGROUND_TREE_TOPS
'        Case 4
'            m_room(col, row).Background = k_BACKGROUND_TREES
'        Case 5
'            m_room(col, row).Floor = k_FLOOR_WALKWAY_THREE_HOLES
'        Case 6
'            m_room(col, row).Creature = k_CREATURE_CONDOR
'        Case 7
'            m_room(col, row).Item = k_ITEM_GOLD_BAR_LEFT
'        Case 8
'            m_room(col, row).Item = k_ITEM_STONE_RAT
'        Case 9
'            m_room(col, row).Waterfall = True
'        Case 10
'            ' investigate
'        Case 11
'            m_room(col, row).Item = k_ITEM_RHONDA_GIRL
'        Case 12
'            m_room(col, row).Item = k_ITEM_DIAMOND_RING
'        Case 13
'            m_room(col, row).Balloon = True
'        Case 14
'            m_room(col, row).Creature = k_CREATURE_FROG
'        Case 15
'            m_room(col, row).Item = k_ITEM_GOLD_BAR_RIGHT
'    End Select
'End Sub
'
'Public Sub SetVCSHighNibble(ByVal col As Integer, ByVal row As Integer, ByVal nibble As Byte)
'    m_room(col, row).Item = k_ITEM_NONE
'    m_room(col, row).Creature = k_CREATURE_NONE
'    m_room(col, row).SavePoint = False
'    m_room(col, row).Waterfall = False
'    m_room(col, row).Balloon = False
'    Select Case nibble
'        Case 0
'            ' nothing
'        Case 1
'            m_room(col, row).SavePoint = 1
'        Case 2
'            ' investigate
'        Case 3
'            m_room(col, row).Item = k_ITEM_QUICKCLAW_CAT
'        Case 4
'            m_room(col, row).Creature = k_CREATURE_SCORPION
'        Case 5
'            m_room(col, row).Creature = k_CREATURE_BAT
'        Case 6
'            m_room(col, row).Creature = k_CREATURE_CONDOR
'        Case 7
'            m_room(col, row).Item = k_ITEM_GOLD_BAR_LEFT
'        Case 8
'            m_room(col, row).Item = k_ITEM_STONE_RAT
'        Case 9
'            m_room(col, row).Waterfall = True
'        Case 10
'            ' investigate
'        Case 11
'            m_room(col, row).Item = k_ITEM_RHONDA_GIRL
'        Case 12
'            m_room(col, row).Item = k_ITEM_DIAMOND_RING
'        Case 13
'            m_room(col, row).Balloon = True
'        Case 14
'            m_room(col, row).Creature = k_CREATURE_FROG
'        Case 15
'            m_room(col, row).Item = k_ITEM_GOLD_BAR_RIGHT
'    End Select
'End Sub

'Public Property Get Serialize() As PropertyBag
'    Dim prop As PropertyBag
'    Dim col As Integer
'    Dim row As Integer
'    Dim cellProp As PropertyBag
'    Dim tempProp As PropertyBag
'
''    formProgress.StartProgress "Compressing Level", "Compressing Level Data", m_rows
'    Set prop = New PropertyBag
'    prop.WriteProperty "Cols", m_mapWidth
'    prop.WriteProperty "Rows", m_mapHeight
'    prop.WriteProperty "StartRoom", m_startRoom
''    prop.WriteProperty "Pic", m_backgroundPicture, Nothing
'    Set cellProp = New PropertyBag
'    Set tempProp = New PropertyBag
'    For row = 0 To m_mapHeight - 1
'        For col = 0 To m_mapWidth - 1
'            cellProp.Contents = tempProp.Contents
'            m_room(col, row).SerializeOut cellProp
'            prop.WriteProperty "Cell-" & Trim(col) & "-" & Trim(row), cellProp.Contents
'        Next
''        formProgress.UpdateProgress row
'    Next
'
'    Set Serialize = prop
''    formProgress.EndProgress
'End Property
'
'
'Public Property Set Serialize(ByRef prop As PropertyBag)
'    Dim col As Integer
'    Dim numCols As Integer
'    Dim row As Integer
'    Dim numRows As Integer
'    Dim cellProp As PropertyBag
'    Dim StartRoom As Integer
'
'    numCols = prop.ReadProperty("Cols", k_MAX_MAP_WIDTH)
'    numRows = prop.ReadProperty("Rows", k_MAX_MAP_HEIGHT)
'    StartRoom = prop.ReadProperty("StartRoom", 0)
''    formProgress.StartProgress "Decompressing Level", "Decompressing Level Data", numRows
''    Set m_backgroundPicture = prop.ReadProperty("Pic")
''    m_initialised = False
'    CreateMap numCols, numRows
'    m_startRoom = StartRoom
'    Set cellProp = New PropertyBag
'    For col = 0 To m_mapWidth - 1
'        For row = 0 To m_mapHeight - 1
'            cellProp.Contents = prop.ReadProperty("Cell-" & Trim(col) & "-" & Trim(row))
'            m_room(col, row).SerializeIn cellProp
'        Next
''        formProgress.UpdateProgress row
'    Next
'
''    formProgress.EndProgress
'End Property
'
'
