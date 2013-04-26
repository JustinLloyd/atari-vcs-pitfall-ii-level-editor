VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lowNibble As Byte
Private m_highNibble As Byte
Private m_exitFlag As Boolean

Private m_hasChanged As Boolean

Public Property Get HasChanged() As Boolean
    HasChanged = m_hasChanged
End Property

Public Sub ClearChange()
    m_hasChanged = False
End Sub

Private Sub FlagRoomChange()
    m_hasChanged = True
End Sub

Public Property Get ExitFlag() As Boolean
    ExitFlag = m_exitFlag
End Property

Public Property Let ExitFlag(ByVal val As Boolean)
    m_exitFlag = val
End Property

Public Property Get LowNibble() As Byte
    LowNibble = m_lowNibble
End Property

Public Property Let LowNibble(ByVal val As Byte)
    Debug.Assert val >= 0 And val <= 15
    m_lowNibble = val
End Property

Public Property Get HighNibble() As Byte
    HighNibble = m_highNibble
End Property

Public Property Let HighNibble(ByVal val As Byte)
    Debug.Assert val >= 0 And val <= 15
    m_highNibble = val
End Property

Public Function NibblePair() As Byte
    NibblePair = LShift(m_highNibble, 4) Or m_lowNibble
End Function

Private Sub Class_Initialize()
    InitialiseRoom
End Sub

Public Sub SerialIn(ByVal id As String, ByRef proj As CProject)
    Dim prop As PropertyBag
    Dim result As Boolean
    
    Set prop = New PropertyBag
    result = proj.GetContent(id, prop)
    Debug.Assert result <> False
    m_lowNibble = prop.ReadProperty("LowNibble", 0)
    m_highNibble = prop.ReadProperty("HighNibble", 0)
    m_exitFlag = prop.ReadProperty("ExitFlag", False)
End Sub

Public Sub SerialOut(ByVal id As String, ByRef proj As CProject)
    Dim prop As PropertyBag
    Dim result As Boolean
    
    Set prop = New PropertyBag
    prop.WriteProperty "LowNibble", m_lowNibble
    prop.WriteProperty "HighNibble", m_highNibble
    prop.WriteProperty "ExitFlag", m_exitFlag
    result = proj.AddContent(id, prop)
    Debug.Assert result <> False
End Sub

Public Sub Copy(ByRef srcRoom As CRoom)
    FlagRoomChange
    m_lowNibble = srcRoom.LowNibble
    m_highNibble = srcRoom.HighNibble
End Sub

Public Sub Reset()
    InitialiseRoom
    FlagRoomChange
End Sub

Private Sub InitialiseRoom()
    m_hasChanged = False
    m_lowNibble = 0
    m_highNibble = 0
End Sub

