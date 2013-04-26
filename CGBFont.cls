VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGBFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TileOrientations
    k_TILE_FIRST = 0
    k_TILE_NORMAL = k_TILE_FIRST
    k_TILE_HFLIP
    k_TILE_VFLIP
    k_TILE_HVFLIP
    k_TILE_LAST = k_TILE_HVFLIP
End Enum

Private m_picFont(k_TILE_FIRST To k_TILE_LAST) As Picture
Private m_tileCount As Integer

Private Function AddTile()
    ' add a tile to the font set if it doesn't already exist
    AddTile = False
End Function

Public Sub RemoveTile()
    ' remove a tile from the font set
End Sub

Public Function PaletteCount() As Integer

End Function

Public Function TileCount() As Integer

End Function

Public Function GetPaletteFrequency(ByVal paletteNum As Integer) As Integer
    GetPaletteFrequency = 0
End Function

Public Function GetPalette(ByVal paletteNum As Integer) As CGBPalette

End Function

Public Function GetPaletteForTile(ByVal tileNum As Integer) As CGBPalette

End Function

'Public Function GetTile(ByVal tileNum As Integer) As CGBTile
'
'End Function

Public Function GetTileBitmap(ByVal tileNum As Integer, ByVal flipCode) As Picture
    Debug.Assert flipCode >= k_TILE_FIRST And flipCode <= k_TILE_LAST
    Debug.Assert tileNum >= 0 And tileNum <= m_tileCount
End Function

Private Sub Class_Initialize()
    m_tileCount = 0
    Debug.Assert False
    m_picFont(0).ScaleMode = vbPixels
    m_picFont(0).ScaleWidth = 100
End Sub

