VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_versionMajor As Integer
Private m_versionMinor As Integer
Private m_versionRevision As Integer

Private m_projectData As Collection

Public Function WriteContents(ByVal filename As String) As Boolean
    Dim prop As PropertyBag
    Dim F As Integer
    Dim itemCount As Integer
    
    On Error GoTo WriteError
    F = FreeFile
    Open filename For Binary Access Write As F
    itemCount = m_projectData.Count
    Put #F, , itemCount
    For Each prop In m_projectData
        Put #F, , prop.Contents
    Next
    
    Close F
    
    WriteContents = True
    On Error GoTo 0
    Exit Function
WriteError:
    Debug.Assert False 'uh oh
    On Error GoTo 0
    WriteContents = False
End Function

Public Function ReadContents(ByVal filename As String) As Boolean
    Dim prop As PropertyBag
    Dim temp As Variant
    Dim F As Integer
    Dim index As Integer
    Dim itemCount As Integer
    Dim id As String
    
    On Error GoTo ReadError
    If Not DoesFileExist(filename) Then
        MsgBox "The specified file was not found", , "File Not Found"
        ReadContents = False
        Exit Function
    End If
    
    If FileLen(filename) = 0 Then
        MsgBox "The specified file contains no data", , "Invalid file"
        ReadContents = False
        Exit Function
    End If
    
    F = FreeFile
    Open filename For Binary Access Read As F
    Get #F, , itemCount
    For index = 1 To itemCount
        Set prop = New PropertyBag
        Get #F, , temp
        prop.Contents = temp
        id = prop.ReadProperty("UniqueCollectionID")
        m_projectData.Add prop, id
    Next
    
    Close F
    
    ReadContents = True
    On Error GoTo 0
    Exit Function
ReadError:
    Debug.Assert False 'uh oh
    On Error GoTo 0
    ReadContents = False
End Function

Public Function AddContent(ByVal id As String, ByRef prop As PropertyBag) As Boolean
    Debug.Assert Len(id) > 0 And Len(id) <= 64
    
    On Error GoTo AddError
    prop.WriteProperty "UniqueCollectionID", id
    m_projectData.Add prop, id
    AddContent = True
    Exit Function
AddError:
    Debug.Assert False 'uh oh
    AddContent = False
    On Error GoTo 0
End Function

Public Function GetContent(ByVal id As String, ByRef prop As PropertyBag) As Boolean
    Debug.Assert Len(id) > 0 And Len(id) <= 64
    
    On Error GoTo GetError
    prop.Contents = m_projectData.Item(id).Contents
    GetContent = True
    On Error GoTo 0
    Exit Function
GetError:
    Debug.Assert False 'uh oh
    GetContent = False
    On Error GoTo 0
End Function

Private Sub Class_Initialize()
    Set m_projectData = New Collection
End Sub

