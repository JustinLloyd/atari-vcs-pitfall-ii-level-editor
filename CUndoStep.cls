VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUndoStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_undoSteps As Collection

Private Sub Class_Initialize()
    Set m_undoSteps = New Collection
End Sub

' add an undo-able step
Public Sub AddStep()

End Sub

' move to the previous step
Public Sub UndoStep()

End Sub

' move to the next step
Public Sub RedoStep()

End Sub

' clear the current undo step and subsequent undo steps
Public Sub ClearStep()

End Sub

' completely reset the undo list
Public Sub Reset()

End Sub
