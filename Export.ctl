VERSION 5.00
Begin VB.UserControl Export 
   CanGetFocus     =   0   'False
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   4485
End
Attribute VB_Name = "Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_export As String

Private Sub UserControl_Initialize()
    m_export = ""
End Sub
