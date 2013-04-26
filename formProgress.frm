VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1185
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar progressBar 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label labelProgress 
      Caption         =   "Progress Dialogue"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "formProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = "Progress Dialogue"
    labelProgress.Caption = "Progress Dialogue"
    progressBar.Min = 0
    progressBar.Value = 0
    progressBar.Max = 100
End Sub


Public Sub StartProgress(ByVal title As String, ByVal info As String, ByVal distance As Integer)
    Me.Caption = title
    labelProgress.Caption = info
    progressBar.Max = distance
    progressBar.Value = 0
    Me.show
    DoEvents
End Sub


Public Sub UpdateProgress(ByVal progress As Integer)
    progressBar.Value = progress
    DoEvents
End Sub

Public Sub EndProgress()
    Unload Me
End Sub
