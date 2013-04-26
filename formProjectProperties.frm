VERSION 5.00
Begin VB.Form formProjectProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Properties"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Export Information"
      Height          =   1815
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   6495
      Begin VB.TextBox textLabel 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   4575
      End
      Begin VB.CommandButton cmdBrowseFilename 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   5280
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox textFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label8 
         Caption         =   "&Label:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "&Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox textVersionInfo 
      Height          =   735
      Left            =   1920
      MaxLength       =   1024
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Version Information"
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   6495
      Begin VB.ListBox listVersionInfo 
         Height          =   645
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "T&ype:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Version Number"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox checkAutoIncrement 
         Caption         =   "A&uto Increment"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox textRevision 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox textMinor 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox textMajor 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "&Major:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "M&inor:"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "&Revision:"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Size"
      Height          =   1455
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.TextBox textH 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox textW 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Height:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Width:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "formProjectProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_MAJOR_VERSION_MIN = 0
Private Const k_MAJOR_VERSION_MAX = 1000
Private Const k_MINOR_VERSION_MIN = 0
Private Const k_MINOR_VERSION_MAX = 1000
Private Const k_REVISION_VERSION_MIN = 0
Private Const k_REVISION_VERSION_MAX = 1000


Private m_majorVersion As Integer
Private m_minorVersion As Integer
Private m_revisionVersion As Integer
Private m_autoIncrement As Boolean
Private m_levelInfo As String
Private m_levelComments As String
Private m_companyName As String
Private m_legalCopyright As String
Private m_mapWidth As Integer
Private m_mapHeight As Integer
Private m_exportFilespec As String
Private m_exportLabel As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not SetProperties Then
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Initialise
End Sub

' validateproperties -- validates the properties before performing a setproperties
Private Function ValidateProperties()
    Dim result As Integer
    
    ' update map width & height
    m_mapWidth = val(textW.Text)
    m_mapHeight = val(textH.Text)
    If m_mapWidth < g_map.Width Or m_mapHeight < g_map.Height Then
        result = MsgBox("The new map size is smaller than the current size. Some information will be lost. Proceed?", vbYesNo, "Change Map Size")
        If result = vbNo Then
            ValidateProperties = False
            Exit Function
        End If
    End If
    
    ' update version information
    m_minorVersion = val(textMinor.Text)
    m_majorVersion = val(textMajor.Text)
    m_revisionVersion = val(textRevision.Text)
    m_autoIncrement = checkAutoIncrement.value
    
    ' update project information
    ' update filename
    m_exportFilespec = textFilename.Text
    ' update label
    m_exportLabel = textLabel.Text
    If Len(m_exportLabel) = 0 Then
        result = MsgBox("A label for the export must be specified.", vbOKOnly, "Export Label")
        ValidateProperties = False
        Exit Function
    End If
    
    ValidateProperties = True
End Function

' updateinformation -- sets the information in the form from the project properties
Private Sub UpdateInformation()
    textW.Text = Trim(m_mapWidth)
    textH.Text = Trim(m_mapHeight)
    listVersionInfo.Clear
    listVersionInfo.AddItem "LevelInfo"
    listVersionInfo.AddItem "Comments"
    listVersionInfo.AddItem "CompanyName"
    listVersionInfo.AddItem "LegalCopyright"
End Sub

' getproperties -- retrieves the properties from the project
Private Sub GetProperties()
    m_mapWidth = g_map.Width
    m_mapHeight = g_map.Height
End Sub

' setproperties -- updates the project with the new properties
Private Function SetProperties() As Boolean
    Dim result As Integer
    
    If Not ValidateProperties Then
        SetProperties = False
        Exit Function
    End If
    
    g_map.Resize m_mapWidth, m_mapHeight
    SetProperties = True
End Function

Private Sub Initialise()
    GetProperties
    UpdateInformation
End Sub

Private Sub textRevision_KeyPress(KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> Asc(vbBack) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub textRevision_Validate(Cancel As Boolean)
    ' If the value is a number and 8<= number <=32, keep the focus
    If Not IsNumeric(textRevision.Text) Or val(textRevision.Text) < k_REVISION_VERSION_MIN Or val(textRevision.Text) > k_REVISION_VERSION_MAX Then
        Cancel = True
        MsgBox "Please insert a number between " & Trim(k_REVISION_VERSION_MIN) & " and " & Trim(k_REVISION_VERSION_MAX), , "Invalid Revision Version"
    End If
End Sub

Private Sub textMinor_KeyPress(KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> Asc(vbBack) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub textMinor_Validate(Cancel As Boolean)
    ' If the value is a number and 8<= number <=32, keep the focus
    If Not IsNumeric(textMinor.Text) Or val(textMinor.Text) < k_MINOR_VERSION_MIN Or val(textMinor.Text) > k_MINOR_VERSION_MAX Then
        Cancel = True
        MsgBox "Please insert a number between " & Trim(k_MINOR_VERSION_MIN) & " and " & Trim(k_MINOR_VERSION_MAX), , "Invalid Minor Version"
    End If
End Sub

Private Sub textMajor_KeyPress(KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> Asc(vbBack) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub textMajor_Validate(Cancel As Boolean)
    ' If the value is a number and 8<= number <=32, keep the focus
    If Not IsNumeric(textMajor.Text) Or val(textMajor.Text) < k_MINOR_VERSION_MIN Or val(textMajor.Text) > k_MINOR_VERSION_MAX Then
        Cancel = True
        MsgBox "Please insert a number between " & Trim(k_MINOR_VERSION_MIN) & " and " & Trim(k_MINOR_VERSION_MAX), , "Invalid Major Version"
    End If
End Sub

Private Sub textH_KeyPress(KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> Asc(vbBack) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub textH_Validate(Cancel As Boolean)
    ' If the value is a number and 8<= number <=32, keep the focus
    If Not IsNumeric(textH.Text) Or val(textH.Text) < 8 Or val(textH.Text) > 32 Then
        Cancel = True
        MsgBox "Please insert a number between 8 and 32", , "Invalid Height"
    End If

End Sub

Private Sub textW_KeyPress(KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> Asc(vbBack) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub textW_Validate(Cancel As Boolean)
    ' If the value is a number and 8<= number <=32, keep the focus
    If Not IsNumeric(textW.Text) Or val(textW.Text) < 8 Or val(textW.Text) > 32 Then
        Cancel = True
        MsgBox "Please insert a number between 8 and 32", , "Invalid Width"
    End If

End Sub
