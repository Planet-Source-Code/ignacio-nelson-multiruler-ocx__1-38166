VERSION 5.00
Begin VB.UserControl MultiRuler 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "MultiRuler.ctx":0000
   ScaleHeight     =   315
   ScaleWidth      =   4815
   ToolboxBitmap   =   "MultiRuler.ctx":0014
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Kontext"
      Visible         =   0   'False
      Begin VB.Menu mnuMode 
         Caption         =   "Centimeter"
         Index           =   0
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Inch"
         Index           =   1
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Pixel"
         Index           =   2
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Twip"
         Index           =   3
      End
   End
End
Attribute VB_Name = "MultiRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Enum sedBorderStyle
    rlrNoBorder = 0
    rlrSunken = 1
    rlrSunkenOuter = 2
    rlrRaised = 3
    rlrRaisedInner = 4
    rlrBump = 5
    rlrEtched = 6
End Enum

Public Enum sedBorderWidth
    rlrNone
    rlrSingle
    rlrDouble
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Const For Border Style

Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000

' Const For Object Sides
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Const m_Def_BorderStyle = 2
Const m_Def_Orientation = 0

Private x1 As Single
Private RScale As Long

Public Enum RulerModeConst
    Millimeters = 0
    Inch = 1
    Pixel = 2
    Twips = 3
End Enum

Public Enum rlrOrientationConstants
    rlrHorizontal = 0
    rlrVertival = 1
End Enum



Private m_BorderStyle As sedBorderStyle
Private m_Mode As RulerModeConst
Private m_Orientation As rlrOrientationConstants

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Dim rctBrd As RECT

Private lBorderDef As Long

Event Click()
Event DblClick()


Public Property Get RulerMode() As RulerModeConst
    RulerMode = m_Mode
End Property

Public Property Let RulerMode(New_Mode As RulerModeConst)
    m_Mode = New_Mode
    Select Case m_Mode
        Case 0
            RScale = 570
        Case 1
            RScale = 1440
        Case 2
            RScale = Screen.TwipsPerPixelX * 100
        Case 3
            RScale = 1000
    End Select
    Picture2.Cls
    DrawRuler
    PropertyChanged "RulerMode"
End Property

Public Property Get Orientation() As rlrOrientationConstants
    Orientation = m_Orientation
End Property

Public Property Let Orientation(New_Val As rlrOrientationConstants)
    m_Orientation = New_Val
    UserControl.Cls
    Picture2.Cls
    ChangeOrientSizes
    DrawRuler
    PropertyChanged "Orientation"
End Property

Public Property Get BorderStyle() As sedBorderStyle
    BorderStyle = m_BorderStyle
    lBorderDef = BorderStyle
End Property

Public Property Let BorderStyle(New_Val As sedBorderStyle)
    m_BorderStyle = New_Val
    Picture2.Cls
    EdgeSubClass Picture2.hWnd, New_Val
    DrawRuler
    lBorderDef = New_Val
    PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = Picture2.BackColor
End Property

Public Property Let BackColor(New_Val As OLE_COLOR)
    Picture2.BackColor = New_Val
    PropertyChanged "BackColor"
    DrawRuler
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Picture2.ForeColor
End Property

Public Property Let ForeColor(New_Val As OLE_COLOR)
    Picture2.ForeColor = New_Val
    PropertyChanged "ForeColor"
    DrawRuler
End Property

Private Sub DrawRuler()
    Dim Sincr As Single
    'Scalemode is in TWIPS 1440 per inch
    Dim I As Integer
    'Number of segment across form
    Sincr = RScale / 10
    With Picture2
        If m_Orientation = rlrHorizontal Then
            Do While Sincr < .ScaleWidth
                'Number of sections
                For I = 1 To 10
                    'Size of Tics
                    If I = 10 Then
                        Picture2.Line (Sincr, 0)-(Sincr, .ScaleHeight)
                        .CurrentY = 0
                        Picture2.Print " " + CStr(Int(Sincr / RScale))
                    ElseIf I = Int(10 * 0.5) Then
                        Picture2.Line (Sincr, .ScaleHeight - _
                        (.ScaleHeight * 0.5))-(Sincr, .ScaleHeight)
                    Else
                        Picture2.Line (Sincr, .ScaleHeight - _
                        (.ScaleHeight * 0.125))-(Sincr, .ScaleHeight)
                    End If
                    Sincr = Sincr + (RScale / 10)
                Next
            Loop
        Else
            Do While Sincr < .ScaleHeight
                'Number of sections
                For I = 1 To 10
                    'Size of Tics
                    If I = 10 Then
                        'Einheiten schreiben
                        Picture2.Line (0, Sincr)-(.ScaleHeight, Sincr)
                        .CurrentX = 0
                        Picture2.Print CStr(Int(Sincr / RScale))
                    ElseIf I = Int(10 * 0.5) Then
                        '50%
                        Picture2.Line (.ScaleWidth - _
                        (.ScaleWidth * 0.5), Sincr)-(.ScaleWidth, Sincr)
                    Else
                        Picture2.Line (.ScaleWidth - _
                        (.ScaleWidth * 0.125), Sincr)-(.ScaleWidth, Sincr)
                    End If
                    Sincr = Sincr + (RScale / 10)
                Next
            Loop
        End If
    End With
End Sub

Public Sub MouseMoved(X As Single)
    
    With Picture2
        .DrawMode = 6
        If m_Orientation = rlrHorizontal Then
            Picture2.Line (X, 0)-(X, .ScaleHeight)
            If x1 > 0 Then
                Picture2.Line (x1, 0)-(x1, .ScaleHeight)
            End If
            x1 = X
        Else
            Picture2.Line (0, X)-(.ScaleWidth, X)
            If x1 > 0 Then
                Picture2.Line (0, x1)-(.ScaleWidth, x1)
            End If
            x1 = X
        End If
        .DrawMode = 13
    End With
End Sub

Private Sub mnuMenu_Click()
    Dim I As Integer
    For I = 0 To mnuMode.Count - 1
        mnuMode(I).Checked = False
    Next I
    mnuMode(m_Mode).Checked = True
End Sub

Private Sub mnuMode_Click(Index As Integer)
    RulerMode = Index
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then UserControl.PopupMenu mnuMenu
End Sub

Private Sub UserControl_Initialize()

    RScale = 570

End Sub

Private Sub UserControl_InitProperties()
    m_Orientation = m_Def_Orientation
    'm_BorderStyle = m_Def_BorderStyle
End Sub

Private Sub UserControl_Resize()
    Picture2.Height = UserControl.Height
    Picture2.Width = UserControl.Width
    UserControl.Cls
    x1 = 0
    Picture2.Cls
    DrawRuler
End Sub

Private Sub UserControl_Show()
    Picture2.Cls
    DrawRuler
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

        Orientation = PropBag.ReadProperty("Orientation", m_Def_Orientation)
        BorderStyle = PropBag.ReadProperty("BorderStyle", m_Def_BorderStyle)
        RulerMode = PropBag.ReadProperty("RulerMode", 0)
        Picture2.BackColor = PropBag.ReadProperty("BackColor", &H80000018)
        Picture2.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
        'Picture2.BorderStyle = m_BorderStyle
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

        Call PropBag.WriteProperty("Orientation", m_Orientation, m_Def_Orientation)
        Call PropBag.WriteProperty("BorderStyle", lBorderDef, m_Def_BorderStyle)
        Call PropBag.WriteProperty("RulerMode", m_Mode, 0)
        Call PropBag.WriteProperty("BackColor", Picture2.BackColor, &H80000018)
        Call PropBag.WriteProperty("ForeColor", Picture2.ForeColor, &H80000012)

End Sub

Private Sub ChangeOrientSizes()
    Dim rlrWidth As Long
    Dim rlrHeight As Long
    
    rlrWidth = UserControl.Width
    rlrHeight = UserControl.Height
    UserControl.Height = rlrWidth
    UserControl.Width = rlrHeight
    UserControl.Cls
    Picture2.Cls
    DrawRuler
End Sub
