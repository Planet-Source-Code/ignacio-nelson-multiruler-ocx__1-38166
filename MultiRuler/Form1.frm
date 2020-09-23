VERSION 5.00
Object = "*\AMultiRuler.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form AlignTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ruler OCX Test Form"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   FillColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin vb6MultiRuler.MultiRuler MultiRulerTXVer 
      Height          =   2895
      Left            =   600
      TabIndex        =   13
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5106
      Orientation     =   1
      BorderStyle     =   6
      BackColor       =   16777215
   End
   Begin vb6MultiRuler.MultiRuler MultirulerTXHor 
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   3000
      Width           =   5760
      _ExtentX        =   5080
      _ExtentY        =   661
      BorderStyle     =   6
      BackColor       =   16777215
   End
   Begin vb6MultiRuler.MultiRuler MultiRuler5 
      Align           =   3  'Align Left
      Height          =   6135
      Left            =   0
      TabIndex        =   3
      Top             =   345
      Width           =   375
      _ExtentX        =   10425
      _ExtentY        =   661
      Orientation     =   1
      BorderStyle     =   5
      BackColor       =   33023
      ForeColor       =   0
   End
   Begin vb6MultiRuler.MultiRuler MultiRuler4 
      Align           =   4  'Align Right
      Height          =   6135
      Left            =   6945
      TabIndex        =   2
      Top             =   345
      Width           =   360
      _ExtentX        =   10425
      _ExtentY        =   635
      Orientation     =   1
      BorderStyle     =   1
      BackColor       =   12648447
   End
   Begin vb6MultiRuler.MultiRuler MultiRuler3 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6480
      Width           =   7305
      _ExtentX        =   661
      _ExtentY        =   12277
      BorderStyle     =   3
      BackColor       =   0
      ForeColor       =   16777215
   End
   Begin vb6MultiRuler.MultiRuler MultiRuler1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   609
      BorderStyle     =   6
      BackColor       =   -2147483633
      ForeColor       =   0
   End
   Begin vb6MultiRuler.MultiRuler MultiRuler2 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   661
      BorderStyle     =   0
      BackColor       =   16777215
      ForeColor       =   0
   End
   Begin vb6MultiRuler.MultiRuler MultiRuler7 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   661
      BorderStyle     =   0
      RulerMode       =   1
      BackColor       =   16777215
      ForeColor       =   0
   End
   Begin vb6MultiRuler.MultiRuler MultiRuler8 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1680
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   661
      BorderStyle     =   0
      RulerMode       =   2
      BackColor       =   16777215
      ForeColor       =   0
   End
   Begin vb6MultiRuler.MultiRuler MultiRuler9 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   661
      BorderStyle     =   0
      RulerMode       =   3
      BackColor       =   16777215
      ForeColor       =   0
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2895
      Left            =   960
      TabIndex        =   14
      Top             =   3360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":038A
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Twips"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   11
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   9
      Top             =   840
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   1320
      Width           =   375
   End
End
Attribute VB_Name = "AlignTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MultirulerTXHor.MouseMoved X
    MultiRulerTXVer.MouseMoved Y
End Sub
