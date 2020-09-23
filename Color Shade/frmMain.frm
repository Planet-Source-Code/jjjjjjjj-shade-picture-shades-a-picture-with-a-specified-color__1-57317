VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Color  Shading"
   ClientHeight    =   2760
   ClientLeft      =   75
   ClientTop       =   330
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   2760
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Processed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   3600
      ScaleHeight     =   960
      ScaleWidth      =   2100
      TabIndex        =   6
      Top             =   1080
      Width           =   2100
   End
   Begin VB.HScrollBar scrRate 
      Height          =   255
      LargeChange     =   2
      Left            =   600
      Max             =   10
      TabIndex        =   2
      Top             =   2350
      Value           =   4
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Original 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   3600
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   900
      ScaleWidth      =   2100
      TabIndex        =   1
      Top             =   120
      Width           =   2100
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1860
      Left            =   120
      Picture         =   "frmMain.frx":14BD
      ScaleHeight     =   1800
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   120
      Width           =   3405
   End
   Begin VB.Label lbThick 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   240
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   240
      Left            =   3120
      TabIndex        =   5
      Top             =   2350
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2350
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shading Thickness"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1545
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Original_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox Original.Point(X, Y)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShadePicture Original, Processed, Picture1.Point(X, Y), scrRate.Value
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Clipboard.SetText ("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57317&lngWId=1")
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "?") = vbYes Then
        MsgBox "Please Rate my code,The site address is already copied to your clipboard", vbInformation, "ThankYou"
    Else
        MsgBox "Please give FeedBack,The site address is already copied to your clipboard", vbInformation, "Please Give FeedBack"
    End If
End Sub
Private Sub scrRate_Change()
    lbThick = scrRate.Value
End Sub
