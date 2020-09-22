VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   222
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox b1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   825
      Left            =   1800
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   4
      Top             =   240
      Width           =   1305
   End
   Begin VB.PictureBox b2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   825
      Left            =   360
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   3
      Top             =   1800
      Width           =   825
   End
   Begin VB.PictureBox b3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   465
      Left            =   1680
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   2
      Top             =   1320
      Width           =   1185
   End
   Begin VB.PictureBox b5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   705
      Left            =   2040
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   1
      Top             =   1920
      Width           =   825
   End
   Begin VB.PictureBox b4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Left            =   240
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Button1 As New clsButton
Dim Button2 As New clsButton
Dim Button3 As New clsButton
Dim Button4 As New clsButton
Dim Button5 As New clsButton


Private Sub Command1_Click()
dT = Timer()
Button1.InitButton b1, "ONE", False, 0.5, 10, 10, True
Button2.InitButton b2, "TWO", True, 0.75, 10, 10, False
Button3.InitButton b3, "THREE", True, 0.5, 5, 5, True
Button4.InitButton b4, "FOUR", False, 0.5, 12, 20, False
Button5.InitButton b5, "FIVE", True, 0.75, 15, 5, True
Me.Caption = Timer - dT
End Sub

Private Sub Command2_Click()
Me.Caption = Button5.InitButton(b5, "FOUR", True, 0.5, 12, 20, False)
End Sub


Private Sub b1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button1.TriggerButton
End Sub


Private Sub b2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button2.TriggerButton
End Sub


Private Sub b3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button3.TriggerButton
End Sub


Private Sub b4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button4.TriggerButton
End Sub


Private Sub b5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button5.TriggerButton
End Sub


Private Sub Form_Load()
Button1.InitButton b1, "ONE", False, 0.75, 15, 0, True
Button2.InitButton b2, "TWO", True, 0.75, 10, 10, False
Button3.InitButton b3, "THREE", True, 0.75, 10, 5, True
Button4.InitButton b4, "FOUR", True, 0.75, 12, 12, False
Button5.InitButton b5, "FIVE", False, 0.75, 5, 5, True
End Sub












