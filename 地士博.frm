VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "地士博 1.3.0 by 凌莞"
   ClientHeight    =   22905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   52245
   Icon            =   "地士博.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   22905
   ScaleWidth      =   52245
   WindowState     =   2  'Maximized
   Begin VB.PictureBox nopic 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox chosen 
      Height          =   255
      Left            =   0
      Picture         =   "地士博.frx":6988A
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      DrawWidth       =   3
      ForeColor       =   &H80000008&
      Height          =   2.33460e5
      Left            =   480
      ScaleHeight     =   15562
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1351
      TabIndex        =   13
      Top             =   0
      Width           =   20295
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FF80FF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "－"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "＋"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "下"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "擦"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   3960
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   21630
      Left            =   1200
      Picture         =   "地士博.frx":69D94
      Top             =   -2760
      Width           =   14565
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Sub clsChosen()
    Picture2.Picture = nopic.Picture
    Picture3.Picture = nopic.Picture
    Picture4.Picture = nopic.Picture
    Picture5.Picture = nopic.Picture
    Picture6.Picture = nopic.Picture
    Picture7.Picture = nopic.Picture
End Sub
Private Sub Command1_Click()
    Picture1.DrawWidth = 20
    Picture1.ForeColor = Picture1.BackColor
    clsChosen
End Sub

Private Sub Command2_Click()
    Picture1.Top = Picture1.Top - 2333
End Sub

Private Sub Command3_Click()
    Picture1.Cls
    Picture1.Top = 0
End Sub

Private Sub Command4_Click()
    Picture1.Top = Picture1.Top + 2333
End Sub

Private Sub Command5_Click()
    Label1.Caption = Label1.Caption + 2 - 1
    Picture1.DrawWidth = Label1.Caption
End Sub

Private Sub Command6_Click()
    Label1.Caption = Label1.Caption - 1
    If Label1.Caption < 1 Then
        Label1.Caption = 1
        frmAlart.Show
    End If
    Picture1.DrawWidth = Label1.Caption
End Sub


Private Sub Form_Load()
    ShowCursor 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ShowCursor 1
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.CurrentX = X
    Picture1.CurrentY = Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Picture1.Line -(X, Y)
End Sub

Private Sub Picture2_Click()
    Picture1.ForeColor = Picture2.BackColor
    Picture1.DrawWidth = Label1.Caption
    clsChosen
    Picture2.Picture = chosen.Picture
End Sub

Private Sub Picture3_Click()
    Picture1.ForeColor = Picture3.BackColor
    Picture1.DrawWidth = Label1.Caption
    clsChosen
    Picture3.Picture = chosen.Picture
End Sub

Private Sub Picture4_Click()
    Picture1.ForeColor = Picture4.BackColor
    Picture1.DrawWidth = Label1.Caption
    clsChosen
    Picture4.Picture = chosen.Picture
End Sub

Private Sub Picture5_Click()
    Picture1.ForeColor = Picture5.BackColor
    Picture1.DrawWidth = Label1.Caption
    clsChosen
    Picture5.Picture = chosen.Picture
End Sub

Private Sub Picture6_Click()
    Picture1.ForeColor = Picture6.BackColor
    Picture1.DrawWidth = Label1.Caption
    clsChosen
    Picture6.Picture = chosen.Picture
End Sub

Private Sub Picture7_Click()
    Picture1.ForeColor = Picture7.BackColor
    Picture1.DrawWidth = Label1.Caption
    clsChosen
    Picture7.Picture = chosen.Picture
End Sub
