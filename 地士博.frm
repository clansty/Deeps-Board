VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "地士博"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   Icon            =   "地士博.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "×"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   11280
      Width           =   255
   End
   Begin VB.PictureBox nopic 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox chosen 
      Height          =   255
      Left            =   0
      Picture         =   "地士博.frx":5000A
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   12
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
      Left            =   3360
      ScaleHeight     =   15562
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   823
      TabIndex        =   11
      Top             =   -240
      Width           =   12375
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H008080FF&
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FF80FF&
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "－"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "＋"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "下"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "擦"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   7800
      Width           =   135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   7320
      Width           =   135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   6360
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   -100
      TabIndex        =   18
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   -100
      TabIndex        =   17
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   -100
      TabIndex        =   16
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   -100
      TabIndex        =   15
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   20450
      TabIndex        =   24
      Top             =   7320
      Width           =   300
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   20460
      TabIndex        =   25
      Top             =   7800
      Width           =   300
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   20040
      TabIndex        =   26
      Top             =   6840
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim dsbMode As Integer
Private Sub Command1_Click()
    Picture1.DrawWidth = 20
    Picture1.ForeColor = Picture1.BackColor

End Sub

Private Sub Command2_Click()
    Picture1.Top = Picture1.Top - 2333
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

Private Sub Command7_Click()
    ShowCursor 1
    End
End Sub

Private Sub Command9_Click()
    Picture1.DrawWidth = 20
    Picture1.ForeColor = Picture1.BackColor

End Sub

Private Sub Form_Load()
    Debug.Print Form1.Width
    Picture1.Left = 20
    Picture1.Width = 20450
    ShowCursor 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ShowCursor 1
End Sub

Private Sub Label10_Click()
    Picture1.Top = Picture1.Top + 2333

End Sub

Private Sub Label11_Click()
    Picture1.Top = Picture1.Top - 2333

End Sub

Private Sub Label12_Click()
    Picture1.Top = Picture1.Top + 2333

End Sub

Private Sub Label13_Click()
    dsbMode = 2
End Sub

Private Sub Label2_Click()
    Picture1.ForeColor = vbWhite
    Picture1.DrawWidth = Label1.Caption
    
    dsbMode = 1

End Sub

Private Sub Label3_Click()
    Picture1.ForeColor = Picture3.BackColor
    Picture1.DrawWidth = Label1.Caption
    dsbMode = 1


End Sub

Private Sub Label4_Click()
    Picture1.ForeColor = Picture4.BackColor
    Picture1.DrawWidth = Label1.Caption
    dsbMode = 1


End Sub

Private Sub Label5_Click()
    Picture1.ForeColor = Picture5.BackColor
    Picture1.DrawWidth = Label1.Caption
    dsbMode = 1

End Sub

Private Sub Label6_Click()
    Picture1.ForeColor = Picture6.BackColor
    Picture1.DrawWidth = Label1.Caption
    dsbMode = 1

End Sub

Private Sub Label7_Click()
    Picture1.ForeColor = Picture7.BackColor
    Picture1.DrawWidth = Label1.Caption
    dsbMode = 1

End Sub

Private Sub Label8_Click()
    Picture1.DrawWidth = 20
    Picture1.ForeColor = Picture1.BackColor
    dsbMode = 1

End Sub

Private Sub Label9_Click()
    Picture1.Top = Picture1.Top - 2333

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Picture1.CurrentX = X
        Picture1.CurrentY = Y
        If dsbMode = 2 Then
            Picture1.Line (X - 200, Y)-(X + 200, Y)
            Picture1.Line (X, Y - 200)-(X, Y + 200)
            dsbMode = 1
        End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 1 Then Picture1.Line -(X, Y)
End Sub
