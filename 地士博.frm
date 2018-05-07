VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Deep 士博"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   Icon            =   "地士博.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   3
      ForeColor       =   &H80000008&
      Height          =   2.33460e5
      Left            =   2520
      ScaleHeight     =   15564
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Height          =   1335
         Left            =   3000
         Top             =   1080
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Visible         =   0   'False
         X1              =   136
         X2              =   400
         Y1              =   208
         Y2              =   312
      End
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000007&
      Caption         =   "Label15"
      Height          =   375
      Left            =   20280
      TabIndex        =   15
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   20400
      TabIndex        =   14
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   -120
      TabIndex        =   13
      Top             =   11160
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7800
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   7320
      Width           =   135
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   6360
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   -105
      TabIndex        =   1
      Top             =   3960
      Width           =   240
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   20325
      TabIndex        =   10
      Top             =   7320
      Width           =   420
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   20340
      TabIndex        =   11
      Top             =   7800
      Width           =   420
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   20280
      TabIndex        =   12
      Top             =   6840
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim dsbMode As Integer, colorID As Integer
Private Sub chColor(tmp As Integer)
    Picture1.DrawWidth = 4
    Select Case tmp
        Case 1
            Picture1.ForeColor = QBColor(15)
        Case 2
            Picture1.ForeColor = QBColor(11)
        Case 3
            Picture1.ForeColor = QBColor(14)
        Case 4
            Picture1.ForeColor = QBColor(10)
        Case 5
            Picture1.ForeColor = QBColor(13)
        Case 6
            Picture1.ForeColor = QBColor(12)
    End Select
    colorID = tmp
    dsbMode = 1
End Sub
Private Sub chEraser()
    dsbMode = 2
    Picture1.ForeColor = QBColor(0)
    Picture1.DrawWidth = 20
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            Picture1.Top = Picture1.Top - 2333
        Case 40
            Picture1.Top = Picture1.Top + 2333
        Case 9
            If dsbMode = 1 Then chEraser Else chColor colorID
    End Select

End Sub

Private Sub Form_Load()
    'Debug.Print Form1.Width
    Picture1.Left = 20
    Picture1.Width = 20450
    ShowCursor 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowCursor 1
End Sub

Private Sub Label1_Click()
    If dsbMode = 1 Then dsbMode = 4 Else MsgBox ("请先选择颜色")
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
    If dsbMode = 1 Then dsbMode = 3 Else MsgBox ("请先选择颜色")
End Sub

Private Sub Label14_Click()
    ShowCursor 1
    End
End Sub

Private Sub Label15_Click()
    If dsbMode = 1 Then dsbMode = 5 Else MsgBox ("请先选择颜色")
End Sub

Private Sub Label2_Click()
    chColor 1
End Sub

Private Sub Label3_Click()
    chColor 2
End Sub

Private Sub Label4_Click()
    chColor 3
End Sub

Private Sub Label5_Click()
    chColor 4
End Sub

Private Sub Label6_Click()
    chColor 5
End Sub

Private Sub Label7_Click()
    chColor 6
End Sub

Private Sub Label8_Click()
    chEraser
End Sub

Private Sub Label9_Click()
    Picture1.Top = Picture1.Top - 2333

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.CurrentX = X
    Picture1.CurrentY = Y
    Select Case dsbMode
        Case 3
            Picture1.Line (X - 200, Y)-(X + 200, Y)
            Picture1.Line (X, Y - 200)-(X, Y + 200)
            dsbMode = 1
        Case 4
            Line1.Visible = True
            Line1.BorderColor = Picture1.ForeColor
            Line1.X1 = X
            Line1.Y1 = Y
            Line1.Y2 = Y
            Line1.X2 = X
        Case 5
            Line1.Visible = True
            Line1.BorderColor = Picture1.ForeColor
            Line1.X1 = X
            Line1.Y1 = Y
            Line1.Y2 = Y
            Line1.X2 = X
    End Select
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Select Case dsbMode
            Case 1
                Picture1.Line -(X, Y)
            Case 2
                Picture1.Line -(X, Y)
            Case 4
                Line1.X2 = X
                Line1.Y2 = Y
            Case 5
                Line1.X2 = X
                Line1.Y2 = Y
        End Select
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case dsbMode
        Case 4
            Picture1.Line -(X, Y)
            Line1.Visible = False
        Case 5
            Picture1.Line -(X, Y), , B
            Line1.Visible = False
    End Select
End Sub
