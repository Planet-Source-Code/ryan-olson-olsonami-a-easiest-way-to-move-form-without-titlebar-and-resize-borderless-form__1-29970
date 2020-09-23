VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   375
      Left            =   4080
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   4650
      TabIndex        =   2
      Top             =   0
      Width           =   4680
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   0
         MouseIcon       =   "Form1.frx":0528
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0832
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Height          =   2715
      Left            =   4575
      MouseIcon       =   "Form1.frx":0B3C
      MousePointer    =   99  'Custom
      ScaleHeight     =   2655
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   375
      Width           =   100
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   100
      Left            =   0
      MouseIcon       =   "Form1.frx":0E46
      MousePointer    =   99  'Custom
      ScaleHeight     =   45
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   3090
      Width           =   4680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim n As Integer
Dim l As Integer
Dim p As Integer

Private Sub Form_Resize()
On Error Resume Next
Picture5.Left = Width - Picture5.Width - Picture2.Width
Picture5.Top = Height - Picture5.Height - Picture1.Height
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
n = 1
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If n = 1 Then
Me.Height = Y + Me.Height
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
n = 0
End Sub




Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = 1
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If i = 1 Then
Me.Width = X + Me.Width
End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = 0
End Sub



Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
l = 1
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If l = 1 Then
Me.Left = X + Me.Left
Me.Top = Y + Me.Top
End If
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
l = 0
End Sub



Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
p = 1
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If p = 1 Then
Width = X + Width
Height = Y + Height
End If
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
p = 0
End Sub
