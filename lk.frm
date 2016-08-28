VERSION 5.00
Begin VB.Form lk 
   Appearance      =   0  'Flat
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   16000
      Left            =   480
      Top             =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"lk.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1335
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   15720
      Picture         =   "lk.frx":00DE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1365
      Index           =   6
      Left            =   1920
      Picture         =   "lk.frx":D918
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1365
      Index           =   5
      Left            =   1920
      Picture         =   "lk.frx":10B90
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1365
      Index           =   4
      Left            =   1920
      Picture         =   "lk.frx":124D7
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1365
      Index           =   3
      Left            =   1920
      Picture         =   "lk.frx":35799
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1365
      Index           =   2
      Left            =   1920
      Picture         =   "lk.frx":5A7BF
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1365
      Index           =   1
      Left            =   1920
      Picture         =   "lk.frx":76B69
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   1365
      Index           =   0
      Left            =   1920
      Picture         =   "lk.frx":99E2B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "lk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lmsg(7) As String
Dim lcolor(7)
Dim ind

Private Sub Form_Load()
Me.Width = MainForm.Width
Me.Left = 0
Me.Top = MainForm.ScaleHeight - 1830
lmsg(0) = "There are many reasons to love India the most important being India is our motherland, our country. We all are in a way emotionally attached to our Mother Nation. There are a lot many other reasons to love the country."
lcolor(0) = &H4000&
lmsg(1) = "There is no velvet so soft as a mother’s lap,no rose as lovely as her smile,no path so flowery as that imprinted with her footsteps. Happy mother day ,But mothers 's day for me is everyday LOVE YOU MOM,Whis you all the best"
lcolor(1) = &H4040&
lmsg(2) = "Many Hugs Only Luv Never Anger Teaching Me Helping Me Every Smile When I Was Sad Raising Me To Be Strong It Spells Mother. Thanks For Being "
lcolor(2) = &H40&
lmsg(3) = "A Mother Serves Her Sugar With A Bit Of Peppermint To Clarify The Passages That Carry What She Meant When She First Set To Bear A Soul Quite Separat"
lcolor(3) = &H404000
lmsg(4) = "Mom...I love you lot.. I have learnt so many things from you..you are my first teacher. From you only,i know how to talk & behave with others. You grew up me in Godly fear. My positin came from you only."
lcolor(4) = &H400000
lmsg(5) = "Most Superior Example 4 Love' In Anyones Life:- When Apples Are Four And Members Are Five In A Family Then Mother Says, I Dont Like Apples."
lcolor(5) = &H400040
lmsg(6) = "When U Are A Mother, U R Never Really Alone In Ur Thoughts. A Mother Always Has To Think Twice, Once For Herself And Once For Her Child"
lcolor(6) = &H40C0&
ind = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Unload AboutApps
End Sub





Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload AboutApps
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload AboutApps
End Sub

Private Sub Timer2_Timer()
If ind = 7 Then
    ind = 0
    End If
If ind = 0 Then
      Image1(ind).Enabled = True
      Image1(ind).Visible = True
      Image1(6).Enabled = False
      Image1(6).Visible = False
      Me.BackColor = lcolor(ind)
      Label1.Caption = lmsg(ind)
      ind = ind + 1
Else
     Image1(ind).Enabled = True
      Image1(ind).Visible = True
      Image1(ind - 1).Enabled = False
      Image1(ind - 1).Visible = False
      Me.BackColor = lcolor(ind)
      Label1.Caption = lmsg(ind)
      ind = ind + 1
End If
End Sub
