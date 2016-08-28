VERSION 5.00
Begin VB.Form msgShow 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17910
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   17910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   600
      Top             =   5040
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   570
      Left            =   8280
      TabIndex        =   1
      Top             =   6360
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "msgShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If india = 0 Then
   MainForm.Hide
   End If

If lalit = 5 Then
      lalit = 0
      Else
        lalit = lalit + 1
        End If
msgShow.BackColor = colorlist(lalit)


lbl.Caption = msg
End Sub

Private Sub Timer1_Timer()
   
   
   Unload Me
   If india = 0 Then
   MainForm.Show
   End If
End Sub
