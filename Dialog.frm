VERSION 5.00
Begin VB.Form Dialog 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   7920
   ClientLeft      =   2715
   ClientTop       =   3345
   ClientWidth     =   15825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   15825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Label okButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Revel Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   11160
      TabIndex        =   2
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   5880
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   6360
      TabIndex        =   0
      Top             =   6000
      Width           =   2055
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
If lalit = 5 Then
      lalit = 0
      Else
        lalit = lalit + 1
        End If
Me.BackColor = colorlist(lalit)
Label2.Caption = fegpass
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
okButton.BorderStyle = 0
End Sub

Private Sub OKButton_Click()
Unload Me
login_form.Show
End Sub

Private Sub okButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
okButton.BorderStyle = 1
End Sub
