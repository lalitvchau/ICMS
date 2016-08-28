VERSION 5.00
Begin VB.Form login_form 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Login To Kids"
   ClientHeight    =   9930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17115
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   17115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   840
      Top             =   6480
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   9480
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   4
      ToolTipText     =   "Enter Your Agganwari Password"
      Top             =   6000
      Width           =   6615
   End
   Begin VB.TextBox TxtAganID 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   540
      Left            =   9480
      MaxLength       =   30
      TabIndex        =   3
      ToolTipText     =   "Enter The Your Agganwari ID Number"
      Top             =   5160
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   3075
      Left            =   2160
      Picture         =   "Form1.frx":1982
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   3480
   End
   Begin VB.Label signIn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "SIGN - IN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   14160
      TabIndex        =   8
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12120
      TabIndex        =   7
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Intensity Care Of Mother and Kids"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   735
      Left            =   960
      TabIndex        =   6
      Top             =   2880
      Width           =   9135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Forget Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   10200
      TabIndex        =   5
      Top             =   7200
      Width           =   1770
   End
   Begin VB.Label acc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Need a new Account"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   7440
      TabIndex        =   2
      ToolTipText     =   "Click on text for create a new aganwari id"
      Top             =   7200
      Width           =   2340
   End
   Begin VB.Label password 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label aganID 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Aanganwadi ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
End
Attribute VB_Name = "login_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim passkey
Dim pass
Dim cnt As Integer
      
Dim id

Private Sub acc_Click()
Unload Me
accForm.Show
End Sub

Private Sub acc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 acc.FontSize = 11
 
End Sub

Private Sub Command1_Click()

End



End Sub



Private Sub Command1_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BorderStyle = 1
Command1.ForeColor = vbWhite
Command1.Enabled = True
End Sub

Private Sub Form_Load()
india = 1
cnt = 1
signIn.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = Me.BackColor
Command1.ForeColor = vbRed
Command1.BorderStyle = 0
signIn.BackColor = Me.BackColor
signIn.BorderStyle = 0
signIn.ForeColor = &HC000&
Command1.Enabled = True
acc.FontSize = 10
Label1.FontSize = 10
End Sub

Private Sub Label1_Click()
Unload Me
forgetPassword.Show
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontSize = 11
End Sub

Private Sub signIn_Click()
pass = TxtPassword.Text
id = TxtAganID.Text
rs.Open "select * from agganwaritable where agganid='" & TxtAganID.Text & "'", conn, adOpenStatic, adLockReadOnly
'rs.Move
Dim rec As Integer
 If id = Null Or pass = Null Then
       msg = "First Fill Aanganwadi  Id and password!"
       msgShow.Show
       rs.Close
  
  Else
     
    If rs.EOF Then
      rs.Close
         msg = "Not Found Any Account In This Application"
         msgShow.Show
     Else
            passkey = rs.Fields(12)
         If pass = passkey Then

              MainForm.Show
              lk.Show
    
             signIn.Visible = False
              ArgNum = id
              rs.Close
              Unload Me
        Else
         msg = "Wrong Password! Please Enter Correct password"
         msgShow.Show
         rs.Close
     End If
  End If
End If
End Sub


Private Sub signIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
signIn.BackColor = Me.BackColor
signIn.ForeColor = vbWhite
signIn.BorderStyle = 1
End Sub

Private Sub Timer1_Timer()
      
      login_form.BackColor = colorlist(cnt)
      Command1.BackColor = Me.BackColor
      signIn.BackColor = Me.BackColor

      If cnt = 5 Then
        cnt = 0
       Else
        cnt = cnt + 1
        End If
End Sub

Private Sub TxtPassword_Change()
signIn.Enabled = True
signIn.Visible = True

End Sub

Private Sub TxtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then signIn_Click
End Sub
