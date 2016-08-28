VERSION 5.00
Begin VB.Form forgetPassword 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   Icon            =   "forgetPassword.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtRgNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      ToolTipText     =   "Enter Agganwari Registration Number"
      Top             =   5400
      Width           =   4575
   End
   Begin VB.TextBox aganID 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      ToolTipText     =   "Enter Your Agganwari ID number"
      Top             =   4920
      Width           =   4575
   End
   Begin VB.Label ok 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   9720
      TabIndex        =   5
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label cancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   11880
      TabIndex        =   6
      Top             =   6120
      Width           =   1935
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
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   9135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Aangan Reg. No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label LblAggID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "Aanganwadi ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   4920
      Width           =   2415
   End
End
Attribute VB_Name = "forgetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pass
Dim rgn
Dim temp
Dim id
Private Sub cancel_Click()
Unload Me
login_form.Show
End Sub

Private Sub cancel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BackStyle = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BackStyle = 0
ok.BackStyle = 0
End Sub

Private Sub ok_Click()
temp = txtRgNO.Text
rs.Open "select * from agganwaritable where agganid='" & aganid.Text & "'", conn, adOpenStatic, adLockReadOnly
'rs.Move
If rs.EOF Then
      rs.Close
         msg = "Not Found Any Account In This Application"
         msgShow.Show
         
     Else
     rgn = rs.Fields(1)

     If rgn = temp Then
     pass = rs.Fields(12)
     rs.Close
     fegpass = pass
     Unload Me
     Dialog.Show
     
Else
  msg = "Incorrect Aanganwadi  id and Registration Number"
  rs.Close
  msgShow.Show
 End If
  End If


End Sub


Private Sub ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ok.BackStyle = 1
End Sub

Private Sub txtRgNO_Change()
ok.Visible = True

End Sub

Private Sub txtRgNO_Validate(cancel As Boolean)
If IsNumeric(txtRgNO.Text) Then
    
    Else
    MsgBox "Enter a Number! AlphaNumeric or Alphbates are Not Allowed!"
    txtRgNO.Text = ""
    txtRgNO.SetFocus
    End If
End Sub
