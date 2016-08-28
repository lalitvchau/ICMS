VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form updateKidReg 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "s"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   15075
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "updateKidReg.frx":0000
      Left            =   7320
      List            =   "updateKidReg.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2520
      Width           =   2175
   End
   Begin VB.ComboBox regno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7200
      TabIndex        =   19
      Text            =   "Select"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtKidWeight 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      ToolTipText     =   "Enter kids weight at birth time"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtBirthDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   15
      ToolTipText     =   "Enter baby birth date example- 02-Apr-2013"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.ComboBox CoupleId 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   7
      Text            =   "Select"
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtKidName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      ToolTipText     =   "Enter baby name"
      Top             =   1080
      Width           =   7215
   End
   Begin VB.TextBox txtImagePath 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   12120
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Upload Picture"
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
      Left            =   12120
      TabIndex        =   23
      Top             =   1920
      Width           =   2775
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
      Left            =   12600
      TabIndex        =   22
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label ok 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save"
      Enabled         =   0   'False
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
      Left            =   10440
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   " Kid Weight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Birth Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label txtFatherName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2040
      Width           =   7215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Father Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label txtMotherName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   1560
      Width           =   7215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   " Mother Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Kid's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Kids Registration No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label cp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Couple No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label kidImageName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Upload Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   12120
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.Image kidImage 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label agganId 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11400
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aanganwadi Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   9720
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "Update Kids Registration"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   20.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "updateKidReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancel_Click()
Unload Me

End Sub

Private Sub Command2_Click()

kidImageName.Caption = txtKidName.Text

Command2.Caption = "Update Picture"
CommonDialog1.CancelError = True
On Error GoTo err
   CommonDialog1.Flags = cdlCFBoth
   CommonDialog1.Filter = "(*.jpg)"
   CommonDialog1.FilterIndex = 2
   
   CommonDialog1.ShowOpen
   txtImagePath.Text = CommonDialog1.FileName
   kidImage.Picture = LoadPicture(txtImagePath.Text)
   kidImageName.Caption = txtKidName.Text
   Exit Sub
err:
MsgBox " Only jpeg Format Allow", vbExclamation
End Sub

Private Sub CoupleId_Click()
regno.Clear
rs.Open "select * from kidtable where coupleno='" & CoupleId.Text & "'", conn, adOpenStatic, adLockReadOnly
  If Not CoupleId.Text Then
            
     While Not rs.EOF
        regno.AddItem rs.Fields("kidregno")
        rs.MoveNext
              Wend
        rs.Close
     End If
End Sub



Private Sub CoupleId_KeyUp(KeyCode As Integer, Shift As Integer)
Dim tt
  tt = Val(CoupleId.Text)
  Dim ct
  ct = CoupleId.ListCount
  If KeyCode = 13 Then
   For i = 1 To ct
     If tt = CoupleId.List(i - 1) Then
       
       CoupleId_Click
       Exit For
       Else
        Me.Refresh
         CoupleId.Text = ""
         txtKidName.Text = ""
         txtMotherName.Caption = ""
         txtFatherName.Caption = ""
         txtBirthDate.Text = ""
         text1.Text = ""
         txtKidWeight.Text = ""
         kidImage.Picture = LoadPicture("")
         regno.Clear
         Me.Refresh
       End If
       
      Next
  End If
End Sub



Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackStyle = 1
Command2.BackColor = &H800000
Command2.ForeColor = vbWhite
End Sub

Private Sub ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ok.BackStyle = 1
ok.BackColor = &H800000
ok.ForeColor = vbWhite
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ok.BackStyle = 0

ok.ForeColor = &H80FF80
cancel.BackStyle = 0

cancel.ForeColor = &HFF&
Command2.BackStyle = 0
Command2.ForeColor = &H80FF80

End Sub
Private Sub cancel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BackStyle = 1
cancel.BackColor = &H800000
cancel.ForeColor = vbWhite
End Sub
Private Sub Form_Load()
Me.Left = 50
Me.Top = 50
agganId.Caption = ArgNum
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
        CoupleId.AddItem rs.Fields("coupleNo")
        rs.MoveNext
              Wend
        rs.Close
ok.Visible = False
End Sub

 

Private Sub ok_Click()
ok.Visible = False
kidImage.Picture = LoadPicture
conn.Execute "update  kidtable set kidname='" & txtKidName.Text & "',birthdate='" & txtBirthDate.Text & "',kidweight='" & txtKidWeight.Text & "',photo='" & txtImagePath.Text & "',gender= '" & text1.Text & "' where kidregno='" & regno.Text & "' "

 
msg = " Update Succefully!"
msgShow.Show

Unload Me

End Sub
Private Static Function nullCheck() As Boolean
Dim rt As Boolean
rt = False
 
     
If txtBirthDate.Text = "" Then
    MsgBox " Birth Date Is Blank !"
    txtBirthDate.SetFocus
    
ElseIf txtKidName.Text = "" Then
    MsgBox " Kid Name is  Blank !"
    txtKidName.SetFocus
    

ElseIf txtImagePath.Text = "" Then
    MsgBox " Photo Not Selected !"
    txtImagePath.SetFocus
    
ElseIf txtKidWeight.Text = "" Then
    MsgBox " Kid Is Blank !"
    txtKidWeight.SetFocus
ElseIf temp = "" Then
    MsgBox " Gender Is Blank !"
    
Else
   rt = True
    End If
nullCheck = rt
End Function
Private Sub regno_Click()
 ok.Visible = True
ok.Enabled = True
 If Not regno.Text Then
     rs.Open "select * from kidtable where kidregno='" & regno.Text & "' ", conn, adOpenStatic, adLockReadOnly
     txtKidName.Text = rs.Fields("kidname")
     txtMotherName.Caption = rs.Fields("mothername")
     txtFatherName.Caption = rs.Fields("fathername")
     txtBirthDate.Text = Format(rs.Fields("birthdate"), "DD-MMM-YYYY")
     txtImagePath.Text = rs.Fields("photo")
     txtKidWeight.Text = rs.Fields("kidweight")
       text1.Text = rs.Fields("gender")
       On Error GoTo err
       kidImage.Picture = LoadPicture(txtImagePath.Text)
       rs.Close
       Exit Sub
err: MsgBox "Photo Not Avialable!"
       rs.Close
     End If
     
End Sub

Private Sub regno_KeyUp(KeyCode As Integer, Shift As Integer)
Dim tt
  tt = Val(regno.Text)
  Dim ct
  ct = regno.ListCount
  If KeyCode = 13 Then
   For i = 1 To ct
     If tt = regno.List(i - 1) Then
       
       regno_Click
       Exit For
       Else
         Me.Refresh
         regno.Text = ""
         
         txtKidName.Text = ""
         txtMotherName.Caption = ""
         txtFatherName.Caption = ""
         txtBirthDate.Text = ""
         text1.Text = ""
         txtKidWeight.Text = ""
         kidImage.Picture = LoadPicture("")
         Me.Refresh
       End If
       
      Next
  End If
End Sub



Private Sub txtBirthDate_Validate(cancel As Boolean)
If IsDate(txtBirthDate.Text) Then
      txtBirthDate.Text = Format(txtBirthDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    txtBirthDate.Text = ""
    txtBirthDate.SetFocus
    End If
End Sub

Private Sub txtImagePath_Change()
Command2.Visible = True
End Sub
Private Sub txtKidWeight_Validate(cancel As Boolean)
If IsNumeric(txtKidWeight.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    txtKidWeight.Text = ""
    txtKidWeight.SetFocus
    End If
End Sub
