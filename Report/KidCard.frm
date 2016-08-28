VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form KidCard 
   Appearance      =   0  'Flat
   BackColor       =   &H00800080&
   BorderStyle     =   0  'None
   Caption         =   "wari "
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17625
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleWidth      =   17625
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Copies          =   1000
      DefaultExt      =   "*.jpg"
      DialogTitle     =   "Print Kid Card"
      FileName        =   "KidCard"
   End
   Begin VB.ComboBox kidregno 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   7680
      TabIndex        =   25
      Text            =   "Select Kid's Regstration No"
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Thnaks You !"
      BeginProperty Font 
         Name            =   "Revel Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   26
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Revel Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      TabIndex        =   27
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label anm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   10440
      TabIndex        =   24
      Top             =   8640
      Width           =   7065
   End
   Begin VB.Label asha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   3840
      TabIndex        =   23
      Top             =   8640
      Width           =   5505
   End
   Begin VB.Label owner 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   5760
      TabIndex        =   22
      Top             =   7680
      Width           =   11745
   End
   Begin VB.Label aganid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   13200
      TabIndex        =   21
      Top             =   6720
      Width           =   4305
   End
   Begin VB.Label no 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   13200
      TabIndex        =   20
      Top             =   5880
      Width           =   4305
   End
   Begin VB.Label weight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   13200
      TabIndex        =   19
      Top             =   4920
      Width           =   4305
   End
   Begin VB.Label center 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   5760
      TabIndex        =   18
      Top             =   6720
      Width           =   3585
   End
   Begin VB.Label gender 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   5760
      TabIndex        =   17
      Top             =   5880
      Width           =   3345
   End
   Begin VB.Label birthdate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   5760
      TabIndex        =   16
      Top             =   4920
      Width           =   3345
   End
   Begin VB.Label fathername 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   7920
      TabIndex        =   15
      Top             =   3960
      Width           =   9465
   End
   Begin VB.Label mothername 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   7920
      TabIndex        =   14
      Top             =   3120
      Width           =   9465
   End
   Begin VB.Label kidname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   7920
      TabIndex        =   13
      Top             =   2400
      Width           =   9465
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "KID's CARD"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   660
      Left            =   2760
      TabIndex        =   12
      Top             =   1200
      Width           =   3285
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ANM"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   9480
      TabIndex        =   11
      Top             =   8640
      Width           =   825
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Asha"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   2880
      TabIndex        =   10
      Top             =   8640
      Width           =   825
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aanganwadi Owner"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   2655
      TabIndex        =   9
      Top             =   7680
      Width           =   3225
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aanganwadi Center"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   2670
      TabIndex        =   8
      Top             =   6720
      Width           =   3195
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aanganwadi Reg No"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   9390
      TabIndex        =   7
      Top             =   6720
      Width           =   3225
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kid's Registration No."
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   9480
      TabIndex        =   6
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   2880
      TabIndex        =   5
      Top             =   5880
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Weight at Born"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   9480
      TabIndex        =   4
      Top             =   4920
      Width           =   2445
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   2880
      TabIndex        =   3
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   5280
      TabIndex        =   2
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mother's Name"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   5280
      TabIndex        =   1
      Top             =   3120
      Width           =   2385
   End
   Begin VB.Label df 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kid's Name"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   5280
      TabIndex        =   0
      Top             =   2400
      Width           =   1785
   End
   Begin VB.Image photo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
End
Attribute VB_Name = "KidCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
 kidregno.Visible = False

 Command2.Visible = False
 
CommonDialog1.CancelError = True
On Error GoTo err
   CommonDialog1.Flags = cdlCFBoth
   CommonDialog1.ShowPrinter
   Exit Sub
err:
  
  


 Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.ForeColor = &HFFFF&
Command2.ForeColor = &HFFFF&
Command2.BorderStyle = 0
Command1.BorderStyle = 0
End Sub
Private Sub Command1_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command1.BorderStyle = 1
Command1.ForeColor = vbWhite
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command2.BorderStyle = 1
Command2.ForeColor = vbWhite
End Sub


Private Sub Form_Load()
If lalit = 5 Then
      lalit = 0
      Else
        lalit = lalit + 1
        End If
Me.BackColor = colorlist(lalit)
rs.Open "select * from kidtable,mothertable where mothertable.agganid='" & ArgNum & "'and kidtable.coupleno=mothertable.coupleno ", conn, adOpenStatic, adLockReadOnly
           
     While Not rs.EOF
        kidregno.AddItem rs.Fields("kidregno")
        rs.MoveNext
              Wend
        rs.Close
End Sub

Private Sub kidregno_Click()
    rs.Open "select * from kidtable where kidregno='" & kidregno.Text & "' ", conn, adOpenStatic, adLockReadOnly
    
    kidname.Caption = rs.Fields("kidname")
    mothername.Caption = rs.Fields("mothername")
    fathername.Caption = rs.Fields("fathername")
    photo = LoadPicture(rs.Fields("photo"))
    no.Caption = rs.Fields("kidregno")
    birthdate.Caption = Format(rs.Fields("birthdate"), "D-DD-MMM-YYYY")
    weight.Caption = rs.Fields("kidweight")
    weight.Caption = weight.Caption + " KG"
    gender.Caption = rs.Fields("gender")
    rs.Close
     rs.Open "select * from agganwaritable where agganid='" & ArgNum & "' ", conn, adOpenStatic, adLockReadOnly
    aganid.Caption = rs.Fields("agganregno")
   asha.Caption = rs.Fields("asha")
   anm.Caption = rs.Fields("anm")
      owner.Caption = rs.Fields("owner")
   center.Caption = rs.Fields("center")
    
    rs.Close
    Command2.Visible = True
  
    Command2.Enabled = True
   
End Sub
