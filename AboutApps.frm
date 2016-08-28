VERSION 5.00
Begin VB.Form AboutApps 
   Appearance      =   0  'Flat
   BackColor       =   &H00400040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   Icon            =   "AboutApps.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "AboutApps.frx":08CA
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Intensity Care Of Mother and Kids"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Registered Kid and mother in the Apps"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   9600
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Kids:-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   10920
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   10920
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Mother:-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   10320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   10320
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   120
      Picture         =   "AboutApps.frx":0AB6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "AboutApps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

If lalit = 5 Then
      lalit = 0
      Else
        lalit = lalit + 1
        End If
Me.BackColor = colorlist(lalit)
Text1.BackColor = colorlist(lalit)


Me.Top = 0
Me.Left = MainForm.ScaleWidth - 3200
On Error GoTo kjas
rs1.Open "select * from mothertable where agganid='" & ArgNum & "'", con, adOpenStatic, adLockReadOnly
Dim ttt
 ttt = 0
  rs1.MoveFirst
  While Not rs1.EOF
        ttt = ttt + 1
        rs1.MoveNext
              Wend
       
rs1.Close
Label1.Caption = ttt
rs1.Open "select * from kidtable where agganid='" & ArgNum & "'", con, adOpenStatic, adLockReadOnly

 ttt = 0
  rs1.MoveFirst
  While Not rs1.EOF
        ttt = ttt + 1
        rs1.MoveNext
              Wend
       
rs1.Close
Label3.Caption = ttt
Exit Sub
kjas:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lk.Left = 0
lk.Top = MainForm.ScaleHeight - 1560
End Sub
