VERSION 5.00
Begin VB.Form abtu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18810
   Icon            =   "abtu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   18810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "abtu.frx":08CA
      Top             =   240
      Width           =   12735
   End
   Begin VB.Image lalit 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2325
      Left            =   13080
      Picture         =   "abtu.frx":0C8E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile - +918386814144"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   23
      Left            =   15240
      TabIndex        =   16
      Top             =   2280
      Width           =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email - lalitvchau@outlook.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   22
      Left            =   15240
      TabIndex        =   15
      Top             =   1920
      Width           =   2865
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S.P.U.(P.G.) College"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   21
      Left            =   15240
      TabIndex        =   14
      Top             =   1440
      Width           =   2040
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BS-IT Part 3 rd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   20
      Left            =   15240
      TabIndex        =   13
      Top             =   960
      Width           =   1710
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LaLit Kumar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   555
      Index           =   19
      Left            =   15240
      TabIndex        =   12
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Deepak 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2325
      Left            =   13080
      Picture         =   "abtu.frx":61E5
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deepak Kumar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   555
      Index           =   0
      Left            =   15240
      TabIndex        =   11
      Top             =   3120
      Width           =   2805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile - +919772703455"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Index           =   1
      Left            =   15240
      TabIndex        =   10
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email - dvaishanav750@gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   15240
      TabIndex        =   9
      Top             =   4800
      Width           =   2925
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S.P.U.(P.G.) College"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   3
      Left            =   15240
      TabIndex        =   8
      Top             =   4320
      Width           =   2040
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BS-IT Part 3 rd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   4
      Left            =   15240
      TabIndex        =   7
      Top             =   3840
      Width           =   1710
   End
   Begin VB.Image ravi 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2325
      Left            =   13080
      Picture         =   "abtu.frx":E58DB
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ravindra Singh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   555
      Index           =   5
      Left            =   15240
      TabIndex        =   6
      Top             =   5880
      Width           =   2805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile - +917891979901"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   6
      Left            =   15240
      TabIndex        =   5
      Top             =   7920
      Width           =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email - raoravi336@gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   15240
      TabIndex        =   4
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S.P.U.(P.G.) College"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   8
      Left            =   15240
      TabIndex        =   3
      Top             =   7080
      Width           =   2040
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BS-IT Part 3 rd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   9
      Left            =   15240
      TabIndex        =   2
      Top             =   6600
      Width           =   1710
   End
   Begin VB.Label lts 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   13200
      TabIndex        =   1
      Top             =   8400
      Width           =   5415
   End
End
Attribute VB_Name = "abtu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lts.ForeColor = vbBlack
lts.BackColor = &H4000&
lts.BorderStyle = 0

End Sub

Private Sub lts_Click()
lts.BackColor = &H80&
 lts.ForeColor = &HFFC0C0
lts.BorderStyle = 1
End
End Sub

Private Sub lts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lts.ForeColor = &HFFC0C0
lts.BorderStyle = 1
End Sub


