VERSION 5.00
Begin VB.Form updateAccForm 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "updateAccForm.frx":0000
   ScaleHeight     =   6435
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox id 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   3120
      TabIndex        =   34
      ToolTipText     =   "Enter your Agganwari Center Registration no "
      Top             =   720
      Width           =   6375
   End
   Begin VB.TextBox txtOldPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   1920
      TabIndex        =   33
      ToolTipText     =   "Enter a New Password"
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox txtRePass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   6360
      TabIndex        =   31
      ToolTipText     =   "Enter return password"
      Top             =   5160
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   1920
      TabIndex        =   29
      ToolTipText     =   "Enter a New Password"
      Top             =   5160
      Width           =   2775
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   6360
      TabIndex        =   26
      ToolTipText     =   "Enter the Email ID"
      Top             =   4560
      Width           =   3135
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   25
      ToolTipText     =   "Enter hte Contect Number of Agganwari "
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox txthopital 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   3120
      TabIndex        =   23
      ToolTipText     =   "enter the Hospital and FRU"
      Top             =   4080
      Width           =   6375
   End
   Begin VB.TextBox TxtFHC 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   3120
      TabIndex        =   20
      ToolTipText     =   "Enter the First Health Center and City"
      Top             =   3600
      Width           =   6375
   End
   Begin VB.TextBox txtClinic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   3120
      TabIndex        =   18
      ToolTipText     =   "Enter thr upper health center or Clinic an"
      Top             =   3120
      Width           =   6375
   End
   Begin VB.TextBox txtANM 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   7200
      TabIndex        =   16
      ToolTipText     =   "Enter name of ANM The Agaanwari"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtAsha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   3120
      TabIndex        =   14
      ToolTipText     =   "Enter name of Asha of The Agganwari"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox TxtBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   7200
      TabIndex        =   11
      ToolTipText     =   "Enter agganwari center and Block"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtOwner 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   3120
      TabIndex        =   9
      ToolTipText     =   "Enter name of T Agganwari Owner Name"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtUppCenNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Enetr the Upper Center Number"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox TxtUppDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7200
      MaxLength       =   11
      TabIndex        =   7
      Text            =   "DD-MON-YYYY"
      ToolTipText     =   "Enter Upper Center  Date in DD-MON-YYYY format Example 05-APR-1993"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7200
      MaxLength       =   11
      TabIndex        =   4
      Text            =   "DD-MON-YYYY"
      ToolTipText     =   "Enter Registration Date in DD-MON-YYYY format Example 05-APR-1993"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtRgNO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   "Enter your Agganwari Center Registration no "
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label create 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   5160
      TabIndex        =   36
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label cancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "CANCEL"
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
      Left            =   7440
      TabIndex        =   35
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Old Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblRePass 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "RE-Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4800
      TabIndex        =   30
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label lblPas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4800
      TabIndex        =   27
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblMob 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Contect No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label hospital 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Hospital and FRU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Updates Aanganwadi Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   9735
   End
   Begin VB.Label lblFir 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "First Health Center / City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Lblclinic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Upper Health Center / Clinic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label lblANM 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "ANM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label LblAsha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Asha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label LblOwner 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Aanganwadi Owner Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Lblblock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Center/Block"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblUpDate 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblUppCenNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Upper Center No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblRgNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Aanganwadi Center Reg. No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label LblID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
End
Attribute VB_Name = "updateAccForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim passkey
Private Sub cancel_Click()
Unload Me

End Sub

Private Sub cancel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BorderStyle = 1
cancel.ForeColor = vbWhite
End Sub
Private Static Function nullCheck() As Boolean
Dim rt As Boolean
rt = False

If txtRgNO.Text = "" Then
    MsgBox " Aaganwadi Registration Number Is Blank !"
    txtRgNO.SetFocus
    
ElseIf txtDate.Text = "" Then
    MsgBox " Aaganwadi Registration Date Is Blank !"
    txtDate.SetFocus
    
ElseIf txtUppCenNo.Text = "" Then
    MsgBox " Aaganwadi Center  Number Is Blank !"
    txtUppCenNo.SetFocus
    
ElseIf TxtUppDate.Text = "" Then
    MsgBox " Aaganwadi Center  Date Is Blank !"
    TxtUppDate.SetFocus
    
ElseIf txtAsha.Text = "" Then
    MsgBox " Asha Name Is Blank !"
    txtAsha.SetFocus
    
ElseIf txtANM.Text = "" Then
    MsgBox " ANM Name Is Blank !"
    txtANM.SetFocus
    
ElseIf txtClinic.Text = "" Then
    MsgBox " Clinic Name Is Blank !"
    txtClinic.SetFocus
    
ElseIf TxtFHC.Text = "" Then
    MsgBox " FHC Name Is Blank !"
    TxtFHC.SetFocus
    
ElseIf txthopital.Text = "" Then
    MsgBox " Hospital Name Is Blank !"
    txthopital.SetFocus
    
ElseIf txtMobile.Text = "" Then
    MsgBox " Mobile Number Is Blank !"
    txtMobile.SetFocus
    
ElseIf txtEmail.Text = "" Then
    MsgBox " Email Is Blank !"
    txtEmail.SetFocus
    
ElseIf txtOwner.Text = "" Then
    MsgBox " Owner Is Blank !"
    txtOwner.SetFocus
    
ElseIf TxtBlock.Text = "" Then
    MsgBox " Block Is Blank !"
    TxtBlock.SetFocus
Else
   rt = True
    End If
nullCheck = rt
End Function

Private Sub create_Click()
If nullCheck = True Then
If txtOldPassword.Text = passkey Then
          If txtRePass.Text = "" And txtPass.Text = "" Then
                txtRePass.Text = "Army"
                conn.Execute " update agganwaritable set agganRegNo='" & txtRgNO.Text & "',rgDate='" & txtDate.Text & "',upperCenter='" & txtUppCenNo.Text & "',cendate='" & TxtUppDate.Text & "',asha='" & txtAsha.Text & "',anm='" & txtANM.Text & "',center='" & TxtBlock.Text & "',owner='" & txtOwner.Text & "',upperhc='" & txtClinic.Text & "',fhc='" & TxtFHC.Text & "',hosfru='" & txthopital.Text & "',contectno='" & txtMobile.Text & "',email='" & txtEmail.Text & "' where agganid='" & ArgNum & "'"
                Unload Me
                msg = "Upadate Your Account"
                msgShow.Show
                
                
          ElseIf txtRePass.Text = txtPass.Text Then
          
                 conn.Execute " update agganwaritable set agganRegNo='" & txtRgNO.Text & "',rgDate='" & txtDate.Text & "',upperCenter='" & txtUppCenNo.Text & "',cendate='" & TxtUppDate.Text & "',asha='" & txtAsha.Text & "',anm='" & txtANM.Text & "',center='" & TxtBlock.Text & "',owner='" & txtOwner.Text & "',upperhc='" & txtClinic.Text & "',fhc='" & TxtFHC.Text & "',hosfru='" & txthopital.Text & "',contectno='" & txtMobile.Text & "',email='" & txtEmail.Text & "',password='" & txtPass.Text & "' where agganid='" & ArgNum & "'"
                
                
                msg = "Upadate Your Account"
                msgShow.Show
                Unload Me
               Else
                 msg = "New Passwod and Re Password miss match! "
                 msgShow.Show
         End If
        Else
         msg = "Your Current Password Worng!"
         msgShow.Show
    End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
create.BorderStyle = 0
create.ForeColor = vbGreen
cancel.BorderStyle = 0
cancel.ForeColor = vbRed
End Sub
Private Sub create_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
create.BorderStyle = 1
create.ForeColor = vbWhite
End Sub
Private Sub Form_Load()
Me.Left = 50
Me.Top = 50
create.Visible = True
       id.Text = ArgNum + "  You Don't Change AggnId"
rs.Open "select * from agganwaritable where agganid= '" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
   txtRgNO.Text = rs.Fields("agganregno")
   txtDate.Text = Format(rs.Fields("rgdate"), "DD-MMM-YYYY")
   txtUppCenNo.Text = rs.Fields("uppercenter")
   TxtUppDate.Text = rs.Fields("cendate")
   TxtUppDate.Text = Format(TxtUppDate.Text, "DD-MMM-YYYY")
   txtAsha.Text = rs.Fields("asha")
   txtANM.Text = rs.Fields("anm")
   txtClinic.Text = rs.Fields("upperhc")
   TxtFHC.Text = rs.Fields("fhc")
   txthopital.Text = rs.Fields("hosfru")
   txtMobile.Text = rs.Fields("contectno")
   txtEmail.Text = rs.Fields("email")
   passkey = rs.Fields("password")
   txtOwner.Text = rs.Fields("owner")
   TxtBlock.Text = rs.Fields("center")
   rs.Close
   
   
   
End Sub





Private Sub id_KeyPress(KeyAscii As Integer)
msg = "  You Don't Change AggnId"
msgShow.Show
End Sub




Private Sub txtDate_Validate(cancel As Boolean)


If IsDate(txtDate.Text) Then
    txtDate.Text = Format(txtDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    txtDate.Text = ""
    txtDate.SetFocus
    End If

End Sub



Private Sub txtMobile_Validate(cancel As Boolean)
If IsNumeric(txtMobile.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    txtMobile.Text = ""
    txtMobile.SetFocus
    End If
End Sub

Private Sub txtOldPassword_Change()
create.Visible = True
create.Enabled = True
End Sub

Private Sub txtPass_Change()
 txtRePass.Visible = True
End Sub

Private Sub txtRgNO_Validate(cancel As Boolean)
If IsNumeric(txtRgNO.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    txtRgNO.Text = ""
    txtRgNO.SetFocus
    End If
End Sub

Private Sub txtUppCenNo_Validate(cancel As Boolean)
If IsNumeric(txtUppCenNo.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    txtUppCenNo.Text = ""
    txtUppCenNo.SetFocus
    End If
End Sub

Private Sub TxtUppDate_Validate(cancel As Boolean)
If IsDate(TxtUppDate.Text) Then
    TxtUppDate.Text = Format(TxtUppDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    TxtUppDate.Text = ""
    TxtUppDate.SetFocus
    End If
End Sub
