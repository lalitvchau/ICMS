VERSION 5.00
Begin VB.Form accForm 
   Appearance      =   0  'Flat
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "New Acoount"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17940
   BeginProperty Font 
      Name            =   "Cooper Black"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "accForm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "accForm.frx":08CA
   ScaleHeight     =   10950
   ScaleWidth      =   17940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   12360
      TabIndex        =   32
      ToolTipText     =   "Enter return password"
      Top             =   9240
      Visible         =   0   'False
      Width           =   4095
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
      Left            =   7080
      TabIndex        =   30
      ToolTipText     =   "Enter a New Password"
      Top             =   9240
      Width           =   3615
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
      Left            =   12360
      TabIndex        =   27
      ToolTipText     =   "Enter the Email ID"
      Top             =   8640
      Width           =   4095
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
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
      Height          =   375
      Left            =   7080
      MaxLength       =   12
      TabIndex        =   26
      ToolTipText     =   "Enter hte Contect Number of Agganwari "
      Top             =   8640
      Width           =   3615
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
      Left            =   7080
      TabIndex        =   24
      ToolTipText     =   "enter the Hospital and FRU"
      Top             =   8160
      Width           =   9375
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
      Left            =   7080
      TabIndex        =   21
      ToolTipText     =   "Enter the First Health Center and City"
      Top             =   7680
      Width           =   9375
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
      Left            =   7080
      TabIndex        =   19
      ToolTipText     =   "Enter thr upper health center or Clinic an"
      Top             =   7200
      Width           =   9375
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
      Left            =   12360
      MaxLength       =   1000
      TabIndex        =   17
      ToolTipText     =   "Enter name of ANM The Agaanwari"
      Top             =   6720
      Width           =   4095
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
      Left            =   7080
      TabIndex        =   15
      ToolTipText     =   "Enter name of Asha of The Agganwari"
      Top             =   6720
      Width           =   3615
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
      Left            =   12360
      TabIndex        =   12
      ToolTipText     =   "Enter agganwari center and Block"
      Top             =   6240
      Width           =   4095
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
      Left            =   7080
      TabIndex        =   10
      ToolTipText     =   "Enter name of T Agganwari Owner Name"
      Top             =   6240
      Width           =   3615
   End
   Begin VB.TextBox txtUppCenNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   9
      ToolTipText     =   "Enetr the Upper Center Number"
      Top             =   5760
      Width           =   3615
   End
   Begin VB.TextBox TxtUppDate 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12360
      MaxLength       =   11
      TabIndex        =   8
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Upper Center  Date in DD-MON-YYYY format Example 05-APR-1993"
      Top             =   5760
      Width           =   4095
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
      Left            =   12360
      MaxLength       =   11
      TabIndex        =   5
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Registration Date in DD-MON-YYYY format Example 05-APR-1993"
      Top             =   5280
      Width           =   4095
   End
   Begin VB.TextBox txtRgNO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
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
      Height          =   375
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   3
      ToolTipText     =   "Enter your Agganwari Center Registration no "
      Top             =   5280
      Width           =   3615
   End
   Begin VB.TextBox txtID 
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
      Left            =   7080
      MaxLength       =   30
      TabIndex        =   1
      ToolTipText     =   "Enter a new Agganwari Account ID example_ baliagan"
      Top             =   4800
      Width           =   9375
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
      Left            =   13560
      TabIndex        =   35
      Top             =   9840
      Width           =   2895
   End
   Begin VB.Label create 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "CREATE "
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
      Left            =   10560
      TabIndex        =   34
      Top             =   9840
      Width           =   2895
   End
   Begin VB.Label note 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   4080
      TabIndex        =   33
      Top             =   9840
      Width           =   75
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
      Left            =   10680
      TabIndex        =   31
      Top             =   9240
      Width           =   1695
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
      Left            =   4080
      TabIndex        =   29
      Top             =   9240
      Width           =   3015
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
      Left            =   10680
      TabIndex        =   28
      Top             =   8640
      Width           =   1695
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
      Left            =   4080
      TabIndex        =   25
      Top             =   8640
      Width           =   3015
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
      Left            =   4080
      TabIndex        =   23
      Top             =   8160
      Width           =   3015
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "New Aanganwadi Account"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   -600
      TabIndex        =   22
      Top             =   2520
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
      Left            =   4080
      TabIndex        =   20
      Top             =   7680
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
      Left            =   4080
      TabIndex        =   18
      Top             =   7200
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
      Left            =   10680
      TabIndex        =   16
      Top             =   6720
      Width           =   1695
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
      Left            =   4080
      TabIndex        =   14
      Top             =   6720
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
      Left            =   4080
      TabIndex        =   13
      Top             =   6240
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
      Left            =   10680
      TabIndex        =   11
      Top             =   6240
      Width           =   1695
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
      Left            =   10680
      TabIndex        =   7
      Top             =   5760
      Width           =   1695
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
      Left            =   4080
      TabIndex        =   6
      Top             =   5760
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
      Left            =   10680
      TabIndex        =   4
      Top             =   5280
      Width           =   1695
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
      Left            =   4080
      TabIndex        =   2
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label LblID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "New Aanganwadi ID"
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
      Left            =   4080
      TabIndex        =   0
      Top             =   4800
      Width           =   3015
   End
End
Attribute VB_Name = "accForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim test
Dim temp
Private Sub cancel_Click()
Unload Me
login_form.Show
End Sub
Private Static Function nullCheck() As Boolean
Dim rt As Boolean
rt = False
 If txtID.Text = "" Then
    MsgBox " Aaganwadi Id Is Blanks !"
    txtID.SetFocus
ElseIf txtRgNO.Text = "" Then
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
Private Sub log()

 If txtPass.Text = txtRePass.Text Then
        create.Visible = False
        conn.Execute "insert into agganwariTable(agganId,agganRegNo,rgDate,upperCenter,cendate,asha,anm,owner,center,upperhc,fhc,hosfru,contectno,email,password)values('" & txtID.Text & "'," & txtRgNO.Text & ",'" & txtDate.Text & "','" & txtUppCenNo.Text & "','" & TxtUppDate.Text & "','" & txtAsha.Text & "','" & txtANM.Text & "','" & TxtBlock.Text & "','" & txtOwner.Text & "','" & txtClinic.Text & "','" & TxtFHC.Text & "','" & txthopital.Text & "','" & txtMobile.Text & "','" & txtEmail.Text & "','" & txtPass.Text & "')"
        msg = "New Aanganwadi  Account Created Successfully! "
        
        txtRePass.Text = ""
        Unload Me
        login_form.Show
        msgShow.Show
    Else
         
         msg = " Password Miss Match! ReEnter Password"
         msgShow.Show
    End If

End Sub

Private Sub cancel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BorderStyle = 1
cancel.ForeColor = vbWhite
End Sub
Private Sub create_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
create.BorderStyle = 1
create.ForeColor = vbWhite
End Sub

Private Sub create_Click()
 Dim lk As Integer
 Dim tan
 lk = 0
  If nullCheck = True Then
  rs.Open "select * from agganwaritable ", conn, adOpenStatic, adLockReadOnly
  
  If rs.BOF = rs.EOF Then
      rs.Close
      If txtID.Text = "" Then
         txtID.SetFocus
         msg = "Please Fill Aanganwadi  Id Field"
         msgShow.Show
       Else
            rs.Open "select * from agganwaritable", conn, adOpenStatic, adLockReadOnly
          While Not rs.EOF
               tan = rs.Fields("agganid")
               If txtID.Text = tan Then
               lk = 1
               End If
         
               rs.MoveNext
            Wend
       
         If lk = 1 Then
            txtID.SetFocus
            msg = "Enter A Unique Aanganwadi  id!"
             msgShow.Show
            rs.Close
         Else
            rs.Close
            log
        End If
      End If
  Else
       rs.Open "select * from agganwaritable", conn, adOpenStatic, adLockReadOnly
          While Not rs.EOF
               tan = rs.Fields("agganid")
               If txtID.Text = tan Then
               lk = 1
               End If
         
               rs.MoveNext
            Wend
       
         If lk = 1 Then
            txtID.SetFocus
            msg = "Enter A Unique Aanganwadi  id!"
             msgShow.Show
            rs.Close
         Else
            rs.Close
            log
        End If
    End If
  End If
End Sub



Private Sub Form_Load()
create.Visible = False
temp = ""
End Sub








Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
create.BorderStyle = 0
create.ForeColor = vbGreen
cancel.BorderStyle = 0
cancel.ForeColor = vbRed
End Sub

Private Sub txtDate_Validate(cancel As Boolean)
If IsDate(txtDate.Text) Then
txtDate.Text = Format(txtDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a DATE! AlphaNumeric or Alphbates or Numbers Not Allowed!"
    txtDate.Text = ""
    txtDate.SetFocus
    End If
End Sub



Private Sub txtMobile_Validate(cancel As Boolean)
If IsNumeric(txtMobile.Text) Then
    
    Else
    MsgBox "Enter a Mobile Numbers! AlphaNumeric or Alphbates Not Allowed!"
    txtMobile.Text = ""
    txtMobile.SetFocus
    End If
End Sub

Private Sub txtPass_Change()
txtRePass.Visible = True
End Sub

Private Sub txtRePass_Change()
create.Visible = True
    create.Enabled = True
End Sub

Private Sub txtRgNO_Validate(cancel As Boolean)
If IsNumeric(txtRgNO.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates Not Allowed!"
    txtRgNO.Text = ""
    txtRgNO.SetFocus
    End If
End Sub

Private Sub txtUppCenNo_Validate(cancel As Boolean)
If IsNumeric(txtUppCenNo.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates Not Allowed!"
    txtUppCenNo.Text = ""
    txtUppCenNo.SetFocus
    End If
End Sub


Private Sub TxtUppDate_Validate(cancel As Boolean)
If IsDate(TxtUppDate.Text) Then
    TxtUppDate.Text = Format(TxtUppDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a DATE! AlphaNumeric or Alphbates or Numbers Not Allowed!"
    TxtUppDate.Text = ""
    TxtUppDate.SetFocus
    End If
End Sub
