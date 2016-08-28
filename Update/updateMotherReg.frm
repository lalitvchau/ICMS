VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form updateMotherReg 
   Appearance      =   0  'Flat
   BackColor       =   &H00400040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox CoupleId 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   2880
      TabIndex        =   33
      Text            =   "Your Couple No"
      Top             =   600
      Width           =   6975
   End
   Begin VB.TextBox txtJSIDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      IMEMode         =   3  'DISABLE
      Left            =   11520
      MaxLength       =   13
      TabIndex        =   32
      Text            =   "DD-MM-YYYY"
      ToolTipText     =   "Enter JSI paid date"
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox txtJSIPaid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   2880
      TabIndex        =   30
      Text            =   " "
      ToolTipText     =   "Enter JSI paid money"
      Top             =   4440
      Width           =   5775
   End
   Begin VB.TextBox txtJSIREGn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   11520
      TabIndex        =   29
      Text            =   " "
      ToolTipText     =   "Enter your registration number"
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox txtChild 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   28
      Text            =   " "
      ToolTipText     =   "Enter your ctotal childs name"
      Top             =   3960
      Width           =   5775
   End
   Begin VB.TextBox txtPin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   11520
      TabIndex        =   22
      Text            =   " "
      ToolTipText     =   "Enter your city  name"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox txtMob 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   2880
      MaxLength       =   13
      TabIndex        =   21
      Text            =   " "
      ToolTipText     =   "Enter your city  name"
      Top             =   3480
      Width           =   5775
   End
   Begin VB.TextBox txtState 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   11520
      TabIndex        =   20
      Text            =   " "
      ToolTipText     =   "Enter your city  name"
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox txtdist 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   5880
      TabIndex        =   18
      Text            =   " "
      ToolTipText     =   "Enter your city  name"
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtCity 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   2880
      TabIndex        =   17
      Text            =   " "
      ToolTipText     =   "Enter your city  name"
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   2880
      TabIndex        =   15
      ToolTipText     =   "Enter your permanent address"
      Top             =   2520
      Width           =   6975
   End
   Begin VB.ComboBox txtEdu 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      ItemData        =   "updateMotherReg.frx":0000
      Left            =   7200
      List            =   "updateMotherReg.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   13
      ToolTipText     =   "Select Education"
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox bdate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   2880
      TabIndex        =   11
      Text            =   "DD-MM-YYY"
      ToolTipText     =   "Enter your birth date"
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtHusbandName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "Enter the lady's husband name"
      Top             =   1560
      Width           =   6975
   End
   Begin VB.TextBox txtMotherName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "Enter the lady name"
      Top             =   1080
      Width           =   6975
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      TabIndex        =   4
      Text            =   "Your Picture Path"
      ToolTipText     =   "Enter path of Lady Image"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Command1 
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
      Left            =   12360
      TabIndex        =   36
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label ok 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
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
      Left            =   10680
      TabIndex        =   35
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Command2 
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
      Left            =   12960
      TabIndex        =   34
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label txtJsiPaiDadte 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "JSI Paid Date"
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
      Left            =   8760
      TabIndex        =   31
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label txtJsiMony 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "JSI Paid Money"
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
      TabIndex        =   27
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label txtJsiRe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "JSI Registration No"
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
      Left            =   8760
      TabIndex        =   26
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label txtTotChild 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "Total Childs"
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
      TabIndex        =   25
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "City Pin"
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
      Left            =   8760
      TabIndex        =   24
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "Mobile Number"
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
      TabIndex        =   23
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "State"
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
      Left            =   8760
      TabIndex        =   19
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "City\Town And District"
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
      TabIndex        =   16
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "Address"
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
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   " Eduction"
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
      Left            =   5880
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
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
      TabIndex        =   10
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "Husband's Name"
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
      TabIndex        =   8
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label txtPragLady 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "Pragnent Lady Name"
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
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label txtCoupleID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label imageName 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
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
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   12360
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label agganId 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11640
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
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
      Height          =   255
      Left            =   9960
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      Caption         =   "Update Mother Registration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
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
Attribute VB_Name = "updateMotherReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bdate_Validate(cancel As Boolean)
  If IsDate(bdate.Text) Then
      bdate.Text = Format(bdate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bdate.Text = ""
    bdate.SetFocus
    End If
End Sub

Private Sub Command1_Click()


imageName.Caption = txtMotherName.Text
Command1.Caption = "Update Image"
CommonDialog1.CancelError = True
On Error GoTo err
   CommonDialog1.Flags = cdlCFBoth
   CommonDialog1.Filter = "(*.jpg)"
   CommonDialog1.FilterIndex = 2
   
   CommonDialog1.ShowOpen
   txtPath.Text = CommonDialog1.FileName
   Image1.Picture = LoadPicture(txtPath.Text)
   imageName.Caption = txtMotherName.Text
   Exit Sub
err:
MsgBox " Only jpeg Format Allow", vbExclamation
  End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub CoupleId_Click()
ok.Visible = True
 ok.Enabled = True

If CoupleId.Text = "Select" Then
     MsgBox " First Select A Couple Id"
     Else
   rs.Open "select * from mothertable where coupleno= '" & CoupleId.Text & "'", conn, adOpenStatic, adLockReadOnly
   txtMotherName.Text = rs.Fields("plname")
   txtHusbandName.Text = rs.Fields("husname")
   bdate.Text = Format(rs.Fields("bdate"), "DD-MMM-YYYY")
   txtEdu.Text = rs.Fields("education")
    Text2.Text = rs.Fields("address")
    txtCity.Text = rs.Fields("city")
    txtdist.Text = rs.Fields("dist")
    txtMob.Text = rs.Fields("mobno")
    txtPin.Text = rs.Fields("citypin")
    txtChild.Text = rs.Fields("totalchild")
    txtJSIREGn.Text = rs.Fields("jsiregno")
    txtJSIPaid.Text = rs.Fields("jsipaidmon")
    txtJSIDate.Text = Format(rs.Fields("jsidate"), "DD-MMM-YYYY")
    txtPath.Text = rs.Fields("photo")
    txtState.Text = rs.Fields("state")
    On Error GoTo err
    Image1.Picture = LoadPicture(txtPath.Text)
    rs.Close
    Exit Sub
err: MsgBox "Photo Not Avialable!"
   rs.Close
     End If
 
End Sub
Private Static Function nullCheck() As Boolean
Dim rt As Boolean
rt = False
 If txtMotherName.Text = "" Then
    MsgBox " Mother Name Is Blanks !"
    txtMotherName.SetFocus
ElseIf txtHusbandName.Text = "" Then
    MsgBox " Father Name Is Blank !"
    txtHusbandName.SetFocus
     
ElseIf bdate.Text = "" Then
    MsgBox " Birth Date Is Blank !"
    bdate.SetFocus
    
ElseIf txtEdu.Text = "" Then
    MsgBox " Edducation Is Blank !"
    txtEdu.SetFocus
    
ElseIf Text2.Text = "" Then
    MsgBox " Address Is Blank !"
    Text2.SetFocus
    
ElseIf txtCity.Text = "" Then
    MsgBox " City Is Blank !"
    txtCity.SetFocus
    
ElseIf txtdist.Text = "" Then
    MsgBox " Disit Is Blank !"
    txtdist.SetFocus
    
ElseIf txtPin.Text = "" Then
    MsgBox " City Pin Is Blank !"
    txtPin.SetFocus
    
ElseIf txtChild.Text = "" Then
    MsgBox " Child Is Blank !"
    txtChild.SetFocus
    
ElseIf txtJSIREGn.Text = "" Then
    MsgBox " JSY Registration Number Is Blank !"
    txtJSIREGn.SetFocus
    
ElseIf txtMob.Text = "" Then
    MsgBox " Mobile Number Is Blank !"
    txtMob.SetFocus
    
ElseIf txtJSIPaid.Text = "" Then
    MsgBox " JSY Paid Is Blank !"
    txtJSIPaid.SetFocus
    
ElseIf txtPath.Text = "" Then
    MsgBox " Photo Not Selected !"
    txtPath.SetFocus
    
ElseIf txtJSIDate.Text = "" Then
    MsgBox " JSY Date Is Blank !"
    txtJSIDate.SetFocus
ElseIf txtState.Text = "" Then
    MsgBox " State Is Blank !"
    txtState.SetFocus
Else
   rt = True
    End If
nullCheck = rt

End Function




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
         txtMotherName.Text = ""
         txtHusbandName.Text = ""
         bdate.Text = ""
         Text2.Text = ""
         txtCity.Text = ""
         txtState.Text = ""
         txtdist.Text = ""
         txtMob.Text = ""
         txtPin = ""
         txtJSIPaid.Text = ""
         txtJSIDate.Text = ""
         txtJSIREGn.Text = ""
         txtChild.Text = ""
         Image1.Picture = LoadPicture("")
         Me.Refresh
       End If
       
      Next
  End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackStyle = 1
Command2.BackColor = &H800080
Command2.ForeColor = vbWhite
End Sub

Private Sub ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ok.BackStyle = 1
ok.BackColor = &H800080
ok.ForeColor = vbWhite
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ok.BackStyle = 0

ok.ForeColor = &H80FF80
Command1.BackStyle = 0

Command1.ForeColor = &H80FF80
Command2.BackStyle = 0
Command2.ForeColor = &HFF&

End Sub
Private Sub Command1_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackStyle = 1
Command1.BackColor = &H800080
Command1.ForeColor = vbWhite
End Sub
Private Sub Form_Load()
Me.Left = 50
Me.Top = 50
agganId.Caption = ArgNum
    rs.Open "select * from mothertable where agganid='" & agganId.Caption & "' ", conn, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
        CoupleId.AddItem rs.Fields("coupleno")
        rs.MoveNext
              Wend
        rs.Close
    End Sub

Private Sub ok_Click()
If nullCheck = True Then
conn.Execute "update  mothertable set  plname='" & txtMotherName.Text & "', husname='" & txtHusbandName.Text & "', bdate='" & bdate.Text & "', education='" & txtEdu.Text & "', address='" & Text2.Text & "', city='" & txtCity.Text & "', dist='" & txtdist.Text & "', state='" & txtState.Text & "', mobno='" & txtMob.Text & "', citypin='" & txtPin.Text & "', totalchild='" & txtChild.Text & "', jsiregno='" & txtJSIREGn.Text & "', jsipaidmon='" & txtJSIPaid.Text & "', jsidate='" & txtJSIDate.Text & "', photo='" & txtPath.Text & "' where coupleno='" & CoupleId.Text & "'"
conn1.Execute "update kidtable set mothername='" & txtMotherName.Text & "', fathername='" & txtHusbandName.Text & "' where coupleno='" & CoupleId.Text & "'"
msg = "Update Successfully!"
msgShow.Show
Unload Me

End If
End Sub



Private Sub txtChild_Validate(cancel As Boolean)
If IsNumeric(txtChild.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    txtChild.Text = ""
    txtChild.SetFocus
    End If
End Sub

Private Sub txtEdu_KeyUp(KeyCode As Integer, Shift As Integer)
 Dim tt
  tt = txtEdu.Text
  Dim ct
  ct = txtEdu.ListCount
  If KeyCode = 13 Then
   For i = 1 To ct
     If tt = txtEdu.List(i - 1) Then
       
       txtEdu = txtEdu.List(i - 1)
       Exit For
       Else
         txtEdu.Text = ""
       End If
       
      Next
  End If
End Sub






Private Sub txtJSIDate_Validate(cancel As Boolean)
If IsDate(txtJSIDate.Text) Then
    txtJSIDate.Text = Format(txtJSIDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    txtJSIDate.Text = ""
    txtJSIDate.SetFocus
    End If
End Sub

Private Sub txtJSIPaid_Validate(cancel As Boolean)
If IsNumeric(txtJSIPaid.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    txtJSIPaid.Text = ""
    txtJSIPaid.SetFocus
    End If
End Sub
Private Sub txtJSIREGn_Validate(cancel As Boolean)
If IsNumeric(txtJSIREGn.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    txtJSIREGn.Text = ""
    txtJSIREGn.SetFocus
    End If
End Sub

Private Sub txtMob_Validate(cancel As Boolean)
If IsNumeric(txtMob.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    txtMob.Text = ""
    txtMob.SetFocus
    End If
End Sub

Private Sub txtPath_Change()
Command1.Visible = True
End Sub


Private Sub txtPin_Validate(cancel As Boolean)
If IsNumeric(txtPin.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    txtPin.Text = ""
    txtPin.SetFocus
    End If
End Sub
