VERSION 5.00
Begin VB.Form beforDel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "beforDel.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Other Check Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      TabIndex        =   40
      ToolTipText     =   "Please check right answer"
      Top             =   6720
      Width           =   17775
      Begin VB.TextBox bloodSugarDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   11760
         MaxLength       =   11
         TabIndex        =   52
         Text            =   "01-Jan-4000"
         ToolTipText     =   "Enter the date of blood suagar check up date"
         Top             =   1440
         Width           =   5895
      End
      Begin VB.TextBox bloodSugar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   3000
         TabIndex        =   49
         ToolTipText     =   "Enter quintity of Urin Sugar"
         Top             =   1440
         Width           =   5895
      End
      Begin VB.TextBox HBSDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   11760
         MaxLength       =   11
         TabIndex        =   48
         Text            =   "01-Jan-4000"
         ToolTipText     =   "Enter the date of HBS Anitgen"
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   3000
         TabIndex        =   43
         ToolTipText     =   "Enter quintity of Urin Sugar"
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox urineDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   11760
         MaxLength       =   11
         TabIndex        =   42
         Text            =   "01-Jan-4000"
         ToolTipText     =   "Enter the date of Urine check up"
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox urine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   3000
         TabIndex        =   41
         ToolTipText     =   "enter Urine Check up"
         Top             =   480
         Width           =   5895
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Blood Sugar"
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
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   8880
         TabIndex        =   50
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "HBS Anitgen"
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
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   8880
         TabIndex        =   46
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Urine "
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
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   8880
         TabIndex        =   44
         Top             =   960
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Important Check Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      ToolTipText     =   "Please check right answer"
      Top             =   5040
      Width           =   17775
      Begin VB.TextBox bloodGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   11760
         TabIndex        =   38
         ToolTipText     =   "Enter Blood Group and Rh type"
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox himo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   3000
         TabIndex        =   34
         ToolTipText     =   "Enter quintity of Hemoglobin"
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox albu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   11760
         TabIndex        =   33
         ToolTipText     =   "Enter quintity of urin Albumin"
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox urinSugar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         ToolTipText     =   "Enter quintity of Urin Sugar"
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Blood group and RH Type"
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
         Height          =   375
         Left            =   8880
         TabIndex        =   39
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Hemoglobin"
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
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Urine Albumin"
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
         Height          =   375
         Left            =   8880
         TabIndex        =   36
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Urine Sugar"
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
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Last Delivery Index"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Please check right answer"
      Top             =   3360
      Width           =   17775
      Begin VB.TextBox breast 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   11760
         TabIndex        =   30
         ToolTipText     =   "Enter Breast"
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   3000
         TabIndex        =   28
         ToolTipText     =   "Enter Lungs"
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox heartDel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   11760
         TabIndex        =   26
         ToolTipText     =   "Enter Heart"
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox normalPhase 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         ToolTipText     =   "Enter Normal Phase"
         Top             =   480
         Width           =   5895
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Breast"
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
         Height          =   375
         Left            =   8880
         TabIndex        =   29
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Lungs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Lungs"
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
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Heart"
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
         Height          =   375
         Left            =   8880
         TabIndex        =   25
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Normal Phase"
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
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Last Delivery Index"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Please check right answer"
      Top             =   2280
      Width           =   17775
      Begin VB.CheckBox tapadik 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "         Tapedik"
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
         TabIndex        =   21
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox bloodPresare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "      High Blood Presaure"
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
         Left            =   3120
         TabIndex        =   20
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox heart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "  Any Heart Deasese "
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
         Left            =   6240
         TabIndex        =   19
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox diabetes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "            Diabetes "
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
         Left            =   9000
         TabIndex        =   18
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox Asthma 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "                   Asthma     "
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
         Left            =   11640
         TabIndex        =   17
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "                Other"
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
         Left            =   14760
         TabIndex        =   16
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Last delivery Problems"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Please check right answer"
      Top             =   1200
      Width           =   17775
      Begin VB.CheckBox other2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "Other"
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
         Left            =   16680
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox weekKid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   " Draw Back in Kid at Born"
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
         Left            =   13560
         TabIndex        =   13
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox lsps 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "ALPS"
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
         Left            =   12600
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox pph 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "PPH"
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
         Left            =   11640
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox infPrag 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "  Infective Delivery"
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
         Left            =   9120
         TabIndex        =   10
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox blood 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   " Low   Blood Quentity"
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
         Left            =   6600
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox pih 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "Preg Induced Hypertension"
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
         Left            =   3360
         TabIndex        =   8
         Top             =   480
         Width           =   3255
      End
      Begin VB.CheckBox Aklempsiya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "         Aklempsiya"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox aph 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "  APH"
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
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.ComboBox coupleNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   3000
      TabIndex        =   1
      Text            =   "Couple No"
      ToolTipText     =   "You should select Couple no whom you want check entry"
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label cancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Revel Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   15720
      TabIndex        =   54
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
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
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   13560
      TabIndex        =   53
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Lady Name"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblMotherName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   9600
      TabIndex        =   2
      Top             =   720
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Care Before Delivery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18015
   End
End
Attribute VB_Name = "beforDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(16)
Dim j As Integer
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.ForeColor = &H4000&
cancel.ForeColor = &HFF&
cancel.BorderStyle = 0
Command1.BorderStyle = 0
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command1.BorderStyle = 1
Command1.ForeColor = &H0&
End Sub

Private Sub cancel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cancel.BorderStyle = 1
cancel.ForeColor = &H0&
End Sub

Private Sub Aklempsiya_Click()
If Aklempsiya.Value = Yes Then
    arr(1) = "No"
  Else
    arr(1) = "Yes"
  End If

End Sub

Private Sub albu_Validate(cancel As Boolean)
If albu.Text = "" Then
     albu.Text = "NO"
     End If
End Sub

Private Sub aph_Click()
If aph.Value = Yes Then
    arr(0) = "No"
  Else
    arr(0) = "Yes"
  End If

End Sub

Private Sub Asthma_Click()
If Asthma.Value = Yes Then
    arr(13) = "No"
  Else
    arr(13) = "Yes"
  End If

End Sub

Private Sub blood_Click()
If blood.Value = Yes Then
    arr(3) = "No"
  Else
    arr(3) = "Yes"
  End If

End Sub

Private Sub bloodGroup_Validate(cancel As Boolean)
If bloodGroup.Text = "" Then
     bloodGroup.Text = "N"
     End If
End Sub

Private Sub bloodPresare_Click()
If bloodPresare.Value = Yes Then
    arr(10) = "No"
  Else
    arr(10) = "Yes"
  End If

End Sub



Private Sub bloodSugar_Validate(cancel As Boolean)
 If bloodSugar.Text = "" Then
     bloodSugar.Text = "NO"
     End If
End Sub

Private Sub bloodSugarDate_Validate(cancel As Boolean)
If IsDate(bloodSugarDate.Text) Then
    bloodSugarDate.Text = Format(bloodSugarDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a DATE! AlphaNumeric or Alphbates or Numbers Not Allowed!"
    bloodSugarDate.Text = "05-Apr-1992"
    bloodSugarDate.SetFocus
    End If
End Sub

Private Sub breast_Validate(cancel As Boolean)
If breast.Text = "" Then
     breast.Text = "NO"
     End If
End Sub

Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Check1_Click()
If Check1.Value = Yes Then
    arr(14) = "No"
  Else
    arr(14) = "Yes"
  End If

End Sub

Private Sub Command1_Click()

conn.Execute " insert into beforedel(coupleno,abh,AKLEMPSIYA,PRGINDUCED,LBLOODQUNT,INFECTDEL,PPH,ALPS,DRAWBACK,OTHER,TYPEDIK,HIGHBLOOD,ANYHEART,DIABITIES,ASTHMA,OTHER2,NORMALPHASE,HEART,LAXYS,BREAST,HEMOGLOBIN,URINEALBUMIN,URINSUGAR,BLOODGROUP,URINE,URIDATE,HBS,HBSDATE,BLOODSUG,BLOODDATE) values('" & coupleNo.Text & "','" & arr(0) & "','" & arr(1) & "','" & arr(2) & "','" & arr(3) & "','" & arr(4) & "','" & arr(5) & "','" & arr(6) & "','" & arr(7) & "','" & arr(8) & "','" & arr(9) & "','" & arr(10) & "','" & arr(11) & "','" & arr(12) & "','" & arr(13) & "','" & arr(14) & "','" & normalPhase.Text & "', '" & heartDel.Text & "','" & Text2.Text & "','" & breast.Text & "','" & himo.Text & "','" & albu.Text & "','" & urinSugar.Text & "','" & bloodGroup.Text & "','" & urine.Text & "','" & urineDate.Text & "','" & Text5.Text & "','" & HBSDate.Text & "','" & bloodSugar.Text & "','" & bloodSugarDate.Text & "')"
msg = "Data Saved!"
msgShow.Show
Unload Me

End Sub

Private Sub coupleNo_Click()
Command1.Visible = True
Command1.Enabled = True
rs.Open " select * from mothertable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
lblMotherName.Caption = rs.Fields("plname")
rs.Close
End Sub

Private Sub diabetes_Click()
If diabetes.Value = Yes Then
    arr(12) = "No"
  Else
    arr(12) = "Yes"
  End If

End Sub

Private Sub Form_Load()
Me.Left = 50
Me.Top = 50

rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  i = 1
  While Not rs.EOF
          
        coupleNo.AddItem rs.Fields("coupleno")
        rs.MoveNext
 
           Wend
           
        rs.Close
 j = 0
  While j < 16
     arr(j) = "No"
     j = j + 1
     Wend

End Sub



Private Sub HBSDate_Validate(cancel As Boolean)
If IsDate(HBSDate.Text) Then
    HBSDate.Text = Format(HBSDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a DATE! AlphaNumeric or Alphbates or Numbers Not Allowed!"
    HBSDate.Text = "05-Apr-1992"
    HBSDate.SetFocus
    End If
End Sub

Private Sub heart_Click()
If heart.Value = Yes Then
    arr(11) = "No"
  Else
    arr(11) = "Yes"
  End If

End Sub

Private Sub heartDel_Validate(cancel As Boolean)
If heartDel.Text = "" Then
     heartDel.Text = "NO"
     End If
End Sub

Private Sub himo_Validate(cancel As Boolean)
If himo.Text = "" Then
     himo.Text = "NO"
     End If
End Sub

Private Sub infPrag_Click()
If infPrag.Value = Yes Then
    arr(4) = "No"
  Else
    arr(4) = "Yes"
  End If

End Sub



Private Sub lsps_Click()
If lsps.Value = Yes Then
    arr(6) = "No"
  Else
    arr(6) = "Yes"
  End If

End Sub

Private Sub normalPhase_Validate(cancel As Boolean)
If normalPhase.Text = "" Then
     normalPhase.Text = "NO"
     End If
End Sub

Private Sub other2_Click()
If other.Value = Yes Then
    arr(8) = "No"
  Else
    arr(8) = "Yes"
  End If

End Sub

Private Sub pih_Click()
If pih.Value = Yes Then
    arr(2) = "No"
  Else
    arr(2) = "Yes"
  End If

End Sub

Private Sub pph_Click()
If pph.Value = Yes Then
    arr(5) = "No"
  Else
    arr(5) = "Yes"
  End If

End Sub

Private Sub tapadik_Click()
If tapadik.Value = Yes Then
    arr(9) = "No"
  Else
    arr(9) = "Yes"
  End If

End Sub

Private Sub Text2_Validate(cancel As Boolean)
If Text2.Text = "" Then
     Text2.Text = "NO"
     End If
End Sub

Private Sub Text5_Validate(cancel As Boolean)
If Text5.Text = "" Then
     Text5.Text = "NO"
     End If
End Sub

Private Sub urine_Validate(cancel As Boolean)
If urine.Text = "" Then
     urine.Text = "NO"
     End If
End Sub

Private Sub urineDate_Validate(cancel As Boolean)
If IsDate(urineDate.Text) Then
    urineDate.Text = Format(urineDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a DATE! AlphaNumeric or Alphbates or Numbers Not Allowed!"
    urineDate.Text = "05-Apr-1992"
    urineDate.SetFocus
    End If
End Sub

Private Sub urinSugar_Validate(cancel As Boolean)
If urinSugar.Text = "" Then
     urinSugar.Text = "NO"
     End If
End Sub

Private Sub weekKid_Click()
If weekKid.Value = Yes Then
    arr(7) = "No"
  Else
    arr(7) = "Yes"
  End If

End Sub
