VERSION 5.00
Begin VB.Form kidVac 
   Appearance      =   0  'Flat
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "From 16 to 36 Month"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   15015
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "From 24 to 36 Month"
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
         Height          =   1695
         Left            =   8040
         TabIndex        =   50
         Top             =   240
         Width           =   6375
         Begin VB.TextBox txtVitaminDose2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2880
            MaxLength       =   11
            TabIndex        =   53
            Text            =   "01-Jan-4000"
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtVitaminDose3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2880
            MaxLength       =   11
            TabIndex        =   52
            Text            =   "01-Jan-4000"
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtVitaminDose4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2880
            MaxLength       =   11
            TabIndex        =   51
            Text            =   "01-Jan-4000"
            Top             =   1200
            Width           =   3375
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Vitamin A Dose(24 Month)"
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
            TabIndex        =   56
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label V 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Vitamin A Dose(30 Month)"
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
            TabIndex        =   55
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Vitamin A Dose(36 Month)"
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
            TabIndex        =   54
            Top             =   1200
            Width           =   2775
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "From 16 to 24 Month"
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
         Height          =   1695
         Left            =   720
         TabIndex        =   43
         Top             =   240
         Width           =   6375
         Begin VB.TextBox txtDPTBoster 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2880
            MaxLength       =   11
            TabIndex        =   46
            Text            =   "01-Jan-4000"
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtPolioBoster 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2880
            MaxLength       =   11
            TabIndex        =   45
            Text            =   "01-Jan-4000"
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtVitaminBoster 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2880
            MaxLength       =   11
            TabIndex        =   44
            Text            =   "01-Jan-4000"
            Top             =   1200
            Width           =   3375
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "DPT Boster(16-24 Month)"
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
            TabIndex        =   49
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Polio Boster(16-24 Month)"
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
            TabIndex        =   48
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Vitamin A Dose(16 Month)"
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
            TabIndex        =   47
            Top             =   1200
            Width           =   2775
         End
      End
   End
   Begin VB.Frame kidVac 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Form Born to 3 Year"
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
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   15015
      Begin VB.Frame Vitami 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Vitamin-A"
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
         Height          =   1695
         Left            =   12840
         TabIndex        =   39
         Top             =   2160
         Width           =   2055
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   41
            Text            =   "DD-MON-YYYY"
            ToolTipText     =   "Ennter BCG Dose Date"
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtVitamin 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   120
            MaxLength       =   11
            TabIndex        =   40
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter VitaminDose Date"
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label txtvit9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Vit - A 9 Month"
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
            TabIndex        =   42
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Hipetaites"
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
         Height          =   1695
         Left            =   6480
         TabIndex        =   32
         Top             =   2160
         Width           =   6135
         Begin VB.TextBox txtHIP3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   35
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter BCG Dose Date"
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox txtHIP2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   34
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter BCG Dose Date"
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtHIP1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   33
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter BCG Dose Date"
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Hip-3 (3 Year 6 Month)"
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
            TabIndex        =   38
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Hip-2 (2 Year 6 Month)"
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
            TabIndex        =   37
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Hipetais 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Hip-1 (1 Year 6 Month)"
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
            TabIndex        =   36
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "DPT"
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
         Height          =   1695
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   6135
         Begin VB.TextBox txtDPT3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   28
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter DPT3 Dose Date"
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox txtDPT2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   27
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter DPT2 Dose Date"
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtDPT1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   26
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter DPT1 Dose Date"
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "DPT 3 (3Year 6 Month)"
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
            TabIndex        =   31
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "DPT 2 (2Year 6 Month)"
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
            TabIndex        =   30
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "DPT 1 (1Year 6 Month)"
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
            TabIndex        =   29
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Khasra"
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
         Height          =   1695
         Left            =   12840
         TabIndex        =   21
         Top             =   240
         Width           =   2055
         Begin VB.TextBox txtKhasra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   120
            MaxLength       =   11
            TabIndex        =   24
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter Khasra Dose Date"
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   22
            Text            =   "DD-MON-YYYY"
            ToolTipText     =   "Ennter BCG Dose Date"
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Khasra 9 Month"
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
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "At Born"
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
         Height          =   1695
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6135
         Begin VB.TextBox txtBCG 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   17
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter BCG Dose Date"
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtPol0 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   16
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter BCG Dose Date"
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txthlpB0 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   15
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter Hipetaits Dose Date"
            Top             =   1200
            Width           =   3375
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "BCG"
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
            TabIndex        =   20
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Poliyo-0*"
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
            TabIndex        =   19
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "HIpetaitis-B 0*"
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
            TabIndex        =   18
            Top             =   1200
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Poliyo"
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
         Height          =   1695
         Left            =   6480
         TabIndex        =   7
         Top             =   240
         Width           =   6135
         Begin VB.TextBox txtpoliyo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   10
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Enter Poliyo 1 Dose Date"
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtpoliyo2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   9
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Enter Poliyo 2 Dose Date"
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtPoliyo3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   2640
            MaxLength       =   11
            TabIndex        =   8
            Text            =   "01-Jan-4000"
            ToolTipText     =   "Ennter Poliyo 3 Dose Date"
            Top             =   1200
            Width           =   3375
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Poliyo 1 ( 1Year 6Month)"
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
            TabIndex        =   13
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Poliyo 2 ( 2Year 6Month)"
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
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Poliyo 3 ( 3Year 6Month)"
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
            TabIndex        =   11
            Top             =   1200
            Width           =   2535
         End
      End
   End
   Begin VB.ComboBox kidsNo 
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
      Left            =   2760
      TabIndex        =   1
      Text            =   "Kid Reg No"
      ToolTipText     =   "Select kids No"
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label save 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   10920
      TabIndex        =   57
      Top             =   7680
      Visible         =   0   'False
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
         Name            =   "Revel Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   13080
      TabIndex        =   58
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label kidName 
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
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Vaccination and Dose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "kidVac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer

Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
save.BackStyle = 0

save.ForeColor = &H80FF80
cancel.BackStyle = 0

cancel.ForeColor = &HFF&
End Sub
Private Sub save_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
save.BackStyle = 1
save.BackColor = &H80FF&
save.ForeColor = vbWhite
End Sub

Private Sub cancel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BackStyle = 1
cancel.BackColor = &H80FF&
cancel.ForeColor = vbWhite
End Sub


Private Sub Form_Load()
Me.Left = 50
Me.Top = 50
rs.Open "select * from kidtable,mothertable where mothertable.agganid='" & ArgNum & "' and kidtable.coupleno=mothertable.coupleno ", conn, adOpenStatic, adLockReadOnly
           
     While Not rs.EOF
        kidsNo.AddItem rs.Fields("kidregno")
        rs.MoveNext
              Wend
        rs.Close
    
End Sub

Private Sub kidsNo_Click()
 rs.Open "select * from kidtable where kidregno='" & kidsNo.Text & "' ", conn, adOpenStatic, adLockReadOnly
     kidname.Caption = rs.Fields("kidname")
    rs.Close
rs.Open " select *from dose where kidregno='" & kidsNo.Text & "'", conn, adOpenStatic, adLockReadOnly
     temp = rs.Fields("hope")
 
   rs.Close
 If temp > 0 Then dis
save.Visible = True
save.Enabled = True
End Sub
Private Sub dis()
    rs.Open " select *from dose where kidregno='" & kidsNo.Text & "'", conn, adOpenStatic, adLockReadOnly
     temp = rs.Fields("hope")
     rs.Close
  If temp = 1 Then
      rs.Open " select *from dose where kidregno='" & kidsNo.Text & "'", conn, adOpenStatic, adLockReadOnly
       txtBCG.Text = Format(rs.Fields("bcg"), "DD-MMM-YYYY")
       txtPol0.Text = Format(rs.Fields("poliy"), "DD-MMM-YYYY")
       txthlpB0.Text = Format(rs.Fields("hlp"), "DD-MMM-YYYY")
       
       txtpoliyo1.Enabled = True
       txtDPT1.Enabled = True
       txtHIP1.Enabled = True
       txtVitaminBoster.Enabled = True
       txtPolioBoster.Enabled = True
       txtDPTBoster.Enabled = True
       txtKhasra.Enabled = True
       txtVitamin.Enabled = True
       rs.Close
     ElseIf temp = 2 Then
       rs.Open " select *from dose where kidregno='" & kidsNo.Text & "'", conn, adOpenStatic, adLockReadOnly
       txtBCG.Text = Format(rs.Fields("bcg"), "DD-MMM-YYYY")
       txtPol0.Text = Format(rs.Fields("poliy"), "DD-MMM-YYYY")
       txthlpB0.Text = Format(rs.Fields("hlp"), "DD-MMM-YYYY")
       
       txtpoliyo1.Text = Format(rs.Fields("poliy1"), "DD-MMM-YYYY")
       txtDPT1.Text = Format(rs.Fields("dpt1"), "DD-MMM-YYYY")
       txtHIP1.Text = Format(rs.Fields("hip1"), "DD-MMM-YYYY")
       txtVitaminBoster.Text = Format(rs.Fields("vit1"), "DD-MMM-YYYY")
       txtPolioBoster.Text = Format(rs.Fields("poliyoboster"), "DD-MMM-YYYY")
       txtDPTBoster.Text = Format(rs.Fields("dptboster"), "DD-MMM-YYYY")
       txtKhasra.Text = Format(rs.Fields("kh9"), "DD-MMM-YYYY")
       txtVitamin.Text = Format(rs.Fields("vit9"), "DD-MMM-YYYY")
       
       
       txtpoliyo1.Enabled = True
       txtDPT1.Enabled = True
       txtHIP1.Enabled = True
       txtVitaminBoster.Enabled = True
       txtPolioBoster.Enabled = True
       txtDPTBoster.Enabled = True
       txtKhasra.Enabled = True
       txtVitamin.Enabled = True
       
       txtpoliyo2.Enabled = True
       txtDPT2.Enabled = True
       txtHIP2.Enabled = True
       txtVitaminDose2.Enabled = True
       rs.Close
     ElseIf temp = 3 Then
       rs.Open " select *from dose where kidregno='" & kidsNo.Text & "'", conn, adOpenStatic, adLockReadOnly
       txtBCG.Text = Format(rs.Fields("bcg"), "DD-MMM-YYYY")
       txtPol0.Text = Format(rs.Fields("poliy"), "DD-MMM-YYYY")
       txthlpB0.Text = Format(rs.Fields("hlp"), "DD-MMM-YYYY")
       
       txtpoliyo1.Text = Format(rs.Fields("poliy1"), "DD-MMM-YYYY")
       txtDPT1.Text = Format(rs.Fields("dpt1"), "DD-MMM-YYYY")
       txtHIP1.Text = Format(rs.Fields("hip1"), "DD-MMM-YYYY")
       txtVitaminBoster.Text = Format(rs.Fields("vit1"), "DD-MMM-YYYY")
       txtPolioBoster.Text = Format(rs.Fields("poliyoboster"), "DD-MMM-YYYY")
       txtDPTBoster.Text = Format(rs.Fields("dptboster"), "DD-MMM-YYYY")
       txtKhasra.Text = Format(rs.Fields("kh9"), "DD-MMM-YYYY")
       txtVitamin.Text = Format(rs.Fields("vit9"), "DD-MMM-YYYY")
       
       txtpoliyo2.Text = Format(rs.Fields("poliy2"), "DD-MMM-YYYY")
       txtDPT2.Text = Format(rs.Fields("dpt2"), "DD-MMM-YYYY")
       txtHIP2.Text = Format(rs.Fields("hip2"), "DD-MMM-YYYY")
       txtVitaminDose2.Text = Format(rs.Fields("vit2"), "DD-MMM-YYYY")
       
       txtpoliyo1.Enabled = True
       txtDPT1.Enabled = True
       txtHIP1.Enabled = True
       txtVitaminBoster.Enabled = True
       txtPolioBoster.Enabled = True
       txtDPTBoster.Enabled = True
       txtKhasra.Enabled = True
       txtVitamin.Enabled = True
       
       txtpoliyo2.Enabled = True
       txtDPT2.Enabled = True
       txtHIP2.Enabled = True
       txtVitaminDose2.Enabled = True
       
       txtPoliyo3.Enabled = True
       txtDPT3.Enabled = True
       txtHIP3.Enabled = True
       txtVitaminDose3.Enabled = True
       rs.Close
       ElseIf temp = 4 Then
       rs.Open " select *from dose where kidregno='" & kidsNo.Text & "'", conn, adOpenStatic, adLockReadOnly
       txtBCG.Text = Format(rs.Fields("bcg"), "DD-MMM-YYYY")
       txtPol0.Text = Format(rs.Fields("poliy"), "DD-MMM-YYYY")
       txthlpB0.Text = Format(rs.Fields("hlp"), "DD-MMM-YYYY")
       
       txtpoliyo1.Text = Format(rs.Fields("poliy1"), "DD-MMM-YYYY")
       txtDPT1.Text = Format(rs.Fields("dpt1"), "DD-MMM-YYYY")
       txtHIP1.Text = Format(rs.Fields("hip1"), "DD-MMM-YYYY")
       txtVitaminBoster.Text = Format(rs.Fields("vit1"), "DD-MMM-YYYY")
       txtPolioBoster.Text = Format(rs.Fields("poliyoboster"), "DD-MMM-YYYY")
       txtDPTBoster.Text = Format(rs.Fields("dptboster"), "DD-MMM-YYYY")
       txtKhasra.Text = Format(rs.Fields("kh9"), "DD-MMM-YYYY")
       txtVitamin.Text = Format(rs.Fields("vit9"), "DD-MMM-YYYY")
       
       txtpoliyo2.Text = Format(rs.Fields("poliy2"), "DD-MMM-YYYY")
       txtDPT2.Text = Format(rs.Fields("dpt2"), "DD-MMM-YYYY")
       txtHIP2.Text = Format(rs.Fields("hip2"), "DD-MMM-YYYY")
       txtVitaminDose2.Text = Format(rs.Fields("vit2"), "DD-MMM-YYYY")
       
       txtPoliyo3.Text = Format(rs.Fields("poliy3"), "DD-MMM-YYYY")
       txtDPT3.Text = Format(rs.Fields("dpt3"), "DD-MMM-YYYY")
       txtHIP3.Text = Format(rs.Fields("hip3"), "DD-MMM-YYYY")
       txtVitaminDose3.Text = Format(rs.Fields("vit3"), "DD-MMM-YYYY")
       
       
       txtpoliyo1.Enabled = True
       txtDPT1.Enabled = True
       txtHIP1.Enabled = True
       txtVitaminBoster.Enabled = True
       txtPolioBoster.Enabled = True
       txtDPTBoster.Enabled = True
       txtKhasra.Enabled = True
       txtVitamin.Enabled = True
       
       txtpoliyo2.Enabled = True
       txtDPT2.Enabled = True
       txtHIP2.Enabled = True
       txtVitaminDose2.Enabled = True
       
       txtPoliyo3.Enabled = True
       txtDPT3.Enabled = True
       txtHIP3.Enabled = True
       txtVitaminDose3.Enabled = True
    
       txtVitaminDose4.Enabled = True
       rs.Close
       ElseIf temp = 5 Then
       rs.Open " select *from dose where kidregno='" & kidsNo.Text & "'", conn, adOpenStatic, adLockReadOnly
       txtBCG.Text = Format(rs.Fields("bcg"), "DD-MMM-YYYY")
       txtPol0.Text = Format(rs.Fields("poliy"), "DD-MMM-YYYY")
       txthlpB0.Text = Format(rs.Fields("hlp"), "DD-MMM-YYYY")
       
       txtpoliyo1.Text = Format(rs.Fields("poliy1"), "DD-MMM-YYYY")
       txtDPT1.Text = Format(rs.Fields("dpt1"), "DD-MMM-YYYY")
       txtHIP1.Text = Format(rs.Fields("hip1"), "DD-MMM-YYYY")
       txtVitaminBoster.Text = Format(rs.Fields("vit1"), "DD-MMM-YYYY")
       txtPolioBoster.Text = Format(rs.Fields("poliyoboster"), "DD-MMM-YYYY")
       txtDPTBoster.Text = Format(rs.Fields("dptboster"), "DD-MMM-YYYY")
       txtKhasra.Text = Format(rs.Fields("kh9"), "DD-MMM-YYYY")
       txtVitamin.Text = Format(rs.Fields("vit9"), "DD-MMM-YYYY")
       
       txtpoliyo2.Text = Format(rs.Fields("poliy2"), "DD-MMM-YYYY")
       txtDPT2.Text = Format(rs.Fields("dpt2"), "DD-MMM-YYYY")
       txtHIP2.Text = Format(rs.Fields("hip2"), "DD-MMM-YYYY")
       txtVitaminDose2.Text = Format(rs.Fields("vit2"), "DD-MMM-YYYY")
       
       txtPoliyo3.Text = Format(rs.Fields("poliy3"), "DD-MMM-YYYY")
       txtDPT3.Text = Format(rs.Fields("dpt3"), "DD-MMM-YYYY")
       txtHIP3.Text = Format(rs.Fields("hip3"), "DD-MMM-YYYY")
       txtVitaminDose3.Text = Format(rs.Fields("vit3"), "DD-MMM-YYYY")
       txtVitaminDose4.Text = Format(rs.Fields("vit4"), "DD-MMM-YYYY")
       
       
       txtpoliyo1.Enabled = True
       txtDPT1.Enabled = True
       txtHIP1.Enabled = True
       txtVitaminBoster.Enabled = True
       txtPolioBoster.Enabled = True
       txtDPTBoster.Enabled = True
       txtKhasra.Enabled = True
       txtVitamin.Enabled = True
       
       txtpoliyo2.Enabled = True
       txtDPT2.Enabled = True
       txtHIP2.Enabled = True
       txtVitaminDose2.Enabled = True
       
       txtPoliyo3.Enabled = True
       txtDPT3.Enabled = True
       txtHIP3.Enabled = True
       txtVitaminDose3.Enabled = True
    
       txtVitaminDose4.Enabled = True
       rs.Close
  End If
 
End Sub
Private Sub inData()
     
    rs.Open " select *from dose where kidregno='" & kidsNo.Text & "'", conn, adOpenStatic, adLockReadOnly
    temp = rs.Fields("hope")
    rs.Close
    If temp = 0 Then
           temp = 1
           conn.Execute "update dose set hope='" & temp & "',KIDNAME='" & kidname.Caption & "', bcg='" & txtBCG.Text & "', poliy='" & txtPol0.Text & "',hlp='" & txthlpB0.Text & "'       where kidregno='" & kidsNo.Text & "'"
        ElseIf temp = 1 Then
           temp = 2
           conn.Execute "update dose set hope='" & temp & "',KIDNAME='" & kidname.Caption & "', bcg='" & txtBCG.Text & "', poliy='" & txtPol0.Text & "',hlp='" & txthlpB0.Text & "',poliy1='" & txtpoliyo1.Text & "',dpt1='" & txtDPT1.Text & "', hip1='" & txtHIP1.Text & "',vit1='" & txtVitaminBoster.Text & "',poliyoboster='" & txtPolioBoster.Text & "', dptboster='" & txtDPTBoster.Text & "', kh9='" & txtKhasra.Text & "',vit9='" & txtVitamin.Text & "'      where kidregno='" & kidsNo.Text & "'"
   
        ElseIf temp = 2 Then
           temp = 3
           conn.Execute "update dose set hope='" & temp & "',KIDNAME='" & kidname.Caption & "', bcg='" & txtBCG.Text & "', poliy='" & txtPol0.Text & "',hlp='" & txthlpB0.Text & "',poliy1='" & txtpoliyo1.Text & "',dpt1='" & txtDPT1.Text & "', hip1='" & txtHIP1.Text & "',vit1='" & txtVitaminBoster.Text & "',poliyoboster='" & txtPolioBoster.Text & "', dptboster='" & txtDPTBoster.Text & "', kh9='" & txtKhasra.Text & "',vit9='" & txtVitamin.Text & "',poliy2='" & txtpoliyo2.Text & "',dpt2='" & txtDPT2.Text & "', hip2='" & txtHIP2.Text & "',vit2='" & txtVitaminDose2.Text & "'      where kidregno='" & kidsNo.Text & "'"
 
       ElseIf temp = 3 Then
           temp = 4
           conn.Execute "update dose set hope='" & temp & "',KIDNAME='" & kidname.Caption & "', bcg='" & txtBCG.Text & "', poliy='" & txtPol0.Text & "',hlp='" & txthlpB0.Text & "',poliy1='" & txtpoliyo1.Text & "',dpt1='" & txtDPT1.Text & "', hip1='" & txtHIP1.Text & "',vit1='" & txtVitaminBoster.Text & "',poliyoboster='" & txtPolioBoster.Text & "', dptboster='" & txtDPTBoster.Text & "', kh9='" & txtKhasra.Text & "',vit9='" & txtVitamin.Text & "',poliy2='" & txtpoliyo2.Text & "',dpt2='" & txtDPT2.Text & "', hip2='" & txtHIP2.Text & "',vit2='" & txtVitaminDose2.Text & "',poliy3='" & txtPoliyo3.Text & "',dpt3='" & txtDPT3.Text & "', hip3='" & txtHIP3.Text & "',vit3='" & txtVitaminDose3.Text & "'      where kidregno='" & kidsNo.Text & "'"
       Else
           temp = 5
           conn.Execute "update dose set hope='" & temp & "',KIDNAME='" & kidname.Caption & "', bcg='" & txtBCG.Text & "', poliy='" & txtPol0.Text & "',hlp='" & txthlpB0.Text & "',poliy1='" & txtpoliyo1.Text & "',dpt1='" & txtDPT1.Text & "', hip1='" & txtHIP1.Text & "',vit1='" & txtVitaminBoster.Text & "',poliyoboster='" & txtPolioBoster.Text & "', dptboster='" & txtDPTBoster.Text & "', kh9='" & txtKhasra.Text & "',vit9='" & txtVitamin.Text & "',poliy2='" & txtpoliyo2.Text & "',dpt2='" & txtDPT2.Text & "', hip2='" & txtHIP2.Text & "',vit2='" & txtVitaminDose2.Text & "',poliy3='" & txtPoliyo3.Text & "',dpt3='" & txtDPT3.Text & "', hip3='" & txtHIP3.Text & "',vit3='" & txtVitaminDose3.Text & "',vit4='" & txtVitaminDose4.Text & "'      where kidregno='" & kidsNo.Text & "'"
    End If
End Sub

Private Sub save_Click()
inData
msg = " Data Saved!"
msgShow.Show
Unload Me
End Sub


Private Sub txtBCG_Validate(cancel As Boolean)
If IsDate(txtBCG.Text) Then
    txtBCG.Text = Format(txtBCG.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtBCG.Text = "01-Jan-4000"
    txtBCG.SetFocus
    End If
End Sub




Private Sub txtDPT1_Validate(cancel As Boolean)
If IsDate(txtDPT1.Text) Then
    txtDPT1.Text = Format(txtDPT1.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtDPT1.Text = ""
    txtDPT1.SetFocus
    End If
End Sub
Private Sub txtDPT2_Validate(cancel As Boolean)
If IsDate(txtDPT2.Text) Then
    txtDPT2.Text = Format(txtDPT2.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtDPT2.Text = ""
    txtDPT2.SetFocus
    End If
End Sub
Private Sub txtDPT3_Validate(cancel As Boolean)
If IsDate(txtDPT3.Text) Then
    txtDPT3.Text = Format(txtDPT3.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtDPT3.Text = ""
    txtDPT3.SetFocus
    End If
End Sub



Private Sub txtDPTBoster_Validate(cancel As Boolean)
If IsDate(txtDPTBoster.Text) Then
    txtDPTBoster.Text = Format(txtDPTBoster.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtDPTBoster.Text = "01-Jan-4000"
    txtDPTBoster.SetFocus
    End If
End Sub

Private Sub txtHIP1_Validate(cancel As Boolean)
If IsDate(txtHIP1.Text) Then
    txtHIP1.Text = Format(txtHIP1.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtHIP1.Text = "01-Jan-4000"
    txtHIP1.SetFocus
    End If
End Sub
Private Sub txtHIP2_Validate(cancel As Boolean)
If IsDate(txtHIP2.Text) Then
    txtHIP2.Text = Format(txtHIP2.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtHIP2.Text = "01-Jan-4000"
    txtHIP2.SetFocus
    End If
End Sub
Private Sub txtHIP3_Validate(cancel As Boolean)
If IsDate(txtHIP3.Text) Then
    txtHIP3.Text = Format(txtHIP3.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtHIP3.Text = "01-Jan-4000"
    txtHIP3.SetFocus
    End If
End Sub

Private Sub txthlpB0_Validate(cancel As Boolean)
If IsDate(txthlpB0.Text) Then
    txthlpB0.Text = Format(txthlpB0.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txthlpB0.Text = "01-Jan-4000"
    txthlpB0.SetFocus
    End If
End Sub


Private Sub txtKhasra_Validate(cancel As Boolean)
If IsDate(txtKhasra.Text) Then
    txtKhasra.Text = Format(txtKhasra.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtKhasra.Text = "01-Jan-4000"
    txtKhasra.SetFocus
    End If
End Sub


Private Sub txtPol0_Validate(cancel As Boolean)
If IsDate(txtPol0.Text) Then
    txtPol0.Text = Format(txtPol0.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtPol0.Text = "01-Jan-4000"
    txtPol0.SetFocus
    End If
End Sub

Private Sub txtPolioBoster_Validate(cancel As Boolean)
If IsDate(txtPolioBoster.Text) Then
    txtPolioBoster.Text = Format(txtPolioBoster.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtPolioBoster.Text = "01-Jan-4000"
    txtPolioBoster.SetFocus
    End If
End Sub

Private Sub txtpoliyo1_Validate(cancel As Boolean)
If IsDate(txtpoliyo1.Text) Then
    txtpoliyo1.Text = Format(txtpoliyo1.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtpoliyo1.Text = "01-Jan-4000"
    txtpoliyo1.SetFocus
    End If
End Sub
Private Sub txtpoliyo2_Validate(cancel As Boolean)
If IsDate(txtpoliyo2.Text) Then
    txtpoliyo2.Text = Format(txtpoliyo2.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtpoliyo2.Text = "01-Jan-4000"
    txtpoliyo2.SetFocus
    End If
End Sub
Private Sub txtPoliyo3_Validate(cancel As Boolean)
If IsDate(txtPoliyo3.Text) Then
    txtPoliyo3.Text = Format(txtPoliyo3.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtPoliyo3.Text = "01-Jan-4000"
    txtPoliyo3.SetFocus
    End If
End Sub


Private Sub txtVitamin_Validate(cancel As Boolean)
If IsDate(txtVitamin.Text) Then
    txtVitamin.Text = Format(txtVitamin.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtVitamin.Text = "01-Jan-4000"
    txtVitamin.SetFocus
    End If
End Sub



Private Sub txtVitaminBoster_Validate(cancel As Boolean)
If IsDate(txtVitaminBoster.Text) Then
    txtVitaminBoster.Text = Format(txtVitaminBoster.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtVitaminBoster.Text = "01-Jan-4000"
    txtVitaminBoster.SetFocus
    End If
End Sub



Private Sub txtVitaminDose2_Validate(cancel As Boolean)
If IsDate(txtVitaminDose2.Text) Then
    txtVitaminDose2.Text = Format(txtVitaminDose2.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtVitaminDose2.Text = "01-Jan-4000"
    txtVitaminDose2.SetFocus
    End If
End Sub
Private Sub txtVitaminDose3_Validate(cancel As Boolean)
If IsDate(txtVitaminDose3.Text) Then
    txtVitaminDose3.Text = Format(txtVitaminDose3.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtVitaminDose3.Text = "01-Jan-4000"
    txtVitaminDose3.SetFocus
    End If
End Sub

Private Sub txtVitaminDose4_Validate(cancel As Boolean)
If IsDate(txtVitaminDose4.Text) Then
    txtVitaminDose4.Text = Format(txtVitaminDose4.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! AlphaNumeric or Alphbates and Numbers are Not Allowed!"
    txtVitaminDose4.Text = ""
    txtVitaminDose4.SetFocus
    End If
End Sub
