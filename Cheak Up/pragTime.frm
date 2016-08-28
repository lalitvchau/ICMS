VERSION 5.00
Begin VB.Form pragTime 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "pragTime.frx":0000
   ScaleHeight     =   4590
   ScaleWidth      =   16035
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox t1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   52
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter First month"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   3360
      MaxLength       =   11
      TabIndex        =   51
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Second month"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   4920
      MaxLength       =   11
      TabIndex        =   50
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Third month"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   6480
      MaxLength       =   11
      TabIndex        =   49
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Fourth month"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   8040
      MaxLength       =   11
      TabIndex        =   48
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Fifth month"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   9600
      MaxLength       =   11
      TabIndex        =   47
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Sixeth month"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   11160
      MaxLength       =   11
      TabIndex        =   46
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Seventh month"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   12720
      MaxLength       =   11
      TabIndex        =   45
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Eighth month"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   14280
      MaxLength       =   11
      TabIndex        =   44
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Nineth month"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox w9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   14280
      MaxLength       =   11
      TabIndex        =   43
      Text            =   "0"
      ToolTipText     =   "Enter Nineth month"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox w8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   12720
      MaxLength       =   11
      TabIndex        =   42
      Text            =   "0"
      ToolTipText     =   "Enter Eighth month"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox w7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   11160
      MaxLength       =   11
      TabIndex        =   41
      Text            =   "0"
      ToolTipText     =   "Enter Seventh month"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox w6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   9600
      MaxLength       =   11
      TabIndex        =   40
      Text            =   "0"
      ToolTipText     =   "Enter Sixeth month"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox w5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   8040
      MaxLength       =   11
      TabIndex        =   39
      Text            =   "0"
      ToolTipText     =   "Enter Fifth month"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox w4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   6480
      MaxLength       =   11
      TabIndex        =   38
      Text            =   "0"
      ToolTipText     =   "Enter Fourth month"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox w3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   4920
      MaxLength       =   11
      TabIndex        =   37
      Text            =   "0"
      ToolTipText     =   "Enter Third month"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox w2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   3360
      MaxLength       =   11
      TabIndex        =   36
      Text            =   "0"
      ToolTipText     =   "Enter Second month"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox w1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   35
      Text            =   "0"
      ToolTipText     =   "Enter First month"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox bd1M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   33
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter First month"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox bd2M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   3360
      MaxLength       =   11
      TabIndex        =   32
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Second month"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox bd3M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   4920
      MaxLength       =   11
      TabIndex        =   31
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Third month"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox bd4M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   6480
      MaxLength       =   11
      TabIndex        =   30
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Fourth month"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox bd5M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   8040
      MaxLength       =   11
      TabIndex        =   29
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Fifth month"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox bd6M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   9600
      MaxLength       =   11
      TabIndex        =   28
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Sixeth month"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox bd7M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   11160
      MaxLength       =   11
      TabIndex        =   27
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Seventh month"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox bd8M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   12720
      MaxLength       =   11
      TabIndex        =   26
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Eighth month"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox bd9m 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   14280
      MaxLength       =   11
      TabIndex        =   25
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Nineth month"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox rg9M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   14280
      MaxLength       =   11
      TabIndex        =   24
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Nineth month"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox rg8M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   12720
      MaxLength       =   11
      TabIndex        =   23
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Eighth month"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox rg7M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   11160
      MaxLength       =   11
      TabIndex        =   22
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Seventh month"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox rg6M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   9600
      MaxLength       =   11
      TabIndex        =   21
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Sixeth month"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox rg5M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   8040
      MaxLength       =   11
      TabIndex        =   20
      Text            =   "01-Jan-400"
      ToolTipText     =   "Enter Fifth month"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox rg4M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   6480
      MaxLength       =   11
      TabIndex        =   19
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Fourth month"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox rg3M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   4920
      MaxLength       =   11
      TabIndex        =   18
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Third month"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox rg2M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Left            =   3360
      MaxLength       =   11
      TabIndex        =   17
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Second month"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox rg1M 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   16
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter First month"
      Top             =   1920
      Width           =   1575
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
      TabIndex        =   2
      Text            =   "Couple No"
      ToolTipText     =   "You should select Couple no whom you want check entry"
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Command1 
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
      Left            =   13800
      TabIndex        =   56
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label ok 
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
      Left            =   11400
      TabIndex        =   55
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "TT SERIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "WEIGHT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   53
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Before Delvry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Ninth Month"
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
      Left            =   14280
      TabIndex        =   14
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Eight Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   12720
      TabIndex        =   13
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Seventh Month"
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
      Left            =   11160
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Sixth Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   9600
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Fifth Month"
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
      Left            =   8040
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Third Month"
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
      Left            =   4920
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Fourth Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Second Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "First Month"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Month >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
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
      TabIndex        =   4
      Top             =   720
      Width           =   6255
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
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Compulsory Check up During Pregnancy "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16095
   End
End
Attribute VB_Name = "pragTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bd5M_Validate(cancel As Boolean)
If IsDate(bd5M.Text) Then
    bd5M.Text = Format(bd5M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd5M.Text = ""
    bd5M.SetFocus
    End If
End Sub
Private Sub bd9sM_Validate(cancel As Boolean)
If IsDate(bd9m.Text) Then
    
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd9m.Text = ""
    bd9m.SetFocus
    End If
End Sub
Private Sub bd8M_Validate(cancel As Boolean)
If IsDate(bd8M.Text) Then
    bd8M.Text = Format(bd8M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd8M.Text = ""
    bd8M.SetFocus
    End If
End Sub
Private Sub bd7M_Validate(cancel As Boolean)
If IsDate(bd7M.Text) Then
    bd7M.Text = Format(bd7M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd7M.Text = ""
    bd7M.SetFocus
    End If
End Sub
Private Sub bd6M_Validate(cancel As Boolean)
If IsDate(bd6M.Text) Then
    bd6M.Text = Format(bd6M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd6M.Text = ""
    bd6M.SetFocus
    End If
End Sub

Private Sub bd4M_Validate(cancel As Boolean)
If IsDate(bd4M.Text) Then
    bd4M.Text = Format(bd4M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd4M.Text = ""
    bd4M.SetFocus
    End If
End Sub
Private Sub bd3M_Validate(cancel As Boolean)
If IsDate(bd3M.Text) Then
    bd3M.Text = Format(bd3M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd3M.Text = ""
    bd3M.SetFocus
    End If
End Sub
Private Sub bd2M_Validate(cancel As Boolean)
If IsDate(bd2M.Text) Then
    bd2M.Text = Format(bd2M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd2M.Text = ""
    bd2M.SetFocus
    End If
End Sub
Private Sub bd1M_Validate(cancel As Boolean)
If IsDate(bd1M.Text) Then
    bd1M.Text = Format(bd1M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd1M.Text = ""
    bd1M.SetFocus
    End If
End Sub

Private Sub bd9m_Validate(cancel As Boolean)
If IsDate(bd9m.Text) Then
    bd9m.Text = Format(bd9m.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    bd9m.Text = ""
    bd9m.SetFocus
    End If
End Sub

Private Sub Command1_Click()
Unload Me


End Sub

Private Sub coupleNo_Click()
rs.Open " select * from mothertable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
lblMotherName.Caption = rs.Fields("plname")
rs.Close
ok.Enabled = True
ok.Visible = True
Dim check As Integer
rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
   check = rs.Fields("hope")
   rs.Close
  If check > 0 Then dis
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.ForeColor = vbRed
ok.ForeColor = vbGreen
ok.BorderStyle = 0
Command1.BorderStyle = 0
End Sub
Private Sub Command1_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command1.BorderStyle = 1
Command1.ForeColor = &H0&
End Sub

Private Sub ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

ok.BorderStyle = 1
ok.ForeColor = &H0&
End Sub

Private Sub Form_Load()
Me.Left = 50
Me.Top = 50
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
        coupleNo.AddItem rs.Fields("coupleNo")
        rs.MoveNext
              Wend
              rs.Close
End Sub
Private Sub dataq()
  Dim check As Integer
   rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
   check = rs.Fields("hope")
   rs.Close
    If check = 0 Then
       check = 1
       conn.Execute "update pragtable set hope='" & check & "',ladyname='" & lblMotherName.Caption & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "' where coupleno='" & coupleNo.Text & "'"
     ElseIf check = 1 Then
       check = 2
       conn.Execute "update pragtable set hope='" & check & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "',reg2='" & rg2M.Text & "',tt2='" & t2.Text & "',bd2='" & bd2M.Text & "',weight2='" & w2.Text & "' where coupleno='" & coupleNo.Text & "'"
     ElseIf check = 2 Then
       check = 3
       conn.Execute "update pragtable set hope='" & check & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "',reg2='" & rg2M.Text & "',tt2='" & t2.Text & "',bd2='" & bd2M.Text & "',weight2='" & w2.Text & "',reg3='" & rg3M.Text & "',tt3='" & t3.Text & "',bd3='" & bd3M.Text & "',weight3='" & w3.Text & "' where coupleno='" & coupleNo.Text & "'"
     ElseIf check = 3 Then
       check = 4
       conn.Execute "update pragtable set hope='" & check & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "',reg2='" & rg2M.Text & "',tt2='" & t2.Text & "',bd2='" & bd2M.Text & "',weight2='" & w2.Text & "',reg3='" & rg3M.Text & "',tt3='" & t3.Text & "',bd3='" & bd3M.Text & "',weight3='" & w3.Text & "',reg4='" & rg4M.Text & "',tt4='" & t4.Text & "',bd4='" & bd4M.Text & "',weight4='" & w4.Text & "' where coupleno='" & coupleNo.Text & "'"
     ElseIf check = 4 Then
       check = 5
       conn.Execute "update pragtable set hope='" & check & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "',reg2='" & rg2M.Text & "',tt2='" & t2.Text & "',bd2='" & bd2M.Text & "',weight2='" & w2.Text & "',reg3='" & rg3M.Text & "',tt3='" & t3.Text & "',bd3='" & bd3M.Text & "',weight3='" & w3.Text & "',reg4='" & rg4M.Text & "',tt4='" & t4.Text & "',bd4='" & bd4M.Text & "',weight4='" & w4.Text & "',reg5='" & rg5M.Text & "',tt5='" & t5.Text & "',bd5='" & bd5M.Text & "',weight5='" & w5.Text & "' where coupleno='" & coupleNo.Text & "'"
     ElseIf check = 5 Then
       check = 6
       conn.Execute "update pragtable set hope='" & check & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "',reg2='" & rg2M.Text & "',tt2='" & t2.Text & "',bd2='" & bd2M.Text & "',weight2='" & w2.Text & "',reg3='" & rg3M.Text & "',tt3='" & t3.Text & "',bd3='" & bd3M.Text & "',weight3='" & w3.Text & "',reg4='" & rg4M.Text & "',tt4='" & t4.Text & "',bd4='" & bd4M.Text & "',weight4='" & w4.Text & "',reg5='" & rg5M.Text & "',tt5='" & t5.Text & "',bd5='" & bd5M.Text & "',weight5='" & w5.Text & "',reg6='" & rg6M.Text & "',tt6='" & t6.Text & "',bd6='" & bd6M.Text & "',weight6='" & w6.Text & "' where coupleno='" & coupleNo.Text & "'"
     ElseIf check = 6 Then
       check = 7
       conn.Execute "update pragtable set hope='" & check & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "',reg2='" & rg2M.Text & "',tt2='" & t2.Text & "',bd2='" & bd2M.Text & "',weight2='" & w2.Text & "',reg3='" & rg3M.Text & "',tt3='" & t3.Text & "',bd3='" & bd3M.Text & "',weight3='" & w3.Text & "',reg4='" & rg4M.Text & "',tt4='" & t4.Text & "',bd4='" & bd4M.Text & "',weight4='" & w4.Text & "',reg5='" & rg5M.Text & "',tt5='" & t5.Text & "',bd5='" & bd5M.Text & "',weight5='" & w5.Text & "',reg6='" & rg6M.Text & "',tt6='" & t6.Text & "',bd6='" & bd6M.Text & "',weight6='" & w6.Text & "',reg7='" & rg7M.Text & "',tt7='" & t7.Text & "',bd7='" & bd7M.Text & "',weight7='" & w7.Text & "' where coupleno='" & coupleNo.Text & "'"
     ElseIf check = 7 Then
       check = 8
       conn.Execute "update pragtable set hope='" & check & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "',reg2='" & rg2M.Text & "',tt2='" & t2.Text & "',bd2='" & bd2M.Text & "',weight2='" & w2.Text & "',reg3='" & rg3M.Text & "',tt3='" & t3.Text & "',bd3='" & bd3M.Text & "',weight3='" & w3.Text & "',reg4='" & rg4M.Text & "',tt4='" & t4.Text & "',bd4='" & bd4M.Text & "',weight4='" & w4.Text & "',reg5='" & rg5M.Text & "',tt5='" & t5.Text & "',bd5='" & bd5M.Text & "',weight5='" & w5.Text & "',reg6='" & rg6M.Text & "',tt6='" & t6.Text & "',bd6='" & bd6M.Text & "',weight6='" & w6.Text & "',reg7='" & rg7M.Text & "',tt7='" & t7.Text & "',bd7='" & bd7M.Text & "',weight7='" & w7.Text & "',reg8='" & rg8M.Text & "',tt8='" & t8.Text & "',bd8='" & bd8M.Text & "',weight8='" & w8.Text & "' where coupleno='" & coupleNo.Text & "'"
     ElseIf check = 8 Then
       check = 9
       conn.Execute "update pragtable set hope='" & check & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "',reg2='" & rg2M.Text & "',tt2='" & t2.Text & "',bd2='" & bd2M.Text & "',weight2='" & w2.Text & "',reg3='" & rg3M.Text & "',tt3='" & t3.Text & "',bd3='" & bd3M.Text & "',weight3='" & w3.Text & "',reg4='" & rg4M.Text & "',tt4='" & t4.Text & "',bd4='" & bd4M.Text & "',weight4='" & w4.Text & "',reg5='" & rg5M.Text & "',tt5='" & t5.Text & "',bd5='" & bd5M.Text & "',weight5='" & w5.Text & "',reg6='" & rg6M.Text & "',tt6='" & t6.Text & "',bd6='" & bd6M.Text & "',weight6='" & w6.Text & "',reg7='" & rg7M.Text & "',tt7='" & t7.Text & "',bd7='" & bd7M.Text & "',weight7='" & w7.Text & "',reg8='" & rg8M.Text & "',tt8='" & t8.Text & "',bd8='" & bd8M.Text & "',weight8='" & w8.Text & "',reg9='" & rg9M.Text & "',tt9='" & t9.Text & "',bd9='" & bd9m.Text & "',weight9='" & w9.Text & "' where coupleno='" & coupleNo.Text & "'"
     Else
       conn.Execute "update pragtable set hope='" & check & "',reg1='" & rg1M.Text & "',tt1='" & t1.Text & "',bd1='" & bd1M.Text & "',weight1='" & w1.Text & "',reg2='" & rg2M.Text & "',tt2='" & t2.Text & "',bd2='" & bd2M.Text & "',weight2='" & w2.Text & "',reg3='" & rg3M.Text & "',tt3='" & t3.Text & "',bd3='" & bd3M.Text & "',weight3='" & w3.Text & "',reg4='" & rg4M.Text & "',tt4='" & t4.Text & "',bd4='" & bd4M.Text & "',weight4='" & w4.Text & "',reg5='" & rg5M.Text & "',tt5='" & t5.Text & "',bd5='" & bd5M.Text & "',weight5='" & w5.Text & "',reg6='" & rg6M.Text & "',tt6='" & t6.Text & "',bd6='" & bd6M.Text & "',weight6='" & w6.Text & "',reg7='" & rg7M.Text & "',tt7='" & t7.Text & "',bd7='" & bd7M.Text & "',weight7='" & w7.Text & "',reg8='" & rg8M.Text & "',tt8='" & t8.Text & "',bd8='" & bd8M.Text & "',weight8='" & w8.Text & "',reg9='" & rg9M.Text & "',tt9='" & t9.Text & "',bd9='" & bd9m.Text & "',weight9='" & w9.Text & "' where COUPLENO='" & coupleNo.Text & "'"
   End If
   msg = "Data Saved!"
    msgShow.Show
End Sub
Private Sub dis()
   Dim ch As Integer
   rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
   ch = rs.Fields("hope")
   rs.Close
   If ch = 1 Then
        
        
        rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
        check = rs.Fields("hope")
        
        w1.Text = rs.Fields("weight1")
        rg1M.Text = Format(rs.Fields("reg1"), "DD-MMM-YYYY")
        t1.Text = Format(rs.Fields("tt1"), "DD-MMM-YYYY")
        bd1M.Text = Format(rs.Fields("bd1"), "DD-MMM-YYYY")
        
        rs.Close
        rg2M.Enabled = True
        t2.Enabled = True
        bd2M.Enabled = True
        w2.Enabled = True
   ElseIf ch = 2 Then
        rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
        check = rs.Fields("hope")
        rg1M.Text = Format(rs.Fields("reg1"), "DD-MMM-YYYY")
        t1.Text = Format(rs.Fields("tt1"), "DD-MMM-YYYY")
        bd1M.Text = Format(rs.Fields("bd1"), "DD-MMM-YYYY")
        w1.Text = rs.Fields("weight1")
        
        rg2M.Text = Format(rs.Fields("reg2"), "DD-MMM-YYYY")
        t2.Text = Format(rs.Fields("tt2"), "DD-MMM-YYYY")
        bd2M.Text = Format(rs.Fields("bd2"), "DD-MMM-YYYY")
        w2.Text = rs.Fields("weight2")
        rs.Close
        
        rg2M.Enabled = True
        t2.Enabled = True
        bd2M.Enabled = True
        w2.Enabled = True
     
        rg3M.Enabled = True
        t3.Enabled = True
        bd3M.Enabled = True
        w3.Enabled = True
   ElseIf ch = 3 Then
        rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
        check = rs.Fields("hope")
        rg1M.Text = Format(rs.Fields("reg1"), "DD-MMM-YYYY")
        t1.Text = Format(rs.Fields("tt1"), "DD-MMM-YYYY")
        bd1M.Text = Format(rs.Fields("bd1"), "DD-MMM-YYYY")
        w1.Text = rs.Fields("weight1")
        
        rg2M.Text = Format(rs.Fields("reg2"), "DD-MMM-YYYY")
        t2.Text = Format(rs.Fields("tt2"), "DD-MMM-YYYY")
        bd2M.Text = Format(rs.Fields("bd2"), "DD-MMM-YYYY")
        w2.Text = rs.Fields("weight2")
        
        rg3M.Text = Format(rs.Fields("reg3"), "DD-MMM-YYYY")
        t3.Text = Format(rs.Fields("tt3"), "DD-MMM-YYYY")
        bd3M.Text = Format(rs.Fields("bd3"), "DD-MMM-YYYY")
        w3.Text = rs.Fields("weight3")
        
        rs.Close
        
        rg2M.Enabled = True
        t2.Enabled = True
        bd2M.Enabled = True
        w2.Enabled = True
     
        rg3M.Enabled = True
        t3.Enabled = True
        bd3M.Enabled = True
        w3.Enabled = True
        
        rg4M.Enabled = True
        t4.Enabled = True
        bd4M.Enabled = True
        w4.Enabled = True
        
   ElseIf ch = 4 Then
        rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
        check = rs.Fields("hope")
        rg1M.Text = Format(rs.Fields("reg1"), "DD-MMM-YYYY")
        t1.Text = Format(rs.Fields("tt1"), "DD-MMM-YYYY")
        bd1M.Text = Format(rs.Fields("bd1"), "DD-MMM-YYYY")
        w1.Text = rs.Fields("weight1")
        
        rg2M.Text = Format(rs.Fields("reg2"), "DD-MMM-YYYY")
        t2.Text = Format(rs.Fields("tt2"), "DD-MMM-YYYY")
        bd2M.Text = Format(rs.Fields("bd2"), "DD-MMM-YYYY")
        w2.Text = rs.Fields("weight2")
        
        rg3M.Text = Format(rs.Fields("reg3"), "DD-MMM-YYYY")
        t3.Text = Format(rs.Fields("tt3"), "DD-MMM-YYYY")
        bd3M.Text = Format(rs.Fields("bd3"), "DD-MMM-YYYY")
        w3.Text = rs.Fields("weight3")
        
        rg4M.Text = Format(rs.Fields("reg4"), "DD-MMM-YYYY")
        t4.Text = Format(rs.Fields("tt4"), "DD-MMM-YYYY")
        bd4M.Text = Format(rs.Fields("bd4"), "DD-MMM-YYYY")
        w4.Text = rs.Fields("weight4")
        
        rs.Close
        
        rg2M.Enabled = True
        t2.Enabled = True
        bd2M.Enabled = True
        w2.Enabled = True
     
        rg3M.Enabled = True
        t3.Enabled = True
        bd3M.Enabled = True
        w3.Enabled = True
        
        rg4M.Enabled = True
        t4.Enabled = True
        bd4M.Enabled = True
        w4.Enabled = True
        
        rg5M.Enabled = True
        t5.Enabled = True
        bd5M.Enabled = True
        w5.Enabled = True
        
   ElseIf ch = 5 Then
        rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
        check = rs.Fields("hope")
        rg1M.Text = Format(rs.Fields("reg1"), "DD-MMM-YYYY")
        t1.Text = Format(rs.Fields("tt1"), "DD-MMM-YYYY")
        bd1M.Text = Format(rs.Fields("bd1"), "DD-MMM-YYYY")
        w1.Text = rs.Fields("weight1")
        
        rg2M.Text = Format(rs.Fields("reg2"), "DD-MMM-YYYY")
        t2.Text = Format(rs.Fields("tt2"), "DD-MMM-YYYY")
        bd2M.Text = Format(rs.Fields("bd2"), "DD-MMM-YYYY")
        w2.Text = rs.Fields("weight2")
        
        rg3M.Text = Format(rs.Fields("reg3"), "DD-MMM-YYYY")
        t3.Text = Format(rs.Fields("tt3"), "DD-MMM-YYYY")
        bd3M.Text = Format(rs.Fields("bd3"), "DD-MMM-YYYY")
        w3.Text = rs.Fields("weight3")
        
        rg4M.Text = Format(rs.Fields("reg4"), "DD-MMM-YYYY")
        t4.Text = Format(rs.Fields("tt4"), "DD-MMM-YYYY")
        bd4M.Text = Format(rs.Fields("bd4"), "DD-MMM-YYYY")
        w4.Text = rs.Fields("weight4")
        
        rg5M.Text = Format(rs.Fields("reg5"), "DD-MMM-YYYY")
        t5.Text = Format(rs.Fields("tt5"), "DD-MMM-YYYY")
        bd5M.Text = Format(rs.Fields("bd5"), "DD-MMM-YYYY")
        w5.Text = rs.Fields("weight5")
        
        rs.Close
        
        rg2M.Enabled = True
        t2.Enabled = True
        bd2M.Enabled = True
        w2.Enabled = True
     
        rg3M.Enabled = True
        t3.Enabled = True
        bd3M.Enabled = True
        w3.Enabled = True
        
        rg4M.Enabled = True
        t4.Enabled = True
        bd4M.Enabled = True
        w4.Enabled = True
        
        rg5M.Enabled = True
        t5.Enabled = True
        bd5M.Enabled = True
        w5.Enabled = True
        
        rg6M.Enabled = True
        t6.Enabled = True
        bd6M.Enabled = True
        w6.Enabled = True
        
   ElseIf ch = 6 Then
        rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
        check = rs.Fields("hope")
        rg1M.Text = Format(rs.Fields("reg1"), "DD-MMM-YYYY")
        t1.Text = Format(rs.Fields("tt1"), "DD-MMM-YYYY")
        bd1M.Text = Format(rs.Fields("bd1"), "DD-MMM-YYYY")
        w1.Text = rs.Fields("weight1")
        
        rg2M.Text = Format(rs.Fields("reg2"), "DD-MMM-YYYY")
        t2.Text = Format(rs.Fields("tt2"), "DD-MMM-YYYY")
        bd2M.Text = Format(rs.Fields("bd2"), "DD-MMM-YYYY")
        w2.Text = rs.Fields("weight2")
        
        rg3M.Text = Format(rs.Fields("reg3"), "DD-MMM-YYYY")
        t3.Text = Format(rs.Fields("tt3"), "DD-MMM-YYYY")
        bd3M.Text = Format(rs.Fields("bd3"), "DD-MMM-YYYY")
        w3.Text = rs.Fields("weight3")
        
        rg4M.Text = Format(rs.Fields("reg4"), "DD-MMM-YYYY")
        t4.Text = Format(rs.Fields("tt4"), "DD-MMM-YYYY")
        bd4M.Text = Format(rs.Fields("bd4"), "DD-MMM-YYYY")
        w4.Text = rs.Fields("weight4")
        
        rg5M.Text = Format(rs.Fields("reg5"), "DD-MMM-YYYY")
        t5.Text = Format(rs.Fields("tt5"), "DD-MMM-YYYY")
        bd5M.Text = Format(rs.Fields("bd5"), "DD-MMM-YYYY")
        w5.Text = rs.Fields("weight5")
        
        rg6M.Text = Format(rs.Fields("reg6"), "DD-MMM-YYYY")
        t6.Text = Format(rs.Fields("tt6"), "DD-MMM-YYYY")
        bd6M.Text = Format(rs.Fields("bd6"), "DD-MMM-YYYY")
        w6.Text = rs.Fields("weight6")
     
        rs.Close
        
        rg2M.Enabled = True
        t2.Enabled = True
        bd2M.Enabled = True
        w2.Enabled = True
     
        rg3M.Enabled = True
        t3.Enabled = True
        bd3M.Enabled = True
        w3.Enabled = True
        
        rg4M.Enabled = True
        t4.Enabled = True
        bd4M.Enabled = True
        w4.Enabled = True
        
        rg5M.Enabled = True
        t5.Enabled = True
        bd5M.Enabled = True
        w5.Enabled = True
        
        rg6M.Enabled = True
        t6.Enabled = True
        bd6M.Enabled = True
        w6.Enabled = True
        
        rg7M.Enabled = True
        t7.Enabled = True
        bd7M.Enabled = True
        w7.Enabled = True
   ElseIf ch = 7 Then
        rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
        check = rs.Fields("hope")
        rg1M.Text = Format(rs.Fields("reg1"), "DD-MMM-YYYY")
        t1.Text = Format(rs.Fields("tt1"), "DD-MMM-YYYY")
        bd1M.Text = Format(rs.Fields("bd1"), "DD-MMM-YYYY")
        w1.Text = rs.Fields("weight1")
        
        rg2M.Text = Format(rs.Fields("reg2"), "DD-MMM-YYYY")
        t2.Text = Format(rs.Fields("tt2"), "DD-MMM-YYYY")
        bd2M.Text = Format(rs.Fields("bd2"), "DD-MMM-YYYY")
        w2.Text = rs.Fields("weight2")
        
        rg3M.Text = Format(rs.Fields("reg3"), "DD-MMM-YYYY")
        t3.Text = Format(rs.Fields("tt3"), "DD-MMM-YYYY")
        bd3M.Text = Format(rs.Fields("bd3"), "DD-MMM-YYYY")
        w3.Text = rs.Fields("weight3")
        
        rg4M.Text = Format(rs.Fields("reg4"), "DD-MMM-YYYY")
        t4.Text = Format(rs.Fields("tt4"), "DD-MMM-YYYY")
        bd4M.Text = Format(rs.Fields("bd4"), "DD-MMM-YYYY")
        w4.Text = rs.Fields("weight4")
        
        rg5M.Text = Format(rs.Fields("reg5"), "DD-MMM-YYYY")
        t5.Text = Format(rs.Fields("tt5"), "DD-MMM-YYYY")
        bd5M.Text = Format(rs.Fields("bd5"), "DD-MMM-YYYY")
        w5.Text = rs.Fields("weight5")
        
        rg6M.Text = Format(rs.Fields("reg6"), "DD-MMM-YYYY")
        t6.Text = Format(rs.Fields("tt6"), "DD-MMM-YYYY")
        bd6M.Text = Format(rs.Fields("bd6"), "DD-MMM-YYYY")
        w6.Text = rs.Fields("weight6")
        
        rg7M.Text = Format(rs.Fields("reg7"), "DD-MMM-YYYY")
        t7.Text = Format(rs.Fields("tt7"), "DD-MMM-YYYY")
        bd7M.Text = Format(rs.Fields("bd7"), "DD-MMM-YYYY")
        w7.Text = rs.Fields("weight7")
     
        rs.Close
        
        rg2M.Enabled = True
        t2.Enabled = True
        bd2M.Enabled = True
        w2.Enabled = True
     
        rg3M.Enabled = True
        t3.Enabled = True
        bd3M.Enabled = True
        w3.Enabled = True
        
        rg4M.Enabled = True
        t4.Enabled = True
        bd4M.Enabled = True
        w4.Enabled = True
        
        rg5M.Enabled = True
        t5.Enabled = True
        bd5M.Enabled = True
        w5.Enabled = True
        
        rg6M.Enabled = True
        t6.Enabled = True
        bd6M.Enabled = True
        w6.Enabled = True
        
        rg7M.Enabled = True
        t7.Enabled = True
        bd7M.Enabled = True
        w7.Enabled = True
        
        rg8M.Enabled = True
        t8.Enabled = True
        bd8M.Enabled = True
        w8.Enabled = True
ElseIf ch = 8 Then
        rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
        check = rs.Fields("hope")
        rg1M.Text = Format(rs.Fields("reg1"), "DD-MMM-YYYY")
        t1.Text = Format(rs.Fields("tt1"), "DD-MMM-YYYY")
        bd1M.Text = Format(rs.Fields("bd1"), "DD-MMM-YYYY")
        w1.Text = rs.Fields("weight1")
        
        rg2M.Text = Format(rs.Fields("reg2"), "DD-MMM-YYYY")
        t2.Text = Format(rs.Fields("tt2"), "DD-MMM-YYYY")
        bd2M.Text = Format(rs.Fields("bd2"), "DD-MMM-YYYY")
        w2.Text = rs.Fields("weight2")
        
        rg3M.Text = Format(rs.Fields("reg3"), "DD-MMM-YYYY")
        t3.Text = Format(rs.Fields("tt3"), "DD-MMM-YYYY")
        bd3M.Text = Format(rs.Fields("bd3"), "DD-MMM-YYYY")
        w3.Text = rs.Fields("weight3")
        
        rg4M.Text = Format(rs.Fields("reg4"), "DD-MMM-YYYY")
        t4.Text = Format(rs.Fields("tt4"), "DD-MMM-YYYY")
        bd4M.Text = Format(rs.Fields("bd4"), "DD-MMM-YYYY")
        w4.Text = rs.Fields("weight4")
        
        rg5M.Text = Format(rs.Fields("reg5"), "DD-MMM-YYYY")
        t5.Text = Format(rs.Fields("tt5"), "DD-MMM-YYYY")
        bd5M.Text = Format(rs.Fields("bd5"), "DD-MMM-YYYY")
        w5.Text = rs.Fields("weight5")
        
        rg6M.Text = Format(rs.Fields("reg6"), "DD-MMM-YYYY")
        t6.Text = Format(rs.Fields("tt6"), "DD-MMM-YYYY")
        bd6M.Text = Format(rs.Fields("bd6"), "DD-MMM-YYYY")
        w6.Text = rs.Fields("weight6")
        
        rg7M.Text = Format(rs.Fields("reg7"), "DD-MMM-YYYY")
        t7.Text = Format(rs.Fields("tt7"), "DD-MMM-YYYY")
        bd7M.Text = Format(rs.Fields("bd7"), "DD-MMM-YYYY")
        w7.Text = rs.Fields("weight7")
     
        rg8M.Text = Format(rs.Fields("reg8"), "DD-MMM-YYYY")
        t8.Text = Format(rs.Fields("tt8"), "DD-MMM-YYYY")
        bd8M.Text = Format(rs.Fields("bd8"), "DD-MMM-YYYY")
        w8.Text = rs.Fields("weight8")
     
        rs.Close
        
        rg2M.Enabled = True
        t2.Enabled = True
        bd2M.Enabled = True
        w2.Enabled = True
     
        rg3M.Enabled = True
        t3.Enabled = True
        bd3M.Enabled = True
        w3.Enabled = True
        
        rg4M.Enabled = True
        t4.Enabled = True
        bd4M.Enabled = True
        w4.Enabled = True
        
        rg5M.Enabled = True
        t5.Enabled = True
        bd5M.Enabled = True
        w5.Enabled = True
        
        rg6M.Enabled = True
        t6.Enabled = True
        bd6M.Enabled = True
        w6.Enabled = True
        
        rg7M.Enabled = True
        t7.Enabled = True
        bd7M.Enabled = True
        w7.Enabled = True
        
        rg8M.Enabled = True
        t8.Enabled = True
        bd8M.Enabled = True
        w8.Enabled = True
        
        rg9M.Enabled = True
        t9.Enabled = True
        bd9m.Enabled = True
        w9.Enabled = True
 Else
        rs.Open "select * from pragtable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
        check = rs.Fields("hope")
        rg1M.Text = Format(rs.Fields("reg1"), "DD-MMM-YYYY")
        t1.Text = Format(rs.Fields("tt1"), "DD-MMM-YYYY")
        bd1M.Text = Format(rs.Fields("bd1"), "DD-MMM-YYYY")
          
          
            w1.Text = rs.Fields("weight1")
          
        
        rg2M.Text = Format(rs.Fields("reg2"), "DD-MMM-YYYY")
        t2.Text = Format(rs.Fields("tt2"), "DD-MMM-YYYY")
        bd2M.Text = Format(rs.Fields("bd2"), "DD-MMM-YYYY")
        w2.Text = rs.Fields("weight2")
        
        rg3M.Text = Format(rs.Fields("reg3"), "DD-MMM-YYYY")
        t3.Text = Format(rs.Fields("tt3"), "DD-MMM-YYYY")
        bd3M.Text = Format(rs.Fields("bd3"), "DD-MMM-YYYY")
        w3.Text = rs.Fields("weight3")
        
        rg4M.Text = Format(rs.Fields("reg4"), "DD-MMM-YYYY")
        t4.Text = Format(rs.Fields("tt4"), "DD-MMM-YYYY")
        bd4M.Text = Format(rs.Fields("bd4"), "DD-MMM-YYYY")
        w4.Text = rs.Fields("weight4")
        
        rg5M.Text = Format(rs.Fields("reg5"), "DD-MMM-YYYY")
        t5.Text = Format(rs.Fields("tt5"), "DD-MMM-YYYY")
        bd5M.Text = Format(rs.Fields("bd5"), "DD-MMM-YYYY")
        w5.Text = rs.Fields("weight5")
        
        rg6M.Text = Format(rs.Fields("reg6"), "DD-MMM-YYYY")
        t6.Text = Format(rs.Fields("tt6"), "DD-MMM-YYYY")
        bd6M.Text = Format(rs.Fields("bd6"), "DD-MMM-YYYY")
        w6.Text = rs.Fields("weight6")
        
        rg7M.Text = Format(rs.Fields("reg7"), "DD-MMM-YYYY")
        t7.Text = Format(rs.Fields("tt7"), "DD-MMM-YYYY")
        bd7M.Text = Format(rs.Fields("bd7"), "DD-MMM-YYYY")
        w7.Text = rs.Fields("weight7")
     
        rg8M.Text = Format(rs.Fields("reg8"), "DD-MMM-YYYY")
        t8.Text = Format(rs.Fields("tt8"), "DD-MMM-YYYY")
        bd8M.Text = Format(rs.Fields("bd8"), "DD-MMM-YYYY")
        w8.Text = rs.Fields("weight8")
        
        rg9M.Text = Format(rs.Fields("reg9"), "DD-MMM-YYYY")
        t9.Text = Format(rs.Fields("tt9"), "DD-MMM-YYYY")
        bd9m.Text = Format(rs.Fields("bd9"), "DD-MMM-YYYY")
        w9.Text = rs.Fields("weight9")
     
        rs.Close
        
        rg2M.Enabled = True
        t2.Enabled = True
        bd2M.Enabled = True
        w2.Enabled = True
     
        rg3M.Enabled = True
        t3.Enabled = True
        bd3M.Enabled = True
        w3.Enabled = True
        
        rg4M.Enabled = True
        t4.Enabled = True
        bd4M.Enabled = True
        w4.Enabled = True
        
        rg5M.Enabled = True
        t5.Enabled = True
        bd5M.Enabled = True
        w5.Enabled = True
        
        rg6M.Enabled = True
        t6.Enabled = True
        bd6M.Enabled = True
        w6.Enabled = True
        
        rg7M.Enabled = True
        t7.Enabled = True
        bd7M.Enabled = True
        w7.Enabled = True
        
        rg8M.Enabled = True
        t8.Enabled = True
        bd8M.Enabled = True
        w8.Enabled = True
        
        rg9M.Enabled = True
        t9.Enabled = True
        bd9m.Enabled = True
        w9.Enabled = True

        
        End If
End Sub

Private Sub ok_Click()

dataq




Unload Me


End Sub

Private Sub rg1M_Validate(cancel As Boolean)
If IsDate(rg1M.Text) Then
    rg1M.Text = Format(rg1M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    rg1M.Text = ""
    rg1M.SetFocus
    End If
End Sub
Private Sub rg2M_Validate(cancel As Boolean)
If IsDate(rg2M.Text) Then
    rg2M.Text = Format(rg2M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    rg2M.Text = ""
    rg2M.SetFocus
    End If
End Sub
Private Sub rg3M_Validate(cancel As Boolean)
If IsDate(rg3M.Text) Then
    rg3M.Text = Format(rg3M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    rg3M.Text = ""
    rg3M.SetFocus
    End If
End Sub
Private Sub rg4M_Validate(cancel As Boolean)
If IsDate(rg4M.Text) Then
    rg4M.Text = Format(rg4M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    rg4M.Text = ""
    rg4M.SetFocus
    End If
End Sub
Private Sub rg9M_Validate(cancel As Boolean)
If IsDate(rg9M.Text) Then
    rg9M.Text = Format(rg9M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    rg9M.Text = ""
    rg9M.SetFocus
    End If
End Sub
Private Sub rg8M_Validate(cancel As Boolean)
If IsDate(rg8M.Text) Then
    rg8M.Text = Format(rg8M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    rg8M.Text = ""
    rg8M.SetFocus
    End If
End Sub
Private Sub rg7M_Validate(cancel As Boolean)
If IsDate(rg7M.Text) Then
    rg7M.Text = Format(rg7M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    rg7M.Text = ""
    rg7M.SetFocus
    End If
End Sub
Private Sub rg6M_Validate(cancel As Boolean)
If IsDate(rg6M.Text) Then
    rg6M.Text = Format(rg6M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    rg6M.Text = ""
    rg6M.SetFocus
    End If
End Sub
Private Sub rg5M_Validate(cancel As Boolean)
If IsDate(rg5M.Text) Then
    rg5M.Text = Format(rg5M.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    rg5M.Text = ""
    rg5M.SetFocus
    End If
End Sub



Private Sub t1_Validate(cancel As Boolean)
If IsDate(t1.Text) Then
    t1.Text = Format(t1.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    t1.Text = ""
    t1.SetFocus
    End If
End Sub

Private Sub t2_Validate(cancel As Boolean)
If IsDate(t2.Text) Then
    t2.Text = Format(t2.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    t2.Text = ""
    t2.SetFocus
    End If
End Sub

Private Sub t3_Validate(cancel As Boolean)
If IsDate(t3.Text) Then
    t3.Text = Format(t3.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    t3.Text = ""
    t3.SetFocus
    End If
End Sub


Private Sub t9_Validate(cancel As Boolean)
If IsDate(t9.Text) Then
    t9.Text = Format(t9.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    t9.Text = ""
    t9.SetFocus
    End If
End Sub
Private Sub t8_Validate(cancel As Boolean)
If IsDate(t8.Text) Then
    t8.Text = Format(t8.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    t8.Text = ""
    t8.SetFocus
    End If
End Sub
Private Sub t7_Validate(cancel As Boolean)
If IsDate(t7.Text) Then
    t7.Text = Format(t7.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    t7.Text = ""
    t7.SetFocus
    End If
End Sub
Private Sub t6_Validate(cancel As Boolean)
If IsDate(t6.Text) Then
    t6.Text = Format(t6.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    t6.Text = ""
    t6.SetFocus
    End If
End Sub
Private Sub t5_Validate(cancel As Boolean)
If IsDate(t5.Text) Then
    t5.Text = Format(t5.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    t5.Text = ""
    t5.SetFocus
    End If
End Sub
Private Sub t4_Validate(cancel As Boolean)
If IsDate(t4.Text) Then
    t4.Text = Format(t4.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a Date! Numeric or Alphbates  are Not Allowed!"
    t4.Text = ""
    t4.SetFocus
    End If
End Sub

Private Sub w1_Change()
ok.Enabled = True
End Sub

Private Sub w1_Validate(cancel As Boolean)
If IsNumeric(w1.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    w1.Text = "0"
    w1.SetFocus
    End If
End Sub
Private Sub w2_Validate(cancel As Boolean)
If IsNumeric(w2.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    w2.Text = "0"
    w2.SetFocus
    End If
End Sub
Private Sub w3_Validate(cancel As Boolean)
If IsNumeric(w3.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    w3.Text = "0"
    w3.SetFocus
    End If
End Sub
Private Sub w4_Validate(cancel As Boolean)
If IsNumeric(w4.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    w4.Text = "0"
    w4.SetFocus
    End If
End Sub
Private Sub w5_Validate(cancel As Boolean)
If IsNumeric(w5.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    w5.Text = "0"
    w5.SetFocus
    End If
End Sub
Private Sub w6_Validate(cancel As Boolean)
If IsNumeric(w6.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    w6.Text = "0"
    w6.SetFocus
    End If
End Sub
Private Sub w7_Validate(cancel As Boolean)
If IsNumeric(w7.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    w7.Text = "0"
    w7.SetFocus
    End If
End Sub
Private Sub w8_Validate(cancel As Boolean)
If IsNumeric(w8.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    w8.Text = "0"
    w8.SetFocus
    End If
End Sub
Private Sub w9_Validate(cancel As Boolean)
If IsNumeric(w9.Text) Then
    
    Else
    MsgBox "Enter a Numbers! AlphaNumeric or Alphbates  are Not Allowed!"
    w9.Text = "0"
    w9.SetFocus
    End If
End Sub
