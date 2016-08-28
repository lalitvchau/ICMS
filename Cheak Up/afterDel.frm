VERSION 5.00
Begin VB.Form afterDel 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17145
   Icon            =   "afterDel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   17145
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox deliveryResult 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      ItemData        =   "afterDel.frx":08CA
      Left            =   9240
      List            =   "afterDel.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3000
      Width           =   7815
   End
   Begin VB.TextBox dechargeDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      TabIndex        =   25
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter date of decharge from hospital"
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox kidWeight 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      Left            =   9240
      TabIndex        =   22
      ToolTipText     =   "Enter weight in KG"
      Top             =   2520
      Width           =   7815
   End
   Begin VB.ComboBox kidCry 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      ItemData        =   "afterDel.frx":08E5
      Left            =   2640
      List            =   "afterDel.frx":08EF
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ComboBox kidGender 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      ItemData        =   "afterDel.frx":08FC
      Left            =   2640
      List            =   "afterDel.frx":0906
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox probleDelivery 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      Left            =   9240
      TabIndex        =   16
      ToolTipText     =   "Enter delivery time problem if any"
      Top             =   2040
      Width           =   7815
   End
   Begin VB.ComboBox pragCombo 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      ItemData        =   "afterDel.frx":0918
      Left            =   2640
      List            =   "afterDel.frx":0925
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox hospitalPeriod 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      Left            =   9240
      TabIndex        =   11
      ToolTipText     =   "Enter hospitalize period if delivery at hospital"
      Top             =   1560
      Width           =   7815
   End
   Begin VB.ComboBox coupleNo 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      Left            =   2640
      TabIndex        =   7
      Text            =   "Couple No"
      ToolTipText     =   "You should select Couple no whom you want check entry"
      Top             =   600
      Width           =   3375
   End
   Begin VB.ComboBox DeliveryType 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      ItemData        =   "afterDel.frx":094B
      Left            =   15120
      List            =   "afterDel.frx":0958
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox deliveryPlace 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      Left            =   8520
      TabIndex        =   4
      ToolTipText     =   "Enter Delivery Place"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox deliveryDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      TabIndex        =   2
      Text            =   "01-Jan-4000"
      ToolTipText     =   "Enter Delivery Date example 12-JUN-2014"
      Top             =   1080
      Width           =   3375
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
      Left            =   15000
      TabIndex        =   28
      Top             =   3600
      Width           =   1935
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
      Left            =   12840
      TabIndex        =   29
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Delivery   Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   6000
      TabIndex        =   26
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Hospital Decharge Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Kids Weight at Born"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   4560
      TabIndex        =   23
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Kid Cry After Born"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Kids Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Any Problem at delivery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      Caption         =   "After Delivery Check Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   17175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Pragnancy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Hospitalize Period (Delivery at Hospital)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label lblMotherName 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      Left            =   8520
      TabIndex        =   10
      Top             =   600
      Width           =   8535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
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
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
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
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Delivery Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   12600
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Delivery Place"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Delivery Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      Caption         =   "After Delivery Check Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17175
   End
End
Attribute VB_Name = "afterDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer
Private Sub cancel_Click()
Unload Me

End Sub
Private Static Function nullCheck() As Boolean
Dim rt As Boolean
rt = False
 
     
If deliveryDate.Text = "" Then
    MsgBox " Delivery Date Is Blank !"
    txtBirthDate.SetFocus
ElseIf coupleNo.Text = "" Then
    MsgBox " Couple No is Not Selected !"
    coupleNo.SetFocus
    
ElseIf dechargeDate.Text = "" Then
    MsgBox " Decharge Date is  Blank !"
    dechargeDate.SetFocus
    

ElseIf deliveryPlace.Text = "" Then
    MsgBox " Delivery Place Not Selected !"
    deliveryPlace.SetFocus
    
ElseIf DeliveryType.Text = "" Then
    MsgBox " Delivery Type Is Blank !"
    DeliveryType.SetFocus
    
ElseIf pragCombo.Text = "" Then
    MsgBox " Pragency Is Blank !"
    pragCombo.SetFocus
ElseIf hospitalPeriod.Text = "" Then
    MsgBox " Hospital Period Is Blank !"
    hospitalPeriod.SetFocus
ElseIf kidGender.Text = "" Then
    MsgBox " Kid's Gender Is Blank !"
    kidGender.SetFocus
ElseIf kidCry.Text = "" Then
    MsgBox " Kid's Cry Is Blank !"
    kidCry.SetFocus
ElseIf kidWeight.Text = "" Then
    MsgBox " Kid's Weight Is Blank !"
    kidWeight.SetFocus
ElseIf deliveryResult.Text = "" Then
    MsgBox " Delivery Result Is Blank !"
    deliveryResult.SetFocus
ElseIf probleDelivery.Text = "" Then
    MsgBox " Delivery Problem Is Blank !"
    probleDelivery.SetFocus
Else
   rt = True
    End If
nullCheck = rt
End Function

Private Sub cancel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BackStyle = 1
cancel.BackColor = &H4000&
cancel.ForeColor = vbWhite
End Sub

Private Sub coupleNo_Click()
   rs.Open " select * from mothertable where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
   lblMotherName.Caption = rs.Fields("plname")
   rs.Close
   
   rs.Open "select * from afterdelivery where coupleno='" & coupleNo.Text & "'", conn, adOpenStatic, adLockReadOnly
   temp = rs.Fields("hope")
   
   If temp > 0 Then
     deliveryDate.Text = Format(rs.Fields("DELIVERYDATE"), "DD-MMM-YYYY")
     dechargeDate.Text = Format(rs.Fields("HOSPITALDECHARGEDATE"), "DD-MMM-YYYY")
     deliveryPlace.Text = rs.Fields("deliveryPlace")
     DeliveryType.Text = rs.Fields("DeliveryType")
     pragCombo.Text = rs.Fields("PRAGNANCEY")
     hospitalPeriod.Text = rs.Fields("HOSPITALPEROID")
     kidGender.Text = rs.Fields("kidGender")
     kidCry.Text = rs.Fields("KIDCRYAFTERBORN")
     kidWeight.Text = rs.Fields("kidWeight")
     deliveryResult.Text = rs.Fields("deliveryresult")
     probleDelivery.Text = rs.Fields("problem")
    
    End If
    
    rs.Close
   save.Enabled = True
   save.Visible = True
End Sub



Private Sub dechargeDate_Validate(cancel As Boolean)
If IsDate(dechargeDate.Text) Then
    dechargeDate.Text = Format(dechargeDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a DATE! AlphaNumeric or Alphbates or Numbers Not Allowed!"
    dechargeDate.Text = ""
    dechargeDate.SetFocus
    End If
End Sub

Private Sub deliveryDate_Validate(cancel As Boolean)
If IsDate(deliveryDate.Text) Then
    deliveryDate.Text = Format(deliveryDate.Text, "DD-MMM-YYYY")
    Else
    MsgBox "Enter a DATE! AlphaNumeric or Alphbates or Numbers Not Allowed!"
    deliveryDate.Text = ""
    deliveryDate.SetFocus
    End If
End Sub

Private Sub deliveryResult_Change()
save.Visible = True
End Sub

Private Sub Form_Load()
Me.Left = 50
Me.Top = 50
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
          
        coupleNo.AddItem rs.Fields("coupleno")
        rs.MoveNext
 
           Wend
           
        rs.Close

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
save.BackStyle = 0

save.ForeColor = &H80FF80
cancel.BackStyle = 0

cancel.ForeColor = &HFF&
End Sub

Private Sub kidWeight_Validate(cancel As Boolean)
If IsNumeric(kidWeight.Text) Then
    
    Else
    MsgBox "Enter a DATE! AlphaNumeric or Alphbates or Numbers Not Allowed!"
    kidWeight.Text = ""
    kidWeight.SetFocus
    End If
End Sub

Private Sub save_Click()
temp = 1
If nullCheck = True Then
conn.Execute "update AFTERDELIVERY set hope='" & temp & "',DELIVERYDATE='" & deliveryDate.Text & "',HOSPITALDECHARGEDATE='" & dechargeDate.Text & "',deliveryPlace= '" & deliveryPlace.Text & "',DeliveryType='" & DeliveryType.Text & "',PRAGNANCEY='" & pragCombo.Text & "',HOSPITALPEROID='" & hospitalPeriod.Text & "',kidGender='" & kidGender.Text & "',KIDCRYAFTERBORN='" & kidCry.Text & "',kidWeight='" & kidWeight.Text & "',deliveryresult='" & deliveryResult.Text & "',problem='" & probleDelivery.Text & "',LADYNAME='" & lblMotherName.Caption & "'    where coupleno='" & coupleNo.Text & "'"
msg = "Data Saved!"
msgShow.Show

Unload Me
End If
End Sub

Private Sub save_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
save.BackStyle = 1
save.BackColor = &H4000&
save.ForeColor = vbWhite
End Sub
