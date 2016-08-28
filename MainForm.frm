VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainForm 
   BackColor       =   &H00400000&
   Caption         =   "Intensity Care Of Mother and Kids"
   ClientHeight    =   5955
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11505
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MainForm.frx":08CA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5580
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   0
            TextSave        =   "8:37 AM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   0
            TextSave        =   "2/7/2014"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnNew 
      Caption         =   "&New"
      Begin VB.Menu mnuNewMother 
         Caption         =   "New &Mother Registration"
      End
      Begin VB.Menu mnuNewKid 
         Caption         =   "New &Kid Registration"
      End
   End
   Begin VB.Menu mnuUpdate 
      Caption         =   "&Update"
      Begin VB.Menu mnuUpdateAgganwari 
         Caption         =   "A&anganwadi "
      End
      Begin VB.Menu mnuUpdateKid 
         Caption         =   "&Kid"
      End
      Begin VB.Menu mnuUpdateMother 
         Caption         =   "&Mother"
      End
   End
   Begin VB.Menu mnuCheckUp 
      Caption         =   "&Check Up"
      Begin VB.Menu mnuCheckUpPregTime 
         Caption         =   "&Pregnancy  Time"
      End
      Begin VB.Menu mnuCheckUpBefDel 
         Caption         =   "&Before Delivery"
      End
      Begin VB.Menu mnuCheckUpAftDel 
         Caption         =   "A&fter Delivery"
      End
      Begin VB.Menu mnuCheckUpDose 
         Caption         =   "Kids &Vaccination and Dose"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Repor&t"
      Begin VB.Menu mnuReportKid 
         Caption         =   "&Kids Card"
      End
      Begin VB.Menu mnuReportMother 
         Caption         =   "&Mother Card"
      End
      Begin VB.Menu mnuReportANC 
         Caption         =   "A&NC Report"
      End
      Begin VB.Menu mnuReportDel 
         Caption         =   "Delivery Reports Of Ag&ganwari"
      End
      Begin VB.Menu mnuReportDose 
         Caption         =   "Vaccination Re&port and Dose"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Hel&p"
      Begin VB.Menu mnuReportAbout 
         Caption         =   "&About Us"
      End
      Begin VB.Menu signOut 
         Caption         =   "Sign Out"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MDIForm_Load()
india = 0
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If X > MainForm.ScaleWidth - 50 Then
    AboutApps.Show
    Else
     Unload AboutApps
    End If
 lk.Left = 0
lk.Top = MainForm.ScaleHeight - 1560
End Sub

Private Sub mnuCheckUpAftDel_Click()
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Mother Account Needed!"
         msgShow.Show
     Else
rs.Close
afterDel.Show
End If
End Sub

Private Sub mnuCheckUpBefDel_Click()
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Mother Account Needed!"
         msgShow.Show
     Else
rs.Close
beforDel.Show
End If
End Sub

Private Sub mnuCheckUpDose_Click()
rs.Open "select * from kidtable a, mothertable b where a.coupleno=b.coupleno AND b.agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Kid Account Needed!"
         msgShow.Show
     Else
rs.Close
kidVac.Show

End If
End Sub

Private Sub mnuCheckUpPregTime_Click()
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Mother Account Needed!"
         msgShow.Show
     Else
rs.Close
pragTime.Show

End If

End Sub

Private Sub mnuNewKid_Click()
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Mother Account Needed!"
         msgShow.Show
     Else
rs.Close
kidReg.Show

End If
End Sub

Private Sub mnuNewMother_Click()
motherReg.Show
End Sub

Private Sub mnuReportAbout_Click()
 abtus.Show
End Sub

Private Sub mnuReportANC_Click()
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Mother Account Needed!"
         msgShow.Show
     Else
rs.Close
anc.Show

End If
End Sub

Private Sub mnuReportDel_Click()
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Mother Account Needed!"
         msgShow.Show
     Else
rs.Close
Del.Show

End If
End Sub

Private Sub mnuReportDose_Click()
   rs.Open "select * from kidtable a, mothertable b where a.coupleno=b.coupleno AND b.agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Kid Account Needed!"
         msgShow.Show
     Else
rs.Close
Dose.Show
End If
End Sub

Private Sub mnuReportKid_Click()
 rs.Open "select * from kidtable a, mothertable b where a.coupleno=b.coupleno AND b.agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Kid Account Needed!"
         msgShow.Show
     Else
rs.Close
KidCard.Show

End If
End Sub

Private Sub mnuReportMother_Click()
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Mother Account Needed!"
         msgShow.Show
     Else
rs.Close
mothercard.Show

End If
End Sub

Private Sub mnuUpdateAgganwari_Click()
updateAccForm.Show
End Sub

Private Sub mnuUpdateKid_Click()
rs.Open "select * from kidtable a, mothertable b where a.coupleno=b.coupleno AND b.agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Kid Account Needed!"
         msgShow.Show
     Else
rs.Close
updateKidReg.Show

End If
End Sub

Private Sub mnuUpdateMother_Click()
rs.Open "select * from mothertable where agganid='" & ArgNum & "'", conn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
      rs.Close
         
         msg = "Mother Account Needed!"
         msgShow.Show
     Else
rs.Close
updateMotherReg.Show

End If
End Sub

Private Sub signOut_Click()
india = 1
Unload Me
login_form.Show
msg = ArgNum + "....... Your Account Sign Out Successfully!"
msgShow.Show


End Sub


