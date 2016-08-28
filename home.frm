VERSION 5.00
Begin VB.Form home 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Intensity Care Of Mother and Kids"
   ClientHeight    =   5850
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   11520
   FillColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   Picture         =   "home.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuNew 
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
         Caption         =   "&Agganwari"
      End
      Begin VB.Menu mnuUpdateKid 
         Caption         =   "K&id"
      End
      Begin VB.Menu mnuUpdateMother 
         Caption         =   "M&other"
      End
   End
   Begin VB.Menu mnuCheak 
      Caption         =   "&Check Up"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCheakPregTime 
         Caption         =   "&Pregnancy  Time"
      End
      Begin VB.Menu mnuCheakBfDel 
         Caption         =   "&Before Delivery"
      End
      Begin VB.Menu mnuCheakAfDel 
         Caption         =   "A&fter Delivery"
         Tag             =   "ef"
      End
      Begin VB.Menu mnuCheakKidSer 
         Caption         =   "Kids &Vaccination"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuReportMother 
         Caption         =   "&Mother Report"
      End
      Begin VB.Menu mnuReportKids 
         Caption         =   "&Kids Report"
      End
      Begin VB.Menu mnuReportANC 
         Caption         =   "AN&C Report"
      End
      Begin VB.Menu mnuReportDel 
         Caption         =   "Delivery Reports Of Ag&ganwari"
      End
      Begin VB.Menu mnuReportVac 
         Caption         =   "Vaccination Re&port"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHealth 
         Caption         =   "&Health"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About U&s"
      End
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuCheakAfDel_Click()
afterDel.Visible = True
End Sub

Private Sub mnuCheakBfDel_Click()
beforDel.Visible = True

End Sub

Private Sub mnuCheakKidSer_Click()
kidVac.Visible = True
End Sub

Private Sub mnuCheakPregTime_Click()
pragTime.Visible = True


End Sub

Private Sub mnuNewKid_Click()
kidReg.Visible = True
End Sub

Private Sub mnuNewMother_Click()
motherReg.Visible = True
End Sub

Private Sub mnuUpdateAgganwari_Click()
updateAccForm.Visible = True

End Sub

Private Sub mnuUpdateKid_Click()
updateKidReg.Visible = True
End Sub

Private Sub mnuUpdateMother_Click()
updateMotherReg.Visible = True
End Sub
