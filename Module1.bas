Attribute VB_Name = "Module1"

Public fegpass
Public ArgNum
Public cal
Public msg, india
Public colorlist(5)
Public conn As New ADODB.Connection
Public conn1 As New ADODB.Connection
Public conn2 As New ADODB.Connection
Public con As New ADODB.Connection
Public lalit As Integer
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset


Public Sub Main()
On Error GoTo err
conn.ConnectionString = "Provider=MSDAORA.1;Password=abc123;User ID=kid;Persist Security Info=True"
conn1.ConnectionString = "Provider=MSDAORA.1;Password=abc123;User ID=kid;Persist Security Info=True"
conn2.ConnectionString = "Provider=MSDAORA.1;Password=abc123;User ID=kid;Persist Security Info=True"
con.ConnectionString = "Provider=MSDAORA.1;Password=abc123;User ID=kid;Persist Security Info=True"
conn.Open
conn1.Open
conn2.Open
con.Open
colorlist(0) = &H400000
colorlist(1) = &H4000&
colorlist(2) = &H400040
colorlist(3) = &H40&
colorlist(4) = &H0&
login_form.Show
Exit Sub
err:
info.Show
End Sub

