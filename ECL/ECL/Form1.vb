Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Public Class fmEcl
    Shared contacts As List(Of Contact)
    Dim test As String

    Private Sub fmEcl_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        contacts = DataAccess.LoadData()

        For Each c In contacts
            lbListBox.Items.Add(c.name)
        Next

        Dim vState() As String = {"ME", "NH", "VT", "MA", "CT", "RI", "NY", "VA", "NC",
                "SC", "GA", "FL", "DE", "NJ", "OH", "MI", "IL", "IN", "IA", "KS", "NE", "OK",
                "TX", "AL", "TN", "MO", "ND", "SD", "WY", "MT", "ID", "NV", "WA", "OR", "CO",
                "NM", "AZ", "WV", "PA", "CA", "AR", "HI", "MN", "WI", "MD", "MS", "AK", "LA",
                "UT", "KY"}
        Array.Sort(vState)
        For Each s As String In vState
            cbState1.Items.Add(s)
            cbState2.Items.Add(s)
        Next

    End Sub

    Private Function isValid(nameBox As TextBox, empNum As MaskedTextBox, phoneNum As MaskedTextBox, emailAdd As TextBox)
        Dim valid As Boolean = True

        If Trim(nameBox.Text) = "" Then
            valid = False
            MessageBox.Show("Name cannot be blank!", "Contact Name",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf Trim(empNum.Text) = "" And valid Then
            valid = False
            MessageBox.Show("Employee number cannot be blank!", "Contact Employee Number",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf Trim(Replace(phoneNum.Text, "(   )    -", "")) = "" And valid Then
            valid = False
            MessageBox.Show("Phone number cannot be blank!", "Contact Phone number",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf Trim(emailAdd.Text = "") And valid Then
            valid = False
            MessageBox.Show("Email cannot be blank!", "Contact Email Address",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        If valid Then
            Try
                Dim em As System.Net.Mail.MailAddress = New System.Net.Mail.MailAddress(emailAdd.Text)
            Catch ex As Exception
                valid = False
                MessageBox.Show("Please enter a valid email address!", "Contact Email Address",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If


        Return valid
    End Function

    Private Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click

        If isValid(tbName1, tbEnumber1, tbPnumber1, tbEmail1) Then
            Dim result As MsgBoxResult = MsgBox("Permanently save changes to this contact's data?", MsgBoxStyle.YesNo)
            If result = MsgBoxResult.Yes Then
                Dim c As Contact = contacts(lbListBox.SelectedIndex)
                c.name = tbName1.Text
                c.phoneNum = tbPnumber1.Text
                c.street = tbStreet1.Text
                c.city = tbCity1.Text
                c.state = cbState1.Text
                c.zipCode = tbZcode1.Text
                c.emailAdd = tbEmail1.Text
                c.employeeNum = tbEnumber1.Text
                DataAccess.Save(contacts)
                Dim i As Integer = lbListBox.SelectedIndex
                lbListBox.Items(i) = c.name
                lbListBox.SelectedIndex = i
                btnSave1.Enabled = False
            End If
        End If


    End Sub

    Private Sub fillListBox()
        lbListBox.Items.Clear()
        For Each c In contacts
            lbListBox.Items.Add(c.name)
        Next
        DataAccess.Save(contacts)
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        If isValid(tbname2, tbEnumber2, tbPnumber2, tbEmail2) Then
            Dim c As New Contact(tbname2.Text, tbPnumber2.Text, tbStreet2.Text, tbCity2.Text,
                                         cbState2.Text, tbZcode2.Text, tbEmail2.Text, tbEnumber2.Text)
            contacts.Add(c)
            fillListBox()
            TabControl1.SelectTab(0)
            clearCreateBoxes()
        End If

    End Sub

    Private Sub lbListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbListBox.SelectedIndexChanged
        If lbListBox.SelectedIndex > -1 And lbListBox.SelectedIndex < lbListBox.Items.Count() Then
            Dim c As Contact = contacts(lbListBox.SelectedIndex)
            tbName1.Text = c.name
            tbPnumber1.Text = c.phoneNum
            tbStreet1.Text = c.street
            tbCity1.Text = c.city
            cbState1.SelectedItem = c.state
            tbZcode1.Text = c.zipCode
            tbEmail1.Text = c.emailAdd
            tbEnumber1.Text = c.employeeNum
            btnSave1.Enabled = False
        End If

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        Dim result As MsgBoxResult = MsgBox("Permanently delete this contact's data?", MsgBoxStyle.YesNo)
        If result = MsgBoxResult.Yes Then
            Dim s As String = lbListBox.Items(lbListBox.SelectedIndex)
            Dim i As Integer = lbListBox.SelectedIndex
            lbListBox.Items.Remove(s)
            contacts.RemoveAt(i)
            clearContactData()
            DataAccess.Save(contacts)
        End If


    End Sub

    Private Sub clearContactData()
        tbName1.Text = ""
        tbEnumber1.Text = ""
        tbPnumber1.Text = ""
        tbEmail1.Text = ""
        tbStreet1.Text = ""
        tbCity1.Text = ""
        cbState1.Text = ""
        tbZcode1.Text = ""
        btnSave1.Enabled = False
    End Sub

    Private Sub clearCreateBoxes()
        tbname2.Text = ""
        tbPnumber2.Text = ""
        tbStreet2.Text = ""
        tbCity2.Text = ""
        cbState2.Text = ""
        tbZcode2.Text = ""
        tbEmail2.Text = ""
        tbEnumber2.Text = ""
    End Sub

    Private Sub TabControl1_TabIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.TabIndexChanged
        Me.Size = New Size(1000, 1500)
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            Me.Size = New Size(622, 410)
        Else
            Me.Size = New Size(350, 410)
        End If

    End Sub

    Private Sub tbName1_TextChanged(sender As Object, e As EventArgs) Handles tbName1.TextChanged
        btnSave1.Enabled = True
    End Sub

    Private Sub tbEmail1_TextChanged(sender As Object, e As EventArgs) Handles tbEmail1.TextChanged
        btnSave1.Enabled = True
    End Sub

    Private Sub tbStreet1_TextChanged(sender As Object, e As EventArgs) Handles tbStreet1.TextChanged
        btnSave1.Enabled = True
    End Sub

    Private Sub tbCity1_TextChanged(sender As Object, e As EventArgs) Handles tbCity1.TextChanged
        btnSave1.Enabled = True
    End Sub

    Private Sub cbState1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbState1.SelectedIndexChanged
        btnSave1.Enabled = True
    End Sub

    Private Sub tbZcode1_TextChanged(sender As Object, e As EventArgs) Handles tbZcode1.TextChanged
        btnSave1.Enabled = True
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "CSV|*.csv"
        saveFileDialog1.Title = "Speichern CSV-Datei"
        If saveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim str As String = ""
            Using sw As System.IO.StreamWriter = New System.IO.StreamWriter(saveFileDialog1.OpenFile)
                str = "Name,Employee Number,Phone Number,Email Address,Street,City,State,Zip Code"
                sw.WriteLine(str)
                For Each c In contacts
                    str = c.name & "," & c.employeeNum & "," & c.phoneNum & "," & c.emailAdd & "," & c.street & "," & c.city & "," & c.state & "," & c.zipCode
                    sw.WriteLine(str)
                Next
                sw.Flush()
                sw.Close()
            End Using
        End If
    End Sub

    Private Sub tbPnumber1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbPnumber1.KeyPress
        btnSave1.Enabled = True
    End Sub

    Private Sub tbPnumber1_TextChanged(sender As Object, e As EventArgs) Handles tbPnumber1.TextChanged
        btnSave1.Enabled = True
    End Sub

    Private Sub tbEnumber1_TextChanged(sender As Object, e As EventArgs) Handles tbEnumber1.TextChanged
        btnSave1.Enabled = True
    End Sub
End Class

Public Class DataAccess
    Public Shared Function LoadData()
        Dim cl As New List(Of Contact)
        If My.Computer.FileSystem.FileExists("data") Then
            Dim fs As Stream = New FileStream("data", FileMode.Open)
            Dim bf As BinaryFormatter = New BinaryFormatter()
            cl = DirectCast(bf.Deserialize(fs), List(Of Contact))
            fs.Close()
        End If
        Return cl
    End Function

    Public Shared Function Save(contacts As List(Of Contact))
        If My.Computer.FileSystem.FileExists("data") = True Then
            My.Computer.FileSystem.DeleteFile("data")
        End If
        Dim fs As Stream = New FileStream("data", FileMode.Create)
        Dim bf As BinaryFormatter = New BinaryFormatter()
        bf.Serialize(fs, contacts)
        fs.Close()
        Return True
    End Function
End Class

<Serializable()> Public Class Contact
    Public name As String
    Public phoneNum As String
    Public street As String
    Public city As String
    Public state As String
    Public zipCode As String
    Public emailAdd As String
    Public employeeNum As Integer

    Sub New(n As String, p As String, s As String, c As String, st As String, z As String, e As String, emploNum As Int64)
        name = n
        phoneNum = p
        street = s
        city = c
        state = st
        zipCode = z
        emailAdd = e
        employeeNum = emploNum
    End Sub

End Class
