Imports System.Data.OleDb

Public Class venituri
    Dim sumaCheltuita As Double = 0

    Private Sub venituri_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PictureBox2.BackColor = Color.FromArgb(255, 194, 14)
        PictureBox1.BackColor = Color.FromArgb(255, 194, 14)
        Label2.Font = New Font("Arial", 15, FontStyle.Bold)
        Label1.Font = New Font("Arial", 20, FontStyle.Bold)
        Label3.Font = New Font("arial", 15, FontStyle.Bold)
        ListBox1.BackColor = Color.FromArgb(255, 194, 14)
        Me.BackColor = Color.FromArgb(71, 70, 70)
        ListBox2.Visible = False
        cheltuieli.Hide()
        cheltuieli.Visible = False
        ListBox1.Font = New Font("Arial", 15, FontStyle.Bold)
        ListBox1.BackColor = Color.FromArgb(255, 194, 14)
        ListBox1.ForeColor = Color.White
        LoadDataFromDatabase()
    End Sub

    Private Sub LoadDataFromDatabase()
        ListBox1.Items.Clear()
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\OneDrive\Desktop\Gestionarea bugetului\DataBase\venit.accdb"

        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Dim commandText As String = "SELECT suma, provenienta FROM Table1"
            Dim command As New OleDbCommand(commandText, connection)
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                Dim suma As Double = Convert.ToDouble(reader("suma"))
                Dim provenienta As String = reader("provenienta").ToString()
                ListBox1.Items.Add(suma.ToString() & " - " & provenienta)

                sumaCheltuita += suma
            End While
        End Using


        UpdateSumaLabel()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ListBox2.Visible = True
        ListBox2.Visible = True
        ListBox2.Items.Clear()
        ListBox2.Items.Add("Cheltuieli")
        ListBox2.Items.Add("Venituri")
        ListBox2.Font = New Font("Arial", 10, FontStyle.Bold)
        ListBox2.ForeColor = Color.White
        ListBox2.BackColor = Color.FromArgb(71, 70, 70)
        cheltuieli.Hide()
    End Sub

    Private Sub UpdateSumaLabel()
        cheltuieli.Visible = False
        Label1.Text = sumaCheltuita.ToString()
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        If ListBox2.SelectedIndex <> -1 Then
            Select Case ListBox2.SelectedIndex
                Case 0
                    Dim cheltuieliForm As New cheltuieli()
                    Me.Hide()
                    cheltuieli.Show()
                Case 1
                    Dim venituriForm As New venituri()
                    Me.Show()
            End Select
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim valoareVenit As Double
        Dim categorie As String = TextBox2.Text.Trim()

        If Double.TryParse(TextBox1.Text, valoareVenit) Then
            sumaCheltuita += valoareVenit
            ListBox1.Items.Add(valoareVenit.ToString() & " - " & categorie)
            AddDataToDatabase(valoareVenit, categorie)
            TextBox1.Clear()
        End If

        TextBox2.Clear()

        UpdateSumaLabel()
    End Sub

    Private Sub AddDataToDatabase(ByVal suma As Double, ByVal provenienta As String)
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\OneDrive\Desktop\Gestionarea bugetului\DataBase\venit.accdb"

        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Dim commandText As String = "INSERT INTO Table1 (suma, provenienta) VALUES (?, ?)"
            Dim command As New OleDbCommand(commandText, connection)
            command.Parameters.AddWithValue("@suma", suma)
            command.Parameters.AddWithValue("@provenienta", provenienta)
            command.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class
