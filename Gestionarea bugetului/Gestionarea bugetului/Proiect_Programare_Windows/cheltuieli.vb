Imports System.Data.OleDb
Imports System.Windows.Forms.DataVisualization.Charting

Public Class cheltuieli
    Dim sumaCheltuita As Double = 0

    Private Sub cheltuieli_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UpdateSumaLabelFromDatabase()
        Me.BackColor = Color.FromArgb(71, 70, 71)
        Label2.Font = New Font("Arial", 20, FontStyle.Bold)
        Label1.Font = New Font("Arial", 20, FontStyle.Bold)
        Label3.Font = New Font("arial", 15, FontStyle.Bold)
        Label4.Font = New Font("arial", 15, FontStyle.Bold)
        Label5.Font = New Font("arial", 15, FontStyle.Bold)
        PictureBox1.BackColor = Color.FromArgb(255, 194, 14)
        ListBox2.Visible = False
        PictureBox2.BackColor = Color.FromArgb(255, 194, 14)
        ListBox1.Font = New Font("Arial", 15, FontStyle.Bold)
        ListBox1.BackColor = Color.FromArgb(255, 194, 14)
        ListBox1.ForeColor = Color.White
        venituri.Visible = False
        Form1.Visible = False
        Chart1.Visible = False
        Me.Show()
    End Sub

    Private Sub LoadDataFromDatabase()
        ListBox1.Items.Clear()
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\OneDrive\Desktop\Gestionarea bugetului\DataBase\cheltuieli.accdb"

        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Dim commandText As String = "SELECT Nume_produs, Pret, domeniu FROM Table1"
            Dim command As New OleDbCommand(commandText, connection)
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                Dim numeProdus As String = If(Not reader.IsDBNull(reader.GetOrdinal("Nume_produs")), reader("Nume_produs").ToString(), String.Empty)
                Dim pret As Double = If(Not reader.IsDBNull(reader.GetOrdinal("Pret")), Convert.ToDouble(reader("Pret")), 0.0)
                Dim domeniu As String = If(Not reader.IsDBNull(reader.GetOrdinal("domeniu")), reader("domeniu").ToString(), String.Empty)

                ListBox1.Items.Add(numeProdus & " - " & pret.ToString())
            End While
        End Using
    End Sub

    Private Sub UpdateSumaLabelFromDatabase()
        LoadDataFromDatabase()

        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\OneDrive\Desktop\Gestionarea bugetului\DataBase\cheltuieli.accdb"
        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Dim commandText As String = "SELECT SUM(Pret) FROM Table1"
            Dim command As New OleDbCommand(commandText, connection)
            Dim result As Object = command.ExecuteScalar()
            If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                sumaCheltuita = Convert.ToDouble(result)
                UpdateSumaLabel()
            End If
        End Using
    End Sub

    Private Sub UpdateSumaLabel()
        Label2.Text = sumaCheltuita.ToString()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim valoareCheltuita As Double
        Dim categorie As String = TextBox3.Text.Trim()
        Dim numeProdus As String = TextBox1.Text.Trim()

        If Double.TryParse(TextBox2.Text, valoareCheltuita) Then
            sumaCheltuita += valoareCheltuita
            UpdateSumaLabel()

            ListBox1.Items.Add(numeProdus & " - " & valoareCheltuita.ToString())

            AddDataToDatabase(numeProdus, valoareCheltuita, categorie)

            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
        Else
            MessageBox.Show("Introduceți o valoare numerică validă pentru preț.")
        End If
    End Sub

    Private Sub AddDataToDatabase(ByVal numeProdus As String, ByVal pret As Double, ByVal categorie As String)
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\OneDrive\Desktop\Gestionarea bugetului\DataBase\cheltuieli.accdb"

        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Dim commandText As String = "INSERT INTO Table1 (Nume_produs, Pret, domeniu) VALUES (?, ?, ?)"
            Dim command As New OleDbCommand(commandText, connection)
            command.Parameters.AddWithValue("@Nume_produs", numeProdus)
            command.Parameters.AddWithValue("@Pret", pret)
            command.Parameters.AddWithValue("@domeniu", categorie)
            command.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ListBox2.Visible = True
        ListBox2.Items.Clear()
        ListBox2.Items.Add("Cheltuieli")
        ListBox2.Items.Add("Venituri")
        ListBox2.Font = New Font("Arial", 10, FontStyle.Bold)
        ListBox2.ForeColor = Color.White
        ListBox2.BackColor = Color.FromArgb(71, 70, 70)
        venituri.Visible = False
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        If ListBox2.SelectedIndex <> -1 Then
            Select Case ListBox2.SelectedIndex
                Case 0
                    Dim cheltuieliForm As New cheltuieli()
                    Me.Show()
                Case 1
                    Dim venituriForm As New venituri()
                    Me.Hide()
                    venituri.Show()
            End Select
        End If
    End Sub

    Private Sub TextBox_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged
        TransferData()
    End Sub

    Private Sub TransferData()
        Dim prices As New List(Of Double)()
        Dim domains As New List(Of String)()
        domains.Add(TextBox3.Text)
    End Sub
End Class
