Public Class Form1

    Dim kodeBarang() As String = {"BRG001", "BRG002", "BRG003"}
    Dim namaBarang() As String = {"Laptop", "Mouse", "Keyboard"}
    Dim satuanBarang() As String = {"Unit", "Buah", "Buah"}
    Dim stok() As Integer = {10, 15, 8}

    Dim idTransaksi(99) As String
    Dim barangTransaksi(99) As Integer
    Dim jumlahTransaksi(99) As Integer
    Dim namaPeminjam(99) As String
    Dim kategoriTransaksi(99) As String
    Dim satuanTransaksi(99) As String
    Dim statusTransaksi(99) As String

    Dim tanggalPinjam(99) As String
    Dim tanggalKembali(99) As String

    Dim banyakTransaksi As Integer = 0
    Dim nomorTransaksi As Integer = 1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Items.Clear()
        For i As Integer = 0 To kodeBarang.Length - 1
            ComboBox1.Items.Add(kodeBarang(i))
        Next

        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        ComboBox2.DropDownStyle = ComboBoxStyle.DropDownList

        TextBox2.ReadOnly = True
        TextBox2.Clear()

        NumericUpDown1.Minimum = 1
        NumericUpDown1.Maximum = 100
        NumericUpDown1.Value = 1

        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("KodeBarang", "Kode Barang")
        DataGridView1.Columns.Add("NamaBarang", "Nama Barang")
        DataGridView1.Columns.Add("Satuan", "Satuan")
        DataGridView1.Columns.Add("Stok", "Stok")

        DataGridView2.Columns.Clear()
        DataGridView2.Columns.Add("ID", "ID Transaksi")
        DataGridView2.Columns.Add("KodeBarang", "Kode Barang")
        DataGridView2.Columns.Add("Peminjam", "Nama Peminjam")
        DataGridView2.Columns.Add("NamaBarang", "Nama Barang")
        DataGridView2.Columns.Add("Satuan", "Satuan")
        DataGridView2.Columns.Add("Kategori", "Kategori")
        DataGridView2.Columns.Add("Jumlah", "Jumlah")
        DataGridView2.Columns.Add("TanggalPinjam", "Tanggal Pinjam")
        DataGridView2.Columns.Add("TanggalKembali", "Tanggal Kembali")
        DataGridView2.Columns.Add("Status", "Status")

        DataGridView1.ReadOnly = True
        DataGridView2.ReadOnly = True

        DataGridView1.AllowUserToAddRows = False
        DataGridView2.AllowUserToAddRows = False

        TampilkanStok()
        TampilkanTransaksi()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(
        sender As Object,
        e As EventArgs
    ) Handles ComboBox1.SelectedIndexChanged

        Dim indeks As Integer = ComboBox1.SelectedIndex

        If indeks = -1 Then
            TextBox2.Clear()
            Return
        End If

        TextBox2.Text = namaBarang(indeks)

    End Sub

    Private Sub TampilkanStok()

        DataGridView1.Rows.Clear()

        For i As Integer = 0 To namaBarang.Length - 1
            DataGridView1.Rows.Add(
                kodeBarang(i),
                namaBarang(i),
                satuanBarang(i),
                stok(i)
            )
        Next

    End Sub

    Private Sub TampilkanTransaksi()

        DataGridView2.Rows.Clear()

        For i As Integer = 0 To banyakTransaksi - 1

            Dim indeksBarang As Integer = barangTransaksi(i)

            DataGridView2.Rows.Add(
                idTransaksi(i),
                kodeBarang(indeksBarang),
                namaPeminjam(i),
                namaBarang(indeksBarang),
                satuanTransaksi(i),
                kategoriTransaksi(i),
                jumlahTransaksi(i),
                tanggalPinjam(i),
                tanggalKembali(i),
                statusTransaksi(i)
            )

        Next

        PerbaruiPilihanTransaksi()

    End Sub

    Private Sub PerbaruiPilihanTransaksi()

        Dim idTerpilih As String = ComboBox2.Text

        ComboBox2.Items.Clear()

        For i As Integer = 0 To banyakTransaksi - 1

            If statusTransaksi(i) = "Dipinjam" Then
                ComboBox2.Items.Add(idTransaksi(i))
            End If

        Next

        If ComboBox2.Items.Contains(idTerpilih) Then
            ComboBox2.SelectedItem = idTerpilih
        ElseIf ComboBox2.Items.Count > 0 Then
            ComboBox2.SelectedIndex = 0
        End If

    End Sub

    Private Sub Button1_Click(
        sender As Object,
        e As EventArgs
    ) Handles Button1.Click

        If TextBox1.Text.Trim() = "" OrElse ComboBox1.SelectedIndex = -1 Then
            MessageBox.Show("Isi nama peminjam dan pilih kode barang!")
            Return
        End If

        If TextBox3.Text.Trim() = "" Then
            MessageBox.Show("Isi kategori barang!")
            Return
        End If

        If banyakTransaksi >= 100 Then
            MessageBox.Show("Batas transaksi sudah tercapai!")
            Return
        End If

        Dim pilihanBarang As Integer = ComboBox1.SelectedIndex
        Dim jumlah As Integer = CInt(NumericUpDown1.Value)

        If stok(pilihanBarang) < jumlah Then
            MessageBox.Show("Stok barang tidak mencukupi!")
            Return
        End If

        idTransaksi(banyakTransaksi) = "PJM-" & nomorTransaksi.ToString("000")
        barangTransaksi(banyakTransaksi) = pilihanBarang
        jumlahTransaksi(banyakTransaksi) = jumlah
        namaPeminjam(banyakTransaksi) = TextBox1.Text.Trim()
        kategoriTransaksi(banyakTransaksi) = TextBox3.Text.Trim()
        satuanTransaksi(banyakTransaksi) = satuanBarang(pilihanBarang)
        tanggalPinjam(banyakTransaksi) = DateTimePicker1.Value.ToString("dd-MM-yyyy")
        tanggalKembali(banyakTransaksi) = "-"
        statusTransaksi(banyakTransaksi) = "Dipinjam"

        stok(pilihanBarang) = stok(pilihanBarang) - jumlah

        banyakTransaksi += 1
        nomorTransaksi += 1

        TampilkanStok()
        TampilkanTransaksi()

        TextBox1.Clear()
        ComboBox1.SelectedIndex = -1
        TextBox2.Clear()
        TextBox3.Clear()
        NumericUpDown1.Value = 1

        MessageBox.Show("Peminjaman berhasil!")

    End Sub

    Private Sub Button2_Click(
        sender As Object,
        e As EventArgs
    ) Handles Button2.Click

        If ComboBox2.SelectedIndex = -1 Then
            MessageBox.Show("Pilih transaksi yang ingin dikembalikan!")
            Return
        End If

        Dim idPilihan As String = ComboBox2.Text

        For i As Integer = 0 To banyakTransaksi - 1

            If idTransaksi(i) = idPilihan Then

                If statusTransaksi(i) = "Dikembalikan" Then
                    MessageBox.Show("Barang sudah dikembalikan!")
                    Return
                End If

                Dim pilihanBarang As Integer = barangTransaksi(i)

                stok(pilihanBarang) = stok(pilihanBarang) + jumlahTransaksi(i)
                tanggalKembali(i) = DateTimePicker2.Value.ToString("dd-MM-yyyy")
                statusTransaksi(i) = "Dikembalikan"

                TampilkanStok()
                TampilkanTransaksi()

                MessageBox.Show("Barang berhasil dikembalikan!")
                Return

            End If

        Next

        MessageBox.Show("Transaksi tidak ditemukan!")

    End Sub

End Class
