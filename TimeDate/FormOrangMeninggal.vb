Public Class FormOrangMeninggal

    'Perhitungan Selamatan Orang Meninggal
    '################################################################################################
    'Pembuat            : Vincentius Andri Kurnianto - VincentAndriK
    'Versi              : 0.0.1.201701
    'Tanggal Pembuatan  : 28 Januari 2017
    'Bahasa             : Bahasa Indonesia
    '################################################################################################
    'Perhitungan otomatis sesudah tombol Hitung! ditekan
    'Penyalinan Data langsung ke Clipboard
    'Penyimpanan Data dengan ekstensi tertentu (*.doc, *.rtf, *.txt, *.*)
    'Pencetakan Data otomatis dapat berjalan
    'Peringatan saat data yang akan diproses tidak ada
    'Peringatan saat akan menutup program
    'Dilengkapi dengan Dark Theme dan Light Theme
    'Pemberian Nama Hari dan Pasaran sedang masuk dalam tahap uji coba
    '################################################################################################
    'Kritik dan Saran : Kirim melalui Line @VincentAndriK atau PM Facebook di vincent.andri.k

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = "Hari Meninggal : " + Tanggal(DateTimePicker1.Value, 0) + vbCrLf +
            "3 hari : " + Tanggal(DateTimePicker1.Value, 2) + vbCrLf +
            "7 hari : " + Tanggal(DateTimePicker1.Value, 6) + vbCrLf +
            "40 hari : " + Tanggal(DateTimePicker1.Value, 39) + vbCrLf +
            "100 hari : " + Tanggal(DateTimePicker1.Value, 99) + vbCrLf +
            "1 tahun : " + TanggalKhusus(DateTimePicker1.Value, 1, -12) + vbCrLf +
            "2 tahun : " + TanggalKhusus(DateTimePicker1.Value, 2, -22) + vbCrLf +
            "1000 hari : " + Tanggal(DateTimePicker1.Value, 999)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Maaf, tidak ada data yang akan disalin", "Peringatan")
            DateTimePicker1.Focus()
        Else
            My.Computer.Clipboard.SetText(TextBox1.Text)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Maaf, tidak ada data yang akan disimpan", "Peringatan")
            DateTimePicker1.Focus()
        Else
            SaveFileDialog1.Filter = "Text Files (*.txt)|*.txt|Word Document (*.doc)|*.doc|Rich Text Files (*.rtf)|*.rtf|All Files (*.*)|*."
            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                IO.File.WriteAllText(SaveFileDialog1.FileName, TextBox1.Text + vbCrLf + vbCrLf + _
                                     "Waktu Penyimpanan : " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss") + _
                                     vbCrLf + "Nama Penyimpan : " + My.Computer.Name)
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Clear()
        DateTimePicker1.Focus()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        MessageBox.Show("Versi : 0.0.1.201701" + vbCrLf + "Pembuat : Vincentius Andri Kurnianto" + _
                        vbCrLf + "Terima kasih kepada : " + _
                        vbCrLf + "Bapak Stephanus Suyatno yang memberikan cara perhitungannya" + _
                        vbCrLf + "Tutorial di Internet (khususnya MSDN dan StackOverFlow)", "Tentang",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub FormOrangMeninggal_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If MessageBox.Show("Apakah Anda yakin akan keluar dari program ini?", "Keluar", _
         MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            MessageBox.Show("Terima kasih Anda sudah menggunakan program ini!", "Keluar",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            e.Cancel = True
        End If

    End Sub

    Private Sub FormOrangMeninggal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Color.White
        Button1.BackColor = Color.White
        Button2.BackColor = Color.White
        Button3.BackColor = Color.White
        Button4.BackColor = Color.White
        Button6.BackColor = Color.White
        Button7.BackColor = Color.White
        Button8.BackColor = Color.White
        Me.ForeColor = Color.Black
        TextBox1.BackColor = Color.White
        TextBox1.ForeColor = Color.Black
        Button7.Text = "Dark Theme"
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If Button7.Text = "Dark Theme" Then
            Me.BackColor = Color.Black
            Button1.BackColor = Color.Black
            Button2.BackColor = Color.Black
            Button3.BackColor = Color.Black
            Button4.BackColor = Color.Black
            Button6.BackColor = Color.Black
            Button7.BackColor = Color.Black
            Button8.BackColor = Color.Black
            Me.ForeColor = Color.White
            TextBox1.BackColor = Color.Black
            TextBox1.ForeColor = Color.White
            Button7.Text = "Light Theme"
        Else
            Me.BackColor = Color.White
            Button1.BackColor = Color.White
            Button2.BackColor = Color.White
            Button3.BackColor = Color.White
            Button4.BackColor = Color.White
            Button6.BackColor = Color.White
            Button7.BackColor = Color.White
            Button8.BackColor = Color.White
            Me.ForeColor = Color.Black
            TextBox1.BackColor = Color.White
            TextBox1.ForeColor = Color.Black
            Button7.Text = "Dark Theme"
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Maaf, tidak ada data yang akan dicetak", "Peringatan")
            DateTimePicker1.Focus()
        Else
            PrintDocument1.DefaultPageSettings.Margins = New Printing.Margins(100, 100, 100, 100)
            PrintDialog1.Document = PrintDocument1
            PrintDialog1.AllowSomePages = True
            If PrintDialog1.ShowDialog = DialogResult.OK Then
                PrintDocument1.Print()
            End If
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Static CurCharInt As Int32
        'SET THE FONT THAT YOU WANT TO USE FOR PRINTING.
        Dim fnt As New Font("Times New Roman", 20)
        Dim printheightInt, printwidthInt, leftmargin, topmargin As Int32
        With PrintDocument1.DefaultPageSettings
            ' SET THE VARIABLE THAT HOLD THE LIMITS OF THE PRINTING AREA RECTANGLE 
            printheightInt = .PaperSize.Height - .Margins.Top - .Margins.Bottom
            printwidthInt = .PaperSize.Width - .Margins.Left - .Margins.Right

            ' SET THE VARIABLES TO HOLD THE VALUES OF THE MARGIN THAT WILL WORK FOR 
            'THE X AND Y COORDINATES OF THE UPPER LEFT CORNER OF THE PRINTING AREA RECTANGLE 
            leftmargin = .Margins.Left ' X coordinate
            topmargin = .Margins.Top ' Y coordinate
        End With

        ' COMPUTE THE TOTAL LINES IN THE DOCUMENT BASED ON THE HEIGHT OF THE FONT AND ITS PRINTING AREA.
        Dim linecountInt As Int32 = CInt(printheightInt / fnt.Height)

        ' SET THE RECTANGLE STRUCTURE THAT SERVES AS THE PRINTING AREA.
        Dim printAreaRec As New RectangleF(leftmargin, topmargin, printwidthInt, printheightInt)

        ' INSTANTIATE THE STRINGFORMAT CLASS THAT CONTAINS TEXT LAYOUT INFORMATION, 
        'DISPLAY MANIPULATIONS AND OPENTYPE FEATURES. 
        Dim format As New StringFormat(StringFormatFlags.LineLimit)

        Dim linesfilledInt As Int32 'MUST BE PASSED WHEN PASSING CHARFIT.

        Dim charsfittedInt As Int32 'USED WHEN CALCULATING CurCharInt AND HasMorePages.

        'CALL MEASURESTRING TO KNOW HOW MANY CHARACTERS THAT WILL FIT IN THE PRINTING AREA RECTANGLE.
        e.Graphics.MeasureString(Mid(TextBox1.Text + vbCrLf + vbCrLf + _
                                     "Waktu Pencetakan : " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss"),
                                     CurCharInt + 1), fnt, _
                    New SizeF(printwidthInt, printheightInt), format, _
                    charsfittedInt, linesfilledInt)

        'IN THIS AREA, THE TEXT WILL BE PRINT IN THE PAGE
        e.Graphics.DrawString(Mid(TextBox1.Text + vbCrLf + vbCrLf + _
                                     "Waktu Pencetakan : " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss"),
                                     CurCharInt + 1), fnt, _
            Brushes.Black, printAreaRec, format)

        'FORMULA FOR ADVANCING THE CURRENT TO THE LAST CHARACTER PRINTED ON THIS PAGE.
        CurCharInt += charsfittedInt

        'CHECKING IF WHETHER THE PRINTING MODULE SHOULD BE FIRE TO ANOTHER PRINTPAGE EVENT
        If CurCharInt < TextBox1.Text.Length Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            'RESET CurCharInt AS IT IS STATIC.
            CurCharInt = 0
        End If
    End Sub

    Function Tanggal(ByVal Tgl As Date, ByVal Num As Integer) As String
        Tanggal = NamaHari(Tgl.AddDays(Num)) + " " + NamaPasaran(Tgl.AddDays(Num)) + ", " + Tgl.AddDays(Num).ToString("dd") + " " + _
            NamaBulan(Tgl.AddDays(Num)) + " " + Tgl.AddDays(Num).ToString("yyyy")
    End Function


    Function TanggalKhusus(ByVal Tgl As Date, ByVal Num1 As Integer, ByVal Num2 As Integer) As String
        TanggalKhusus = NamaHari(Tgl.AddYears(Num1).AddDays(Num2)) + " " + NamaPasaran(Tgl.AddYears(Num1).AddDays(Num2)) + ", " + Tgl.AddYears(Num1).AddDays(Num2).ToString("dd") + " " + _
            NamaBulan(Tgl.AddYears(Num1).AddDays(Num2)) + " " + Tgl.AddYears(Num1).AddDays(Num2).ToString("yyyy")
    End Function

    Function NamaHari(ByVal Tgl As Date) As String
        NamaHari = Choose(Weekday(Tgl), "Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jum'at", "Sabtu")
    End Function

    Function NamaBulan(ByRef Tgl As Date) As String
        NamaBulan = Choose(Month(Tgl), "Januari", "Februari", "Maret", "April", "Mei", "Juni", _
                           "Juli", "Agustus", "September", "Oktober", "November", "Desember")
    End Function

    Function NamaPasaran(ByVal Tgl As Date) As String
        Dim l
        Dim s
        Dim InitialDate As Date

        InitialDate = DateValue("02/01/1970")
        l = DateDiff("s", InitialDate, Tgl) * 1000
        s = l + 86400000
        s = s / 432000000
        s = Math.Round((s - Int(s)) * 10) / 2
        l = Math.Abs(Math.Round(s))
        If l > 4 Then l = 0
        NamaPasaran = Choose(l + 1, "Wage", "Kliwon", "Legi", "Pahing", "Pon")
    End Function

End Class