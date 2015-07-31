
' ; Seperator
'
'
'
'
'
'
'
Public Class Form1

    Dim DataImport As String
    Dim DataImportSplit(12) As String

    Dim Person(50, 3) As String 'ID Name Kaution
    Dim PersonPointer As Integer 'Counter im Personarray

    Dim Konto(1000, 4) As String ' Datum Bezeichnung Konto Betrag Saldo
    Dim KontoPointer As Integer 'Counter im Kontoarray

    Dim BarKonto(1000, 4) As String ' Datum Bezeichnung Konto Betrag Saldo
    Dim BarKontoPointer As Integer 'Counter im BarKontoarray

    Dim Tresenrechnung(1000, 11) As String 'Nummer Datum Name Fassbier Flaschenbier Bierjunge Fassbrause Softgetränke Wasser Fassspenden Sonstiges Gesamt
    Dim TresenrechnungPointer As Integer 'Counter im Kontoarray

    Dim Lieferabwicklung(500, 11) As String 'Nummer Datum Name Fassbier Flaschenbier Bierjunge Fassbrause Softgetränke Wasser Fassspenden Sonstiges Gesamt
    Dim LieferabwicklungPointer As Integer 'Counter im Kontoarray

    Dim GetränkeWert(7) As Single 'Fassbier 1 Flaschenbier 1 Bierjunge 1,5 Fassbrause 0,5 Softgetränke 1 Wasser 0,5 Fassspenden 1 Sonstiges 1

    Dim HighID As Integer
    Dim count As Integer

    Private Sub BeendenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BeendenToolStripMenuItem.Click

        Me.Close()

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TabControl1.SelectedTab = TabPage5

        LoadIn() ' Einladen aller Files
        LoadIn2() ' Load in in die Listviews

        EXP()

    End Sub
    Private Sub LoadIn()

        Dim PersonFile As New System.IO.StreamReader("Person.csv")

        PersonPointer = 0

        Do While PersonFile.Peek <> -1

            DataImport = PersonFile.ReadLine
            If DataImport <> "" Then


                DataImportSplit = DataImport.Split(";")
                Person(PersonPointer, 0) = DataImportSplit(0) ' ID
                Person(PersonPointer, 1) = DataImportSplit(1) ' Name
                Person(PersonPointer, 2) = DataImportSplit(2) ' Kaution
                PersonPointer += 1

            End If
        Loop
        PersonFile.Close()
        PersonFile = Nothing

        Dim KontoFile As New System.IO.StreamReader("Kontobuch.csv")

        KontoPointer = 0

        Do While KontoFile.Peek <> -1

            DataImport = KontoFile.ReadLine
            If DataImport <> "" Then


                DataImportSplit = DataImport.Split(";")
                Konto(KontoPointer, 0) = DataImportSplit(0) ' Datum
                Konto(KontoPointer, 1) = DataImportSplit(1) ' Beschreibung
                Konto(KontoPointer, 2) = DataImportSplit(2) ' Gegenkonto
                Konto(KontoPointer, 3) = DataImportSplit(3) ' Betrag
                Konto(KontoPointer, 4) = DataImportSplit(4) ' Saldo
                KontoPointer += 1
            End If

        Loop
        KontoFile.Close()
        KontoFile = Nothing

        Dim TresenrechnungFile As New System.IO.StreamReader("Tresenrechnung.csv")

        TresenrechnungPointer = 0

        Do While TresenrechnungFile.Peek <> -1

            DataImport = TresenrechnungFile.ReadLine
            If DataImport <> "" Then

                DataImportSplit = DataImport.Split(";")
                Tresenrechnung(TresenrechnungPointer, 0) = DataImportSplit(0) ' Nummer
                Tresenrechnung(TresenrechnungPointer, 1) = DataImportSplit(1) ' Datum
                Tresenrechnung(TresenrechnungPointer, 2) = DataImportSplit(2) ' Name
                Tresenrechnung(TresenrechnungPointer, 3) = DataImportSplit(3) ' Fassbier
                Tresenrechnung(TresenrechnungPointer, 4) = DataImportSplit(4) ' Flaschenbier
                Tresenrechnung(TresenrechnungPointer, 5) = DataImportSplit(5) ' Bierjunge
                Tresenrechnung(TresenrechnungPointer, 6) = DataImportSplit(6) ' Fassbrause
                Tresenrechnung(TresenrechnungPointer, 7) = DataImportSplit(7) ' Softgetränke
                Tresenrechnung(TresenrechnungPointer, 8) = DataImportSplit(8) ' Wasser
                Tresenrechnung(TresenrechnungPointer, 9) = DataImportSplit(9) ' Fassspenden
                Tresenrechnung(TresenrechnungPointer, 10) = DataImportSplit(10) ' Sonstiges
                Tresenrechnung(TresenrechnungPointer, 11) = DataImportSplit(11) ' Gesamt
                TresenrechnungPointer += 1
            End If

        Loop
        TresenrechnungFile.Close()
        TresenrechnungFile = Nothing


        Dim BarKontoFile As New System.IO.StreamReader("BarKontoBuch.csv")

        BarKontoPointer = 0

        Do While BarKontoFile.Peek <> -1

            DataImport = BarKontoFile.ReadLine
            If DataImport <> "" Then


                DataImportSplit = DataImport.Split(";")
                BarKonto(BarKontoPointer, 0) = DataImportSplit(0) ' Datum
                BarKonto(BarKontoPointer, 1) = DataImportSplit(1) ' Beschreibung
                BarKonto(BarKontoPointer, 2) = DataImportSplit(2) ' Gegenkonto
                BarKonto(BarKontoPointer, 3) = DataImportSplit(3) ' Betrag
                BarKonto(BarKontoPointer, 4) = DataImportSplit(4) ' Saldo
                BarKontoPointer += 1
            End If

        Loop
        BarKontoFile.Close()
        BarKontoFile = Nothing

       

    End Sub
    Private Sub LoadIn2()

        ListView1.Items.Clear()

        For i = 0 To 50
            If Person(i, 0) <> "" Then
                ListView1.Items.Add(Person(i, 0))
                ListView1.Items(i).SubItems.Add(Person(i, 1))
                ListView1.Items(i).SubItems.Add(Person(i, 2))
            End If
        Next


        ListView2.Items.Clear()

        For i = 0 To 1000
            If Konto(i, 0) <> "" Then
                ListView2.Items.Add(Konto(i, 0))
                ListView2.Items(i).SubItems.Add(Konto(i, 1))
                ListView2.Items(i).SubItems.Add(Konto(i, 2))
                ListView2.Items(i).SubItems.Add(Konto(i, 3))
                ListView2.Items(i).SubItems.Add(Konto(i, 4))
            End If
        Next

        ListView3.Items.Clear()

        For i = 0 To 1000
            If Tresenrechnung(i, 0) <> "" Then
                ListView3.Items.Add(Tresenrechnung(i, 0))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 1))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 2))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 3))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 4))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 5))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 6))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 7))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 8))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 9))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 10))
                ListView3.Items(i).SubItems.Add(Tresenrechnung(i, 11))
            End If
        Next

        ListView4.Items.Clear()

        For i = 0 To 1000
            If BarKonto(i, 0) <> "" Then
                ListView4.Items.Add(BarKonto(i, 0))
                ListView4.Items(i).SubItems.Add(BarKonto(i, 1))
                ListView4.Items(i).SubItems.Add(BarKonto(i, 2))
                ListView4.Items(i).SubItems.Add(BarKonto(i, 3))
                ListView4.Items(i).SubItems.Add(BarKonto(i, 4))
            End If
        Next

    End Sub

    Private Sub BestandsdatenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BestandsdatenToolStripMenuItem.Click

        TabControl1.SelectedTab = TabPage4

    End Sub
    Private Sub GiroToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GiroToolStripMenuItem.Click

        TabControl1.SelectedTab = TabPage3
        ListBox1.Items.Clear()

        For i = 0 To 50
            If Person(i, 0) <> "" Then
                ListBox1.Items.Add(Person(i, 1))
            End If
        Next


    End Sub
    Private Sub RechnungToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RechnungToolStripMenuItem.Click

        TabControl1.SelectedTab = TabPage1

        ListBox2.Items.Clear()

        For i = 0 To 50
            If Person(i, 0) <> "" Then
                ListBox2.Items.Add(Person(i, 1))
            End If
        Next

    End Sub
    Private Sub BarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BarToolStripMenuItem.Click
        TabControl1.SelectedTab = TabPage2

        ListBox3.Items.Clear()

        For i = 0 To 50
            If Person(i, 0) <> "" Then
                ListBox3.Items.Add(Person(i, 1))
            End If
        Next


    End Sub
    Private Sub LieferscheineToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LieferscheineToolStripMenuItem.Click
        TabControl1.SelectedTab = TabPage6
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'ID Check
        For i = 0 To 51
            For j = 0 To 50
                If Person(j, 0) <> "" Then


                    If i = CInt(Person(j, 0)) Then

                        HighID = i + 1

                    End If
                End If
            Next
        Next

        Person(PersonPointer, 0) = CStr(HighID)
        Person(PersonPointer, 1) = TextBox1.Text & " " & TextBox2.Text

        If CheckBox1.Checked = True Then
            Person(PersonPointer, 2) = "50"
        End If
        If CheckBox1.Checked = False Then
            Person(PersonPointer, 2) = "0"
        End If

        TextBox1.Text = ""
        TextBox2.Text = ""
        CheckBox1.Checked = False

        PersonPointer += 1

        LoadIn2()

        MsgBox("Abgeschlossen")

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)

        For i = 0 To 50

            If CInt(TextBox3.Text) = Person(i, 0) Then

                Person(i, 0) = ""
                Person(i, 1) = ""
                Person(i, 2) = ""
                Person(i, 3) = ""

            End If

        Next

        TextBox3.Text = ""

        LoadIn2()
        MsgBox("Abgeschlossen")


    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)

        Konto(KontoPointer, 0) = DateTimePicker1.Value
        Konto(KontoPointer, 1) = TextBox4.Text
        Konto(KontoPointer, 2) = ListBox1.SelectedItem
        Konto(KontoPointer, 3) = TextBox5.Text
        If KontoPointer = 0 Then
            Konto(KontoPointer, 4) = CSng(TextBox5.Text)
        End If
        If KontoPointer <> 0 Then
            Konto(KontoPointer, 4) = CSng(Konto((KontoPointer - 1), 4)) + CSng(TextBox5.Text)
        End If

        KontoPointer += 1
        LoadIn2()

        MsgBox("Abgeschlossen")

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Nummer Datum Name Fassbier Fassbrause Softgetränke Wasser Fassspenden Sonstiges Gesamt

        Tresenrechnung(TresenrechnungPointer, 0) = TextBox6.Text
        Tresenrechnung(TresenrechnungPointer, 1) = DateTimePicker3.Value
        Tresenrechnung(TresenrechnungPointer, 2) = ListBox2.SelectedItem
        Tresenrechnung(TresenrechnungPointer, 3) = TextBox7.Text
        Tresenrechnung(TresenrechnungPointer, 4) = TextBox8.Text
        Tresenrechnung(TresenrechnungPointer, 5) = TextBox9.Text
        Tresenrechnung(TresenrechnungPointer, 6) = TextBox10.Text
        Tresenrechnung(TresenrechnungPointer, 7) = TextBox11.Text
        Tresenrechnung(TresenrechnungPointer, 8) = TextBox12.Text
        Tresenrechnung(TresenrechnungPointer, 9) = TextBox13.Text
        Tresenrechnung(TresenrechnungPointer, 10) = TextBox14.Text
        Tresenrechnung(TresenrechnungPointer, 11) = ((CInt(TextBox7.Text) * GetränkeWert(0)) + (CInt(TextBox8.Text) * GetränkeWert(1)) + (CInt(TextBox9.Text) * GetränkeWert(2)) + (CInt(TextBox10.Text) * GetränkeWert(3)) + (CInt(TextBox11.Text) * GetränkeWert(4)) + (CInt(TextBox12.Text) * GetränkeWert(5)) + (CInt(TextBox13.Text) * GetränkeWert(6)) + (CInt(TextBox14.Text) * GetränkeWert(7)))

        TresenrechnungPointer += 1

        LoadIn2()


        MsgBox("Abgeschlossen")

        TextBox7.Text = "0"
        TextBox8.Text = "0"
        TextBox9.Text = "0"
        TextBox10.Text = "0"
        TextBox11.Text = "0"
        TextBox12.Text = "0"
        TextBox13.Text = "0"
        TextBox14.Text = "0"



    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs)

        BarKonto(BarKontoPointer, 0) = DateTimePicker2.Value
        BarKonto(BarKontoPointer, 1) = TextBox16.Text
        BarKonto(BarKontoPointer, 2) = ListBox3.SelectedItem
        BarKonto(BarKontoPointer, 3) = TextBox15.Text
        If BarKontoPointer = 0 Then
            BarKonto(KontoPointer, 4) = CSng(TextBox15.Text)
        End If
        If BarKontoPointer <> 0 Then
            BarKonto(BarKontoPointer, 4) = CSng(BarKonto((BarKontoPointer - 1), 4)) + CSng(TextBox15.Text)
        End If

        BarKontoPointer += 1
        LoadIn2()

        MsgBox("Abgeschlossen")
    End Sub





    Private Sub EXP()
        'Fassbier 1 Flaschenbier 1 Bierjunge 1,5 Fassbrause 0,5 Softgetränke 1 Wasser 0,5 Fassspenden 1 Sonstiges 1

        GetränkeWert(0) = 1
        GetränkeWert(1) = 1
        GetränkeWert(2) = 1.5
        GetränkeWert(3) = 0.5
        GetränkeWert(4) = 1
        GetränkeWert(5) = 0.5
        GetränkeWert(6) = 1
        GetränkeWert(7) = 1


    End Sub




    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Nummer Datum Name Fassbier Flaschenbier Bierjunge Fassbrause Softgetränke Wasser Fassspenden Sonstiges Gesamt
        Lieferabwicklung(LieferabwicklungPointer, 0) = DateTimePicker4.Value





    End Sub


End Class
