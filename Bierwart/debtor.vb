Public Class Debtor


    Private IdentNummer As String
    Private KomplettName As String

    Private UnterNummer As String
    Private BierKaution As String
    Private Rechnungsnummer As String
    Private Datum As String
    Private Fassbier As String
    Private Flaschenbier As String
    Private Bierjunge As String
    Private Fassbrause As String
    Private Softgetränke As String
    Private Wasser As String
    Private Fassspenden As String
    Private Sonstiges As String
    Private Bezahlt As String
    Private Restschuld As String

    Public Sub New(ByVal ID As String, ByVal KName As String, ByVal UNummer As String)

        IdentNummer = ID
        KomplettName = KName
        UnterNummer = UNummer

    End Sub

    Public Function Output(index As Integer) As String
        Select Case index

            Case 0
                Return IdentNummer
            Case 1
                Return KomplettName
            Case 2
                Return UnterNummer
            Case 3
                Return BierKaution
            Case 4
                Return Rechnungsnummer
            Case 5
                Return Datum
            Case 6
                Return Fassbier
            Case 7
                Return Flaschenbier
            Case 8
                Return Bierjunge
            Case 9
                Return Fassbrause
            Case 10
                Return Softgetränke
            Case 11
                Return Wasser
            Case 12
                Return Fassspenden
            Case 13
                Return Sonstiges
            Case 14
                Return Bezahlt
            Case 15
                Return Restschuld





        End Select


    End Function




End Class
