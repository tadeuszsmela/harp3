
Module basAudio

    Public DJ_Studio_output_device As Short = 0
    Public dlugosc_bufora_audio As Integer = 1000

    Public okno_odtwarzacza As frmPlayer



    Public Function ustaw_parametry_pracy_playera(ByRef ap As AudioDjStudio.AudioDjStudio) As Integer
        Dim msg As String
        Dim wynik As Integer = 0



        Try
            ap.EnableSpeakers = True
            ap.EnableMixingFeatures = False
            ap.BufferLength = dlugosc_bufora_audio

        Catch ex As Exception
            msg = "Wystąpił problem podczas ustawiania parametrów pracy playera: " & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1
        End Try

        Return wynik

    End Function



    Public Function kopiuj_plik_audio(ByVal plik_zrodlowy As String, ByVal sciezka_zapisu As String) As Integer
        Dim n_pl_zr As String
        Dim znak As String
        Dim tmp_str As String
        Dim plik_cel As String
        Dim msg As String = ""
        Dim wynik As Integer = 0

        n_pl_zr = ustal_nazwe_pliku_audio(plik_zrodlowy)

        plik_cel = sciezka_zapisu

        If Right(plik_cel, 1) <> "\" Then
            plik_cel = plik_cel & "\"
        End If

        plik_cel = plik_cel & n_pl_zr

        Try
            System.IO.File.Copy(plik_zrodlowy, plik_cel, True)

        Catch ex As Exception
            msg = "Wystąpił problem podczas kopiowania pliku audio: "
            msg = msg & vbCrLf & ex.Message
            wskaznik_myszy(0)
            MessageBox.Show(msg)
            wynik = -1
        End Try

        Return wynik



    End Function

    Public Function ustal_nazwe_pliku_audio(ByVal pl_audio As String) As String

        'funkcja jako parametr dostaje peną nazwe pliku razem ze ścieżką
        'a zwraca tylko nazwę samego pliku

        Dim elementy_sciezki() As String
        Dim wynik As String = ""


        elementy_sciezki = Split(pl_audio, "\")
        If elementy_sciezki.Length > 0 Then
            wynik = elementy_sciezki(elementy_sciezki.Length - 1)
        End If

        Return wynik


    End Function
End Module
