Module basControlPIK

    Public colListaRaportuControlPIK As New Collection


    Public Function zaladuj_raport_poemisyjny_ControlPIK(ByVal id_programu As Integer, ByVal d As Date, ByVal gp As String, ByVal gk As String, ByVal tryb As Integer) As Integer

        'tryb 0 - zwykły iomport pozycji
        '     1 - dołączenie 


        Dim a As Integer
        a = zaladuj_plik_raportu_ControPIK(id_programu, d, gp, gk, tryb)



    End Function



    Private Function zaladuj_dlugosc_muzyki_ControPIK(ByVal nr_programu As Integer, ByVal d As Date, ByVal gp As String, gk As String) As Integer
      

        Dim plk As String
        Dim dd As String
        Dim mm As String
        Dim yy As String
        Dim tmp_czas_startu As String
        Dim tmp_czas_emisji As String
        Dim tmp_tytul As String
        Dim tmp_plk_audio As String

        Dim wczytana_linia As String
        Dim koniec_pliku As Boolean = False
        '  Dim tmp_str As String
        Dim poz As Integer
        Dim rekord_prawidlowy As Boolean = False
        Dim r As clsPozycjaRaportuControlPIK
        Dim k As Integer = 0
        Dim msg As String
        Dim wynik As Integer = 0
        Dim k2 As Boolean
        Dim znak As String
        Dim p As clsProgram
        Dim kat_raportow As String = ""

        Dim gp_d As Date
        Dim gk_d As Date
        Dim tmp_czas_startu_d As Date

        Dim koniec_rekordu As Boolean
        Dim tmp_dlugosc As Integer = 0

        Try
            gp_d = CDate(gp)

        Catch ex As Exception

        End Try


        Try
            gk_d = CDate(gk)

        Catch ex As Exception

        End Try


        For Each p In colProgramy
            If p.id = nr_programu Then
                kat_raportow = p.parametry_polaczenia
                Exit For
            End If
        Next


        dd = Format(d, "dd")
        mm = Format(d, "MM")
        yy = Right(Year(d).ToString, 2)


        plk = kat_raportow
        If Right(plk, 1) <> "\" Then
            plk = plk & "\"
        End If

        plk = plk & "Muzyka\"
        plk = plk & yy & "-" & mm & "-" & dd & ".txt"

        If System.IO.File.Exists(plk) = False Then
            Return 0
        End If

        Dim enc As System.Text.Encoding
        enc = System.Text.Encoding.Default

        Dim fStreamReader As New System.IO.StreamReader(plk, enc, True)

        'Start emisji   :  01:00:11
        'Czas emisji    :  00:04:21
        'Czas utworu    :  00:04:20
        'Tytuł          :  Beautiful(Stranger)
        'Wykonawca      : MADONNA()
        'Kompozytor     :  MADONNA/ORBIT, W
        'Autor textu    :
        'Wydawca        : WB
        'Rok nagrania   :  1999
        'Kraj           :   Z
        '====================================

        Dim tmp_str As String
        Dim tmp_start As String
        Dim tmp_start_d As Date
        Dim tmp_stop_d As Date
        Dim ts As TimeSpan
        Dim tmp_sek As Integer

        Try
            Do

                koniec_rekordu = False
                tmp_start = "00:00:00"
                tmp_czas_emisji = "00:00:00"

                Do
                    wczytana_linia = fStreamReader.ReadLine()
                    If wczytana_linia = Nothing Then
                        koniec_pliku = True
                        koniec_rekordu = True
                    Else
                        tmp_str = Trim(UCase(wczytana_linia))
                        If Left(tmp_str, 12) = "START EMISJI" Then
                            tmp_start = Right(wczytana_linia, 8)
                        ElseIf Left(tmp_str, 11) = "CZAS EMISJI" Then
                            tmp_czas_emisji = Right(wczytana_linia, 8)

                        ElseIf Left(tmp_str, 1) = "=" Then
                            koniec_rekordu = True
                        End If
                    End If
                Loop While koniec_rekordu = False


                Try
                    tmp_start_d = CDate(tmp_start)
                Catch ex As Exception

                End Try


                If tmp_start_d > gp_d Then
                    If tmp_start_d < gk_d Then
                        tmp_sek = wyznacz_liczbe_sekund(tmp_czas_emisji)

                        tmp_dlugosc = tmp_dlugosc + tmp_sek
                    End If
                End If



            Loop While koniec_pliku = False

            wynik = tmp_dlugosc
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu pliku raportu poemisyjnego (wyznaczanie długości muzyki):"
            msg = msg & ex.Message
            MessageBox.Show(msg)
            wynik = -1
        End Try


        Try
            fStreamReader.Close()

        Catch ex As Exception

        End Try

        Return wynik



    End Function


    Private Function zaladuj_dlugosc_reklam_ControPIK(ByVal nr_programu As Integer, ByVal d As Date, ByVal gp As String, gk As String) As Integer


        Dim plk As String
        Dim dd As String
        Dim mm As String
        Dim yy As String
        Dim tmp_czas_startu As String
        Dim tmp_czas_emisji As String
        Dim tmp_tytul As String
        Dim tmp_plk_audio As String

        Dim wczytana_linia As String
        Dim koniec_pliku As Boolean = False
        '  Dim tmp_str As String
        Dim poz As Integer
        Dim rekord_prawidlowy As Boolean = False
        Dim r As clsPozycjaRaportuControlPIK
        Dim k As Integer = 0
        Dim msg As String
        Dim wynik As Integer = 0
        Dim k2 As Boolean
        Dim znak As String
        Dim p As clsProgram
        Dim kat_raportow As String = ""

        Dim gp_d As Date
        Dim gk_d As Date
        Dim tmp_czas_startu_d As Date

        Dim koniec_rekordu As Boolean
        Dim tmp_dlugosc As Integer = 0

        Try
            gp_d = CDate(gp)

        Catch ex As Exception

        End Try


        Try
            gk_d = CDate(gk)

        Catch ex As Exception

        End Try


        For Each p In colProgramy
            If p.id = nr_programu Then
                kat_raportow = p.parametry_polaczenia
                Exit For
            End If
        Next


        dd = Format(d, "dd")
        mm = Format(d, "MM")
        yy = Right(Year(d).ToString, 2)


        plk = kat_raportow
        If Right(plk, 1) <> "\" Then
            plk = plk & "\"
        End If

        plk = plk & "reklamy\"
        plk = plk & yy & "-" & mm & "-" & dd & ".txt"

        If System.IO.File.Exists(plk) = False Then
            Return 0
        End If

        Dim enc As System.Text.Encoding
        enc = System.Text.Encoding.Default

        Dim fStreamReader As New System.IO.StreamReader(plk, enc, True)

        'Pocz. emisji:  08:07:15
        'A -WESELE - 1 - 2017
        'Koniec emisji: 08:07:50
        '===================================
        'Pocz. emisji:  08:44:02
        'A -CHEAP - 2017
        'Koniec emisji: 08:44:34
        '===================================
        'Pocz. emisji:  09:07:16
        'A -WESELE - 1 - 2017
        'Koniec emisji: 09:07:51
        '===================================

        Dim tmp_str As String
        Dim tmp_start As String
        Dim tmp_stop As String
        Dim tmp_start_d As Date
        Dim tmp_stop_d As Date
        Dim ts As TimeSpan

        Try
            Do

                koniec_rekordu = False
                tmp_start = "00:00:00"
                tmp_stop = "00:00:00"
               
                Do
                    wczytana_linia = fStreamReader.ReadLine()
                    If wczytana_linia = Nothing Then
                        koniec_pliku = True
                        koniec_rekordu = True
                    Else
                        tmp_str = Trim(UCase(wczytana_linia))
                        If Left(tmp_str, 12) = "POCZ. EMISJI" Then
                            tmp_start = Right(wczytana_linia, 8)
                        ElseIf Left(tmp_str, 13) = "KONIEC EMISJI" Then
                            tmp_stop = Right(wczytana_linia, 8)

                        ElseIf Left(tmp_str, 1) = "=" Then
                            koniec_rekordu = True
                        End If
                    End If
                Loop While koniec_rekordu = False


                Try
                    tmp_start_d = CDate(tmp_start)
                Catch ex As Exception

                End Try

                Try
                    tmp_stop_d = CDate(tmp_stop)
                Catch ex As Exception

                End Try

                If tmp_start_d > gp_d Then
                    If tmp_stop_d < gk_d Then
                        ts = tmp_stop_d - tmp_start_d
                        tmp_dlugosc = tmp_dlugosc + ts.TotalSeconds
                    End If
                End If



            Loop While koniec_pliku = False

            wynik = tmp_dlugosc
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu pliku raportu poemisyjnego:"
            msg = msg & ex.Message
            MessageBox.Show(msg)
            wynik = -1
        End Try


        Try
            fStreamReader.Close()

        Catch ex As Exception

        End Try

        Return wynik



    End Function

    Private Function zaladuj_plik_raportu_ControPIK(ByVal nr_programu As Integer, ByVal d As Date, ByVal gp As String, gk As String, ByVal tryb As Integer) As Integer

        'tryb 0 - zwykły iomport pozycji
        '     1 - dołączenie 
        'w trybie 1 nie jest łądowana muzyka i jongle

        Dim plk As String
        Dim dd As String
        Dim mm As String
        Dim yy As String
        Dim tmp_czas_startu As String
        Dim tmp_czas_emisji As String
        Dim tmp_tytul As String
        Dim tmp_plk_audio As String

        Dim wczytana_linia As String
        Dim koniec_pliku As Boolean = False
        '  Dim tmp_str As String
        Dim poz As Integer
        Dim rekord_prawidlowy As Boolean = False
        Dim r As clsPozycjaRaportuControlPIK
        Dim k As Integer = 0
        Dim msg As String
        Dim wynik As Integer = 0
        Dim k2 As Boolean
        Dim znak As String
        Dim p As clsProgram
        Dim kat_raportow As String = ""

        Dim gp_d As Date
        Dim gk_d As Date
        Dim tmp_czas_startu_d As Date

        Try
            gp_d = CDate(gp)

        Catch ex As Exception

        End Try


        Try
            gk_d = CDate(gk)

        Catch ex As Exception

        End Try


        For Each p In colProgramy
            If p.id = nr_programu Then
                kat_raportow = p.parametry_polaczenia
                Exit For
            End If
        Next

        If colListaRaportuControlPIK.Count > 0 Then
            Do
                colListaRaportuControlPIK.Remove(1)
            Loop While colListaRaportuControlPIK.Count > 0
        End If

        dd = Format(d, "dd")
        mm = Format(d, "MM")
        yy = Right(Year(d).ToString, 2)


        plk = kat_raportow
        If Right(plk, 1) <> "\" Then
            plk = plk & "\"
        End If

        plk = plk & "redakcje\"
        plk = plk & yy & "-" & mm & "-" & dd & ".txt"

        If System.IO.File.Exists(plk) = False Then
            Return 0
        End If



        Dim a As Integer


        If tryb = 0 Then
            a = zaladuj_dlugosc_muzyki_ControPIK(nr_programu, d, gp, gk)

            If a > 0 Then
                k += 1
                r = New clsPozycjaRaportuControlPIK
                r.data_emisji = d
                r.tytul = "MUZYKA"
                r.godzina_emisji = gp ' Format(gp, "HH:mm:ss")
                r.dlugosc = zamien_sekundy_na_str(a) ' tmp_czas_emisji
                r.plik_audio = ""
                colListaRaportuControlPIK.Add(r, k & "_")
            End If


            a = 0
            a = zaladuj_dlugosc_reklam_ControPIK(nr_programu, d, gp, gk)

            If a > 0 Then
                k += 1
                r = New clsPozycjaRaportuControlPIK
                r.data_emisji = d
                r.tytul = "REKLAMA"
                r.godzina_emisji = gp ' Format(gp, "HH:mm:ss")
                r.dlugosc = zamien_sekundy_na_str(a) ' tmp_czas_emisji
                r.plik_audio = ""
                colListaRaportuControlPIK.Add(r, k & "_")

            End If


        End If


        Dim enc As System.Text.Encoding
        enc = System.Text.Encoding.Default

        Dim fStreamReader As New System.IO.StreamReader(plk, enc, True)

        'to sa dane w pliku
        '*** PIKI (KRÓTKIE) *01:59:54*      00:00:05 - Hotkeyk:\JINGLE\46814EFB.s48
        'JINGLE-BANIA-2007/2 *04:59:45*      00:00:11 - LISTA_Muzyczna_A^C:\Backup\Sounds\00006400.s48
        'PIKI *04:59:56*      00:00:05 - Hotkey^k:\JINGLE\Hotkeje\4FA0D140.s48
        'JINGLE-BANIA-2007/5 *05:28:40*      00:00:11 - LISTA_Muzyczna_B^C:\Backup\Sounds\00006403.s48
        '10.01 05:30 Z malowanej skrzyni *05:28:51*      00:30:59" - MLIST^k:\Nocemisj\7290EE63.s48
        '10.01 PORANNY 1 Wiadomości *06:00:05*      00:04:16" - Aktual  ^k:\AKTUALNO\72A14C4A.s48
        'B038 Domino-FilmowaMiłość *00:33:45*      00:04:03 - LISTA_Muzyczna_A^k:\muzyka\500101E7.s48
        'B PODKŁAD ZIMA JingleBells *00:37:47*      00:15:22 - LISTA_Muzyczna_B^k:\muzyka\58D125DD.s48

        ' nie zawcze po długości jest znak "
        'ale zauwazyłem ze zawsze po czasie startu jest 6 spacji


        Try
            Do

                tmp_czas_startu = "00:00:00"
                tmp_czas_emisji = "00:00:00"
                tmp_tytul = ""
                tmp_plk_audio = ""
                rekord_prawidlowy = False

                wczytana_linia = fStreamReader.ReadLine()
                If wczytana_linia = Nothing Then
                    koniec_pliku = True
                Else
                    If Len(wczytana_linia) > 0 Then

                        '
                        'poz = InStr(wczytana_linia, "^")
                        'poprawka z dnia 2 sierpnia 2011
                        'nie wiem dlaczego byc może ktoś dodał znak ^ w nazwie dźwieku był ten znak na początku i program źle interpretował linię
                        'dlatego poniższa poprawka
                        poz = InStr(10, wczytana_linia, "^")

                        'teraz ścieżka do pliku audio
                        If poz > 0 Then
                            tmp_plk_audio = Right(wczytana_linia, Len(wczytana_linia) - poz)
                        End If
                        wczytana_linia = Left(wczytana_linia, poz - 1)
                        'szukanie myslnika 
                        k2 = False
                        Do
                            znak = Right(wczytana_linia, 1)
                            wczytana_linia = Left(wczytana_linia, Len(wczytana_linia) - 1)
                            If znak = "-" Then
                                k2 = True
                            End If
                        Loop While k2 = False

                        wczytana_linia = Trim(wczytana_linia)
                        If Right(wczytana_linia, 1) = Chr(34) Then
                            'obcięcie cudzysłowy
                            'w warunku bo nie zawsze wystpeuje
                            wczytana_linia = Left(wczytana_linia, Len(wczytana_linia) - 1)
                        End If

                        'w Rzeszowie używaja myslników w nazwach folderów - to burzy procedurę
                        'przykłądowa linia:
                        '       12.05 23:05 Noc z RR1 *23:05:35*      00:54:13 - 18-24   ^n:\ANTENA\18-24\4_CZWART\10350BA35.s48
                        '
                        'dlatego tu sprawdzenie czy na trzecim  miejscu od końca jest dwukropek
                        'jeżeli nie a wśród znaów jest myślnik to usunięcie kolejnego myślnika

                        If Mid(wczytana_linia, Len(wczytana_linia) - 2, 1) <> ":" Then
                            'szukanie myslnika 
                            k2 = False
                            Do
                                znak = Right(wczytana_linia, 1)
                                wczytana_linia = Left(wczytana_linia, Len(wczytana_linia) - 1)
                                If znak = "-" Then
                                    k2 = True
                                End If
                            Loop While k2 = False

                        End If

                        wczytana_linia = Trim(wczytana_linia)
                        tmp_czas_emisji = Right(wczytana_linia, 8)
                        wczytana_linia = Left(wczytana_linia, Len(wczytana_linia) - 9)
                        wczytana_linia = Trim(wczytana_linia)
                        wczytana_linia = Left(wczytana_linia, Len(wczytana_linia) - 1)
                        tmp_czas_startu = Right(wczytana_linia, 8)
                        wczytana_linia = Left(wczytana_linia, Len(wczytana_linia) - 9)
                        tmp_tytul = Trim(wczytana_linia)
                        rekord_prawidlowy = True

                    End If
                End If

                If rekord_prawidlowy Then
                    Try
                        tmp_czas_startu_d = CDate(tmp_czas_startu)
                    Catch ex As Exception

                    End Try
                    If tmp_czas_startu_d > gp_d Then
                        If tmp_czas_startu_d < gk_d Then
                            k += 1
                            r = New clsPozycjaRaportuControlPIK
                            r.data_emisji = d
                            r.tytul = tmp_tytul
                            r.godzina_emisji = tmp_czas_startu
                            r.dlugosc = tmp_czas_emisji
                            r.plik_audio = tmp_plk_audio
                            If InStr(UCase(r.plik_audio), "JINGL") = 0 Then 'wstawiane sa tylko nagania nie jingle
                                colListaRaportuControlPIK.Add(r, k & "_")
                            End If

                        End If
                    End If
                End If

                '   If tmp_czas_startu = "15:10:30" Then
                'wynik = 0
                '    End If
            Loop While koniec_pliku = False


        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu pliku raportu poemisyjnego:"
            msg = msg & ex.Message
            MessageBox.Show(msg)
            wynik = -1
        End Try



        Try
            fStreamReader.Close()
        Catch ex As Exception

        End Try

        Return wynik



    End Function


    Private Function zaladuj_opis_nagrania(ByRef j As clsPozycjaRaportuControlPIK) As Integer


        Try
            If System.IO.File.Exists(j.plik_audio) Then
                Dim plik_bwf As New clsPlikBWF(j.plik_audio, False)

                If Right(UCase(j.plik_audio), 3) = "S48" Then
                    j.autor_tekstu = plik_bwf.PAD_autor_tekstu
                    j.kompozytor = plik_bwf.PAD_kompozytor
                    j.wykonawca = plik_bwf.PAD_wykonawca
                    j.producent = plik_bwf.PAD_producent
                    j.pad_autor = plik_bwf.PAD_autor
                    j.pad_tytul = plik_bwf.PAD_tytul
                    j.gatunek = UCase(Trim(plik_bwf.PAD_gatunek))


                    j.jest_plik_audio = True
                    Try
                        plik_bwf = Nothing
                    Catch ex As Exception

                    End Try
                    j.dane_opisowe_ok = True

                    '
                    'ten fragment z CoMY - rozpoznawanie Jingli w radiu Lublin - tu na razie niepotrzebne
                    'If InStr(UCase(j.gatunek), "JING") > 0 Then
                    ' j.dane_opisowe_ok = True
                    '                Else
                    '                   If system_emisyjny = 1 Then
                    'j.dane_opisowe_ok = True
                    'End If
                Else
                    'w plikach WAV dane pobierane są z chunka CART

                    j.pad_tytul = plik_bwf.CART_title
                    j.wykonawca = plik_bwf.CART_artist
                    j.gatunek = plik_bwf.CART_classification

                    j.kompozytor = podaj_dane_z_CART_TAGTEXT(plik_bwf.CART_tag_text, "KOMPOZYTOR")
                    j.autor_tekstu = podaj_dane_z_CART_TAGTEXT(plik_bwf.CART_tag_text, "AUTOR TEKSTU")
                    j.producent = podaj_dane_z_CART_TAGTEXT(plik_bwf.CART_tag_text, "PRODUCENT")

                    j.pad_autor = plik_bwf.PAD_autor

                    j.jest_plik_audio = True
                    Try
                        plik_bwf = Nothing
                    Catch ex As Exception

                    End Try

                    j.dane_opisowe_ok = True

                    ' If InStr(UCase(j.gatunek), "JING") > 0 Then
                    ' j.dane_opisowe_ok = True
                    'Else
                    'If system_emisyjny = 1 Then
                    'j.dane_opisowe_ok = True
                    'End If
                End If


            End If


        Catch ex As Exception

        End Try



    End Function


    Public Function podaj_dane_z_CART_TAGTEXT(ByVal tresc As String, ByVal znacznik As String) As String
        Dim wynik As String = ""
        Dim tmp_str As String
        Dim poz1 As Integer = 0
        Dim poz2 As Integer = 0
        Dim szukany_znacznik As String

        szukany_znacznik = "<" & znacznik & ">"


        tmp_str = Trim(tresc)

        If Len(tmp_str) > 0 Then
            Try
                If InStr(tmp_str, szukany_znacznik) > 0 Then
                    poz1 = InStr(tmp_str, szukany_znacznik)
                    tmp_str = Right(tmp_str, Len(tmp_str) - poz1 - Len(szukany_znacznik) + 1)
                    szukany_znacznik = "</" & znacznik & ">"
                    poz2 = InStr(tmp_str, szukany_znacznik)
                    If poz2 = 0 Then
                        poz2 = InStr(tmp_str, "</")
                    End If
                    If poz2 > 0 Then
                        Try
                            tmp_str = Left(tmp_str, poz2 - 1)
                        Catch ex11 As Exception

                        End Try
                    End If
                    wynik = Trim(tmp_str)
                End If
            Catch ex As Exception

            End Try


        End If

        Return wynik
    End Function


End Module
