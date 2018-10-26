Module basEksport

    Public colPozycjeEksportoweCDN As New Collection

    Public Function eksport_wycen_CDN() As Integer
        Dim a As Integer
        Dim upr As Boolean = False
        Dim dp As Date
        Dim dk As Date
        Dim dp1 As Date
        Dim dp_str As String
        Dim dk_str As String

        Dim msg As String
        Dim tmp_str As String = ""

        If aktualny_uzytkownik.uprawnienia_administratora Then
            upr = True
        End If

        If aktualny_uzytkownik.uprawnienia_kadry_place Then
            upr = True
        End If


        If upr = False Then
            MessageBox.Show("Brak uprawnień", naglowek_komunikatow)
            Return 1
        End If


        Dim o_dialog As New frmZakresDniDOZestawienia
        o_dialog.ShowDialog()

        If o_dialog.DialogResult = Windows.Forms.DialogResult.Cancel Then
            Return 1
        End If
        'tu dziwne podstawienia dp1 i dp
        ' chodzi o to że dialog zwraca datę i aktualną godzinę

        ' w zapytaniu sql w klauzuli BETWEEN data z aktualną godziną wyklucza wiersze z 
        'dnia poczatkowego
        'dlatego tu wyzerowanie godziny z dnia początkowego

        dp1 = o_dialog.dtpOdDNia.Value
        dp_str = Format(dp1, "yyyy-MM-dd")
        dp = CDate(dp_str)
        '            dp = o_dialog.dtpOdDNia.Value
        dk = o_dialog.dtpDoDNia.Value
        dk_str = Format(dk, "yyyy-MM-dd")


        msg = "Uruchomiono funkcję eksportu pozycji wniosków za okres od " & Format(dp, "yyyy-MM-dd")
        msg = msg & " do " & Format(dk, "yyyy-MM-dd")
        msg = msg & vbCrLf & "Wszystkie wnioski w ww okresie zostana zablokowane do edycji."

        msg = msg & vbCrLf & "Czy kontynuować ?"
        a = MessageBox.Show(msg, naglowek_komunikatow, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If a = vbNo Then Return 0



        Application.DoEvents()
        Dim o_progresu As New frmProgress
        Dim msg_stat As String

        o_progresu.ProgressBar1.Visible = False

        o_progresu.lblNaglowek.Text = "Eksport danych do systemu CDN"
        msg_stat = "Trwa kontrola zatwierdzenia wniosków ....."
        o_progresu.lblStatus.Text = msg_stat
        o_progresu.Show()
        Application.DoEvents()


        a = sprawdz_zatwierdzenie_wnioskow(dp, dk)
        If a > 0 Then
            msg = "W wybranym okresie są wnioski nie zatwierdzone przez Zarząd."
            msg = msg & vbCrLf & "Eksport nie jest możliwy."
            wskaznik_myszy(0)
            o_progresu.Close()
            'o_progresu.Dispose()
            MessageBox.Show(msg, naglowek_komunikatow, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return -1
        End If

        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa kontrola otwarcia wniosków ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        System.Threading.Thread.Sleep(1000)


        a = sprawdz_czy_wnioski_sa_edytowane(dp, dk, tmp_str)
        If a < 0 Then
            o_progresu.Close()
            ' o_progresu.Dispose()
            Return -1
        ElseIf a > 0 Then
            msg = "W wybranym okresie wniosek " & tmp_str & " jest otwarty z prawem do zapisu (jest edytowany)."
            msg = msg & vbCrLf & "Eksport będzie możliwy po zamknięciu wniosku"
            o_progresu.Close()
            'o_progresu.Dispose()
            MessageBox.Show(msg, naglowek_komunikatow)
            Return 1
        End If


        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa ładowanie pozycji wniosków ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        System.Threading.Thread.Sleep(1000)

        a = zaladuj_liste_wycen_do_eksportu(dp, dk)
        If a <> 0 Then
            o_progresu.Close()
            '  o_progresu.Dispose()
            Return a
        End If

        If colPozycjeEksportoweCDN.Count = 0 Then
            o_progresu.Close()
            '   o_progresu.Dispose()
            MessageBox.Show("Brak danych do eksportu")
            Return 0
        End If

        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa kontrola peseli pracowników....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()
        System.Threading.Thread.Sleep(1000)

        a = ustal_pesele_pracownikow()
        If a <> 0 Then
            o_progresu.Close()
            '  o_progresu.Dispose()
            Return a
        End If

        If zadaniowanie_dostepne Then
            a = ustal_konta_zadan_pozycji()
        End If

        If rozroznianie_zrodel_finansowania Then
            a = ustal_konta_zrodel()
        End If


        If oznaczanie_pracownikow_ryczaltowych_dostepne Then
            a = kontrola_wynagrodzen_ryczaltowych()
        End If

        If licencjobiorca = "Radio Lublin SA" Then
            msg_stat = msg_stat & "OK" & vbCrLf & "Trwa korekta identyfikatora usługi pracowników czasowo-premiowych ....."
            o_progresu.lblStatus.Text = msg_stat
            Application.DoEvents()
            System.Threading.Thread.Sleep(1000)

            a = skoryguj_identyfikatory_uslugi()
            If a <> 0 Then
                o_progresu.Close()
                '  o_progresu.Dispose()
                Return a
            End If
        End If

        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa kontrola już zapisanych pozycji ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()
        System.Threading.Thread.Sleep(1000)

        a = sprawdz_czy_juz_zapisano_pozycje(o_progresu)
        If a < 0 Then
            o_progresu.Close()
            '    o_progresu.Dispose()
            Return -1
        End If


        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa ustawianie blokady edycji wniosków ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        a = ustaw_blokade_wnioskow_w_wybr_okresie(dp, dk)

        If a < 0 Then
            o_progresu.Close()
            '     o_progresu.Dispose()
            Return -1
        End If


        o_progresu.Close()
        '   o_progresu.Dispose()

        Application.DoEvents()

        Dim o As New frmEksportCDN
        o.lblNaglowek.Text = "Lista wycen gotowych do eksportu za okres od " & dp_str & " do " & dk_str
        o.grdListaWYcen.DataSource = colPozycjeEksportoweCDN
        o.grdListaWYcen.RetrieveStructure()
        o.ustaw_siatke_spisu_pozycji()

        o.ShowDialog()


        Return 0


    End Function


    Private Function sprawdz_czy_juz_zapisano_pozycje(ByRef o As frmProgress) As Integer
        Dim wynik As Integer = 0
        Dim l_poz As Integer
        Dim k As Integer = 0
        Dim poz As clsPozycjaEKsportowa
        Dim a As Integer
        l_poz = colPozycjeEksportoweCDN.Count

        o.ProgressBar1.Visible = True

        o.ProgressBar1.Maximum = l_poz + 1

        For Each poz In colPozycjeEksportoweCDN
            k += 1
            o.ProgressBar1.Value = k
            a = kontrola_pozycji_eksportowej_CDN(poz)
            If a > 0 Then
                poz.zapisany = True
            ElseIf a < 0 Then
                wynik = -1
                Exit For
            End If
        Next

        Return wynik

    End Function


    Private Function kontrola_pozycji_eksportowej_CDN(ByRef poz As clsPozycjaEKsportowa) As Integer
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        ' Dim rok As String
        ' Dim a As Integer
        Dim msg As String
        Dim wynik As Integer = 0

        sql = sql & " select id "
        sql = sql & " from CDN_transfer "
        sql = sql & " WHERE rekord_zrodlowy=" & poz.rekord_zrodlowy


        Try
            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                wynik = dr.GetValue(0)
            End If
        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli zapisu pozycji:" & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            Cmd.Dispose()
            dr.Close()

        Catch ex As Exception

        End Try

        Return wynik

    End Function



    Private Function ustal_konta_zrodel() As Integer
        Dim poz As clsPozycjaEKsportowa



        For Each poz In colPozycjeEksportoweCDN
            poz.konto_zrodla = ustal_konto_zrodla(poz.zrodlo)
        Next

        Return 0


    End Function


    Public Function ustal_konto_zrodla(ByVal id_zrodla As Integer) As String
        Dim z As clsZrodloFinansowania
        Dim wynik As String = ""

        For Each z In colZrodlaFinansowania
            If z.id = id_zrodla Then
                wynik = z.konto
                Exit For
            End If
        Next

        Return wynik

    End Function


    Private Function ustal_konta_zadan_pozycji() As Integer
        Dim poz As clsPozycjaEKsportowa



        For Each poz In colPozycjeEksportoweCDN
            poz.konto_zadania = ustal_konto_zadania(poz.zadanie)
        Next

        Return 0


    End Function

    Public Function ustal_konto_zadania(ByVal id_zadania As Integer) As String
        Dim z As clsZadanie
        Dim wynik As String = ""

        For Each z In colZadania
            If z.id = id_zadania Then
                wynik = z.konto
                Exit For
            End If
        Next

        Return wynik

    End Function


    Public Function ustal_pesele_pracownikow() As Integer
        Dim a As Integer
        Dim poz As clsPozycjaEKsportowa
        Dim tmp_pesel As String
        Dim msg As String

        a = zaladuj_liste_pracownikow()
        If a <> 0 Then Return a

        For Each poz In colPozycjeEksportoweCDN
            tmp_pesel = ustal_pesel_pracownika(poz.id_autora)
            If Len(tmp_pesel) <> 11 Then
                msg = poz.imie_nazwisko & " posiada nieprawidłowo określony pesel."
                msg = msg & vbCrLf & "Eksport nie jest możliwy."
                MessageBox.Show(msg, naglowek_komunikatow)
                Return -1
            Else
                poz.pesel = tmp_pesel
            End If
        Next

    End Function


    Private Function ustal_czy_pracownik_ryczaltowy(ByVal id As Integer) As Boolean
        Dim wynik As Boolean = False

        Dim p As clsPracownik

        wynik = False

        For Each p In colPracownicy
            If p.id = id Then
                wynik = p.ryczalt
                Exit For
            End If
        Next

        Return wynik


    End Function

    Private Function skoryguj_identyfikatory_uslugi() As Integer
        'funkcja używana tylko w Radiu Lublin
        'pracownikcy czasowo premiowi będą miału zmienione końcowe cyfry identyfikatora usługi na 4

        Dim a As Integer
        Dim poz As clsPozycjaEKsportowa
        Dim tmp_pesel As String
        Dim msg As String
        Dim data_em As String
        Dim akt_status As Integer = 0
        Dim wsp As Decimal
        Dim wynik As Integer = 0
        Dim tmp_id_usl As String


        For Each poz In colPozycjeEksportoweCDN
            data_em = Format(poz.data_emisji, "yyyy-MM-dd")
            If Right(poz.identyfikator, 1) = "1" Then
                'powyższy warunek dlatego bo dotoczy to tylko wycen z 1 na końcu

                akt_status = ustal_aktualny_status_pracownika(poz.id_autora, data_em, wsp)
                'zwraca -1 - jeżeli nie ma umowy w danym dniu (umowy obowiązują na czas określony
                '       -2 - jeżeli błąd podczas sprawdzania

                If akt_status < 0 Then
                    wynik = -1
                    If akt_status = -1 Then
                        msg = poz.imie_nazwisko & " nie ma aktualnej umowy w dniu " & data_em
                        MessageBox.Show(msg)
                    End If
                    Exit For
                Else
                    If akt_status = 3 Then
                        tmp_id_usl = Left(poz.identyfikator, 13) & "4"
                        poz.identyfikator = tmp_id_usl
                    End If
                End If
            End If
        Next

        Return wynik

    End Function

    Private Function kontrola_wynagrodzen_ryczaltowych() As Integer
        Dim a As Integer
        Dim poz As clsPozycjaEKsportowa
        Dim tmp_pesel As String
        Dim msg As String
        Dim ryczaltowiec As Boolean = False

        For Each poz In colPozycjeEksportoweCDN
            ryczaltowiec = False
            ryczaltowiec = ustal_czy_pracownik_ryczaltowy(poz.id_autora)
            If ryczaltowiec Then
                poz.wycena = 0
            End If
        Next

    End Function




    Private Function ustal_pesel_pracownika(ByVal id As Integer) As String
        Dim wynik As String = ""
        Dim p As clsPracownik

        For Each p In colPracownicy
            If p.id = id Then
                wynik = p.pesel
                Exit For
            End If
        Next

        Return wynik
    End Function


    Public Function zaladuj_liste_wycen_do_eksportu(ByVal dp As Date, _
                                                    ByVal dk As Date) As Integer

        Dim dp_str As String
        Dim dk_str As String
        Dim sql As String = ""
        Dim msg As String
        Dim k As Integer = 0
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim poz As clsPozycjaEKsportowa


        If colPozycjeEksportoweCDN.Count > 0 Then
            Do
                colPozycjeEksportoweCDN.Remove(1)
            Loop While colPozycjeEksportoweCDN.Count > 0
        End If

        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")



        sql = "SET DATEFORMAT YMD "
        sql = sql & " SELECT "
        sql = sql & " pozycje_wniosku.id, "
        sql = sql & " wnioski.tytul_audycji, "
        sql = sql & " wnioski.data_emisji, "
        sql = sql & " wnioski.godzina_rozpoczecia, "
        sql = sql & " pozycje_wniosku.id_pracownika, "
        sql = sql & " pozycje_wniosku.imie_nazwisko_pracownika, "
        sql = sql & " pozycje_wniosku.nazwa_pozycji , "
        sql = sql & " pozycje_wniosku.zadanie, "
        sql = sql & " wnioski.zrodlo_finansowania, "
        sql = sql & " pozycje_wniosku.identyfikator, "
        sql = sql & " (pozycje_wniosku.stawka_podstawowa * pozycje_wniosku.wspolczynnik_wyceny * pozycje_wniosku.ilosc) as wycena, "
        If tryb_obslugi_MPK = 0 Then
            'ta zmiana dopisana w dniu 3 lutego 2012
            'Radio Kraków uzywa tego eksportu, uzywa teraz trybu 0 ale może ewentualnie przejśc na tryb 1 
            sql = sql & " wnioski.mpk "
        Else
            'ta zmiana dopisana w dniu 3 lutego 2012
            'Radio Lublin uzywa tego eksportu, uzywa teraz trybu 0 ale może ewentualnie przejśc na tryb 1 
            sql = sql & " pozycje_wniosku.mpk "
            'do tej pory program pobierał MPK z nagłowka wniosku
        End If
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' "
        sql = sql & " AND pozycje_wniosku.identyfikator NOT LIKE '%3'"
        sql = sql & " ORDER BY wnioski.data_emisji"



        Try
            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            Do
                If dr.Read Then
                    odczytano = True
                    poz = New clsPozycjaEKsportowa
                    poz.rekord_zrodlowy = dr.GetValue(0)
                    poz.tytul_audycji = dr.GetValue(1)
                    poz.data_emisji = dr.GetValue(2)
                    poz.godzina_emisji = dr.GetValue(3)
                    poz.id_autora = dr.GetValue(4)
                    poz.imie_nazwisko = dr.GetValue(5)
                    poz.nazwa_wyceny = dr.GetValue(6)
                    poz.zadanie = dr.GetValue(7)
                    poz.zrodlo = dr.GetValue(8)
                    poz.identyfikator = dr.GetValue(9)
                    poz.wycena = dr.GetValue(10)
                    poz.mpk = dr.GetValue(11)
                    colPozycjeEksportoweCDN.Add(poz, poz.rekord_zrodlowy & "_")
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu pozycji do eksportu: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            Cmd.Dispose()
            dr.Close()
        Catch ex As Exception

        End Try


        Return wynik



    End Function


    Public Function zapisz_liste_eksportowa_CDN() As Integer
        Dim poz As clsPozycjaEKsportowa
        Dim wynik As Integer = 0
        Dim a As Integer
        Dim l_poz As Integer
        Dim k As Integer = 0
        Dim o As New frmProgress


        l_poz = colPozycjeEksportoweCDN.Count

        o.ProgressBar1.Visible = True

        o.ProgressBar1.Maximum = l_poz + 1
        o.lblNaglowek.Text = "Zapis listy wycen."
        o.lblStatus.Text = "Trwa zapis ....."
        o.Show()


        If colPozycjeEksportoweCDN.Count = 0 Then Return 0

        For Each poz In colPozycjeEksportoweCDN
            k += 1
            o.ProgressBar1.Value = k
            If poz.zapisany = False Then
                a = zapisz_pozycje_eksportowa_CDN(poz)
                If a <> 0 Then
                    wynik = a
                    Exit For
                Else
                    poz.zapisany = True
                End If
            End If
        Next

        o.Close()
        ' o.Dispose()

        Return wynik

    End Function

    Private Function zapisz_pozycje_eksportowa_CDN(ByRef poz As clsPozycjaEKsportowa) As Integer
        Dim sql As String
        Dim a As Integer
        Dim nazwisko As String
        Dim nazwa As String
        Dim tytul As String
        Dim d_em As String
        Dim g_em As String
        Dim wycena As String

        If InStr(poz.imie_nazwisko, "'") > 0 Then
            nazwisko = skoryguj_apostrofy_do_SQL(poz.imie_nazwisko)
        Else
            nazwisko = poz.imie_nazwisko
        End If

        If InStr(poz.nazwa_wyceny, "'") > 0 Then
            nazwa = skoryguj_apostrofy_do_SQL(poz.nazwa_wyceny)
        Else
            nazwa = poz.nazwa_wyceny
        End If

        If InStr(poz.tytul_audycji, "'") > 0 Then
            tytul = skoryguj_apostrofy_do_SQL(poz.tytul_audycji)
        Else
            tytul = poz.tytul_audycji
        End If

        wycena = formatuj_wycene(poz.wycena)

        d_em = Format(poz.data_emisji, "yyyy-MM-dd")
        g_em = poz.godzina_emisji ' Format(poz.godzina_emisji, "HH:mm")

        sql = "SET DATEFORMAT YMD "
        sql = sql & " INSERT INTO CDN_transfer "
        sql = sql & "(rekord_zrodlowy, "
        sql = sql & "data_emisji, "
        sql = sql & "godzina_emisji, "
        sql = sql & "tytul_audycji, "
        sql = sql & "imie_nazwisko, "
        sql = sql & "pesel, "
        sql = sql & "identyfikator, "
        sql = sql & "nazwa_wyceny, "
        sql = sql & "wycena, "
        sql = sql & "konto_redakcji, "
        sql = sql & "konto_zadania, "
        sql = sql & "konto_zrodla_finansowania, "
        sql = sql & "mpk ) "

        sql = sql & " VALUES("
        sql = sql & poz.rekord_zrodlowy & ","
        sql = sql & "'" & d_em & "',"
        sql = sql & "'" & g_em & "',"
        sql = sql & "'" & tytul & "',"
        sql = sql & "'" & nazwisko & "',"
        sql = sql & "'" & poz.pesel & "',"
        sql = sql & "'" & poz.identyfikator & "',"
        sql = sql & "'" & nazwa & "',"
        sql = sql & "'" & wycena & "',"
        sql = sql & "'" & poz.konto_redakcji & "',"
        sql = sql & "'" & poz.konto_zadania & "',"
        sql = sql & "'" & poz.konto_zrodla & "',"
        sql = sql & "'" & poz.mpk & "')"

        a = wykonaj_polecenie_SQL(sql)
        Return a


    End Function


End Module
