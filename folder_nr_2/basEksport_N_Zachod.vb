Module basEksport_N_Zachod


    Public colEx_N_Zach_ListaPracownikow As New Collection
    Public colEx_N_Zach_ListaWspolPracownikow As New Collection



    Public Function N_eksport_wycen_zachod() As Integer
        'od grudnia 2009 od wersji 2.0.84 
        'Radio Zach�d ma dwa modu�y eksportowe - zmienili system kadrowo-p�acowy

        Dim liczba_pozycji As Integer = 0

        Dim a As Integer
        Dim upr As Boolean = False
        Dim dp As Date
        Dim dk As Date
        Dim dp1 As Date
        Dim dp_str As String
        Dim dk_str As String

        Dim msg As String
        Dim tmp_str As String = ""
        Dim poz As clsPozycjaEksportowaNZachod
        Dim kw As Decimal
        Dim ko As Decimal

        If aktualny_uzytkownik.uprawnienia_administratora Then
            upr = True
        End If

        If aktualny_uzytkownik.uprawnienia_kadry_place Then
            upr = True
        End If


        If upr = False Then
            MessageBox.Show("Brak uprawnie�", naglowek_komunikatow)
            Return 1
        End If


        Dim o_dialog As New frmZakresDniDOZestawienia
        o_dialog.ShowDialog()

        If o_dialog.DialogResult = Windows.Forms.DialogResult.Cancel Then
            Return 1
        End If
        'tu dziwne podstawienia dp1 i dp
        ' chodzi o to �e dialog zwraca dat� i aktualn� godzin�

        ' w zapytaniu sql w klauzuli BETWEEN data z aktualn� godzin� wyklucza wiersze z 
        'dnia poczatkowego
        'dlatego tu wyzerowanie godziny z dnia pocz�tkowego

        dp1 = o_dialog.dtpOdDNia.Value
        dp_str = Format(dp1, "yyyy-MM-dd")
        dp = CDate(dp_str)
        '            dp = o_dialog.dtpOdDNia.Value
        dk = o_dialog.dtpDoDNia.Value
        dk_str = Format(dk, "yyyy-MM-dd")


        msg = "Uruchomiono funkcj� eksportu pozycji wniosk�w za okres od " & Format(dp, "yyyy-MM-dd")
        msg = msg & " do " & Format(dk, "yyyy-MM-dd")
        msg = msg & vbCrLf & "Wszystkie wnioski w ww okresie zostana zablokowane do edycji."

        msg = msg & vbCrLf & "Czy kontynuowa� ?"
        a = MessageBox.Show(msg, naglowek_komunikatow, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If a = vbNo Then Return 0



        Application.DoEvents()
        Dim o_progresu As New frmProgress
        Dim msg_stat As String

        o_progresu.ProgressBar1.Visible = False

        o_progresu.lblNaglowek.Text = "Eksport danych."
        msg_stat = "Trwa kontrola zatwierdzenia wniosk�w ....."
        o_progresu.lblStatus.Text = msg_stat
        o_progresu.Show()
        Application.DoEvents()


        a = sprawdz_zatwierdzenie_wnioskow(dp, dk)
        If a > 0 Then
            msg = "W wybranym okresie s� wnioski nie zatwierdzone przez Zarz�d."
            msg = msg & vbCrLf & "Eksport nie jest mo�liwy."
            wskaznik_myszy(0)
            o_progresu.Close()
            'o_progresu.Dispose()
            MessageBox.Show(msg, naglowek_komunikatow, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return -1
        End If

        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa kontrola otwarcia wniosk�w ....."
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
            msg = msg & vbCrLf & "Eksport b�dzie mo�liwy po zamkni�ciu wniosku"
            o_progresu.Close()
            'o_progresu.Dispose()
            MessageBox.Show(msg, naglowek_komunikatow)
            Return 1
        End If





        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa �adowanie listy pracownik�w z wystawionych wniosk�w ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        System.Threading.Thread.Sleep(1000)


        a = EX_N_Zach_zaladuj_liste_pracownikow(dp, dk)

        If a < 0 Then
            o_progresu.Close()
            ' o_progresu.Dispose()
            Return -1
        End If


        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa �adowanie listy wsp�pracownik�w z wystawionych wniosk�w ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        System.Threading.Thread.Sleep(1000)

        a = EX_N_Zach_zaladuj_liste_wspolpracownikow(dp, dk)
        If a < 0 Then
            o_progresu.Close()
            ' o_progresu.Dispose()
            Return -1
        End If

        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa �adowanie listy os�b  ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        System.Threading.Thread.Sleep(1000)

        a = zaladuj_liste_pracownikow()
        If a < 0 Then
            o_progresu.Close()
            ' o_progresu.Dispose()
            Return -1
        End If


        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa kontrola peseli pracownik�w  ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        System.Threading.Thread.Sleep(1000)

        a = Ex_N_Zach_kontrola_peseli_pracownikow()
        If a < 0 Then
            o_progresu.Close()
            ' o_progresu.Dispose()
            Return -1
        End If


        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa kontrola peseli wsp�pracownik�w  ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        System.Threading.Thread.Sleep(1000)

        a = Ex_N_Zach_kontrola_peseli_wspolpracownikow()
        If a < 0 Then
            o_progresu.Close()
            ' o_progresu.Dispose()
            Return -1
        End If

        liczba_pozycji = colEx_N_Zach_ListaPracownikow.Count + colEx_N_Zach_ListaWspolPracownikow.Count

        If liczba_pozycji = 0 Then
            o_progresu.Close()
            MessageBox.Show("Brak danych do eksportu")
            Return -1
        End If

        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa kontrola �adowanie wynagrodze� pracownik�w  ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        If colEx_N_Zach_ListaPracownikow.Count > 0 Then
            For Each poz In colEx_N_Zach_ListaPracownikow
                a = EX_N_ZACH_zaladuj_kwoty(poz, dp, dk, 1, kw)
                If a <> 0 Then
                    o_progresu.Close()
                    Return -1
                Else
                    poz.kwota_wlasne = kw
                End If
            Next
        End If


        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa kontrola �adowanie wynagrodze� wp�pracownik�w  ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        If colEx_N_Zach_ListaWspolPracownikow.Count > 0 Then
            For Each poz In colEx_N_Zach_ListaWspolPracownikow
                a = EX_N_ZACH_zaladuj_kwoty(poz, dp, dk, 2, kw)
                If a <> 0 Then
                    o_progresu.Close()
                    Return -1
                Else
                    poz.kwota_wlasne = kw
                End If
            Next
        End If

        Dim klucz As String

        If colEx_N_Zach_ListaPracownikow.Count > 0 Then
            'usuni�cie wierszy z zerowymi kwotami
            For Each poz In colEx_N_Zach_ListaPracownikow
                If poz.kwota_wlasne = 0 Then
                    klucz = poz.id & "_" & poz.program
                    colEx_N_Zach_ListaPracownikow.Remove(klucz)
                End If
            Next
        End If


        If colEx_N_Zach_ListaWspolPracownikow.Count > 0 Then
            For Each poz In colEx_N_Zach_ListaWspolPracownikow
                If poz.kwota_wlasne = 0 Then
                    klucz = poz.id & "_" & poz.program
                    colEx_N_Zach_ListaWspolPracownikow.Remove(klucz)
                End If
            Next
        End If



        msg_stat = msg_stat & "OK" & vbCrLf & "Trwa ustawianie blokady edycji wniosk�w ....."
        o_progresu.lblStatus.Text = msg_stat
        Application.DoEvents()

        a = ustaw_blokade_wnioskow_w_wybr_okresie(dp, dk)


        o_progresu.Close()

        Dim o As New frmEksportZachod
        o.Text = "Eksport Radio Zach�d - Symfonia"
        o.data_poczatkowa = dp
        o.data_koncowa = dk
        o.lblNaglowek.Text = "Pe�na lista honorari�w za okres od " & Format(dp, "yyyy-MM-dd") & " do " & Format(dk, "yyyy-MM-dd") & " gotowa do eksportu do systemu Symfonia."
        o.grdPracownicy.DataSource = colEx_N_Zach_ListaPracownikow
        o.grdPracownicy.RetrieveStructure()
        o.Ex_N_Zach_ustaw_siatke_pracownikow()


        o.grdWspolpracownicy.DataSource = colEx_N_Zach_ListaWspolPracownikow
        o.grdWspolpracownicy.RetrieveStructure()
        o.Ex_N_Zach_ustaw_siatke_wspolpracownikow()
        o.ShowDialog()




    End Function



    Public Function EX_N_ZACH_zaladuj_kwoty(ByRef poz As clsPozycjaEksportowaNZachod, _
                                        ByVal dp As Date, _
                                        ByVal dk As Date, _
                                        ByVal status As Integer, _
                                        ByRef kwota_wlasne As Decimal) As Integer


        Dim sql As String
        Dim r1 As String
        Dim r2 As String
        Dim msg As String
        Dim dp_str As String
        Dim dk_str As String
        Dim aa As Decimal
        Dim wynik As Integer = 0

        kwota_wlasne = 0
        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
        sql = sql & " * pozycje_wniosku.ilosc "
        sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
        sql = sql & " AS Suma_wydatkow"
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE pozycje_wniosku.identyfikator Like '%" & status & "' "
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & dp & "' AND '" & dk & "'"
        sql = sql & " AND pozycje_wniosku.id_pracownika=" & poz.id
        sql = sql & " AND wnioski.zrodlo_finansowania=1"
        sql = sql & " AND wnioski.nr_programu=" & poz.program

        aa = wczytaj_sume_kosztow(sql)
        If aa < 0 Then
            kwota_wlasne = 0
            Return -1
        Else
            kwota_wlasne = aa
        End If


        Return wynik



    End Function



    Public Function EX_N_Zach_zaladuj_liste_pracownikow(ByVal dp As Date, ByVal dk As Date) As Integer
        'od grudnia 2009 od wersji 2.0.84 
        'Radio Zach�d ma dwa modu�y eksportowe - zmienili system kadrowo-p�acowy

        Dim sql As String

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim ii As Integer = 0
        Dim id_prac As Integer = 0
        Dim p As clsPozycjaEksportowaNZachod

        If colEx_N_Zach_ListaPracownikow.Count > 0 Then
            Do
                colEx_N_Zach_ListaPracownikow.Remove(1)
            Loop While colEx_N_Zach_ListaPracownikow.Count > 0
        End If


        d_p_str = Format(dp, "yyyy-MM-dd")
        d_k_str = Format(dk, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & " SELECT  DISTINCT  pozycje_wniosku.id_pracownika "
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE pozycje_wniosku.identyfikator like '%1' "
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak po��czenia z serwerem SQL."
                        msg = msg & "Prosz� skontaktowa� si� z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    'procedura tworzy instancj� w kolekcji dla kolejnych program�w oddzielnie 
                    id_prac = dr.GetValue(0)
                    p = New clsPozycjaEksportowaNZachod
                    p.id = id_prac
                    p.program = 1
                    colEx_N_Zach_ListaPracownikow.Add(p, p.id & "_1")
                    If colProgramy.Count > 1 Then
                        p = New clsPozycjaEksportowaNZachod
                        p.id = id_prac
                        p.program = 2
                        colEx_N_Zach_ListaPracownikow.Add(p, p.id & "_2")
                    End If
                    If colProgramy.Count > 2 Then
                        p = New clsPozycjaEksportowaNZachod
                        p.id = id_prac
                        p.program = 3
                        colEx_N_Zach_ListaPracownikow.Add(p, p.id & "_3")
                    End If

                    odczytano = True
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wyst�pi� problem podczas odczytu listy pracownik�w z wystawionymi wycenami: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "B��D !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik

    End Function


    Public Function EX_N_Zach_zaladuj_liste_wspolpracownikow(ByVal dp As Date, ByVal dk As Date) As Integer
        'od grudnia 2009 od wersji 2.0.84 
        'Radio Zach�d ma dwa modu�y eksportowe - zmienili system kadrowo-p�acowy

        Dim sql As String

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim ii As Integer = 0
        Dim id_prac As Integer = 0
        Dim p As clsPozycjaEksportowaNZachod

        If colEx_N_Zach_ListaWspolPracownikow.Count > 0 Then
            Do
                colEx_N_Zach_ListaWspolPracownikow.Remove(1)
            Loop While colEx_N_Zach_ListaWspolPracownikow.Count > 0
        End If


        d_p_str = Format(dp, "yyyy-MM-dd")
        d_k_str = Format(dk, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & " SELECT  DISTINCT  pozycje_wniosku.id_pracownika "
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE pozycje_wniosku.identyfikator like '%2' "
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak po��czenia z serwerem SQL."
                        msg = msg & "Prosz� skontaktowa� si� z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    'procedura tworzy instancj� w kolekcji dla kolejnych program�w oddzielnie 
                    id_prac = dr.GetValue(0)
                    p = New clsPozycjaEksportowaNZachod
                    p.id = id_prac
                    p.program = 1
                    colEx_N_Zach_ListaWspolPracownikow.Add(p, p.id & "_1")
                    If colProgramy.Count > 1 Then
                        p = New clsPozycjaEksportowaNZachod
                        p.id = id_prac
                        p.program = 2
                        colEx_N_Zach_ListaWspolPracownikow.Add(p, p.id & "_2")
                    End If
                    If colProgramy.Count > 2 Then
                        p = New clsPozycjaEksportowaNZachod
                        p.id = id_prac
                        p.program = 3
                        colEx_N_Zach_ListaWspolPracownikow.Add(p, p.id & "_3")
                    End If

                    odczytano = True
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wyst�pi� problem podczas odczytu listy wsp�pracownik�w z wystawionymi wycenami: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "B��D !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik

    End Function




    Public Function Ex_N_Zach_kontrola_peseli_pracownikow() As Integer
        Dim poz As clsPozycjaEksportowaNZachod
        Dim a As Integer

        If colEx_N_Zach_ListaPracownikow.Count = 0 Then Return 0


        For Each poz In colEx_N_Zach_ListaPracownikow

            a = EX_N_Zach_ustal_pesel_osoby(poz)
            If a <> 0 Then
                Return a
            End If
        Next

    End Function

    Public Function Ex_N_Zach_kontrola_peseli_wspolpracownikow() As Integer
        Dim poz As clsPozycjaEksportowaNZachod
        Dim a As Integer

        If colEx_N_Zach_ListaWspolPracownikow.Count = 0 Then Return 0


        For Each poz In colEx_N_Zach_ListaWspolPracownikow
            a = EX_N_Zach_ustal_pesel_osoby(poz)
            If a <> 0 Then
                Return a
            End If
        Next

    End Function


    Private Function EX_N_Zach_ustal_pesel_osoby(ByRef poz As clsPozycjaEksportowaNZachod) As Integer
        Dim p As clsPracownik
        Dim wynik As Integer = -1
        Dim msg As String


        For Each p In colPracownicy
            If p.id = poz.id Then
                poz.pesel = p.pesel
                poz.imie_nazwisko = p.imie_nazwisko
                wynik = 0
                Exit For
            End If
        Next

        If Len(poz.pesel) <> 11 Then
            msg = poz.imie_nazwisko & " (id: " & poz.id & ") ma nieprawid�owo okre�lony pesel."
            msg = msg & vbCrLf & "Eksport wycen nie jest mo�liwy."
            MessageBox.Show(msg)
            wynik = -1
        End If

        Return wynik
    End Function


    Public Function RZ_N_zapisz_pliki_eksportowe(ByVal dp As Date, ByVal dk As Date)
        Dim a As Integer
        Dim n_plk As String

        Dim nazwa_eksportowanego_pliku As String = ""
        Dim rok As String
        Dim miesiac As String
        Dim msg As String = ""
        Dim p As clsPozycjaEksportowaNZachod
        Dim wynik As Integer = 0

        Dim l_rek As Integer


        l_rek = colEx_N_Zach_ListaPracownikow.Count + colEx_N_Zach_ListaWspolPracownikow.Count

        If l_rek = 0 Then
            msg = "Brak danych do eksportu"
            MessageBox.Show(msg)
            Return 0
        End If


        rok = Year(dp)

        If Month(dp) < 10 Then
            miesiac = "0" & Month(dk)
        Else
            miesiac = Month(dk)
        End If

        nazwa_eksportowanego_pliku = katalog_eksportowy_ZACHOD
        If Right(nazwa_eksportowanego_pliku, 1) <> "\" Then
            nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & "\"
        End If
        nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & "Honoraria pracownik�w_"
        nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & rok
        '        nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & miesiac & ".sym"
        'zmiana z dnia 20 stycznia 2011
        'HARP 2.0.101 - zmieniono rozszerzenie pliku i kodowanie znak�w na WIN 1250
        nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & miesiac & ".txt"


        Try
            If System.IO.File.Exists(nazwa_eksportowanego_pliku) Then
                System.IO.File.Delete(nazwa_eksportowanego_pliku)
            End If
        Catch ex As Exception
            msg = "Wyst�pi� problem podczas usuwania wcze�niej zapisanego pliku eksportowego " & nazwa_eksportowanego_pliku
            msg = msg & vbCrLf & ex.Message
            MessageBox.Show(msg)
            Return -1
        End Try

        If colEx_N_Zach_ListaPracownikow.Count > 0 Then
            a = RZ_N_zapisz_naglowek_pliku(nazwa_eksportowanego_pliku, 1)
            If a <> 0 Then Return -1

            For Each p In colEx_N_Zach_ListaPracownikow
                a = RZ_N_zapisz_rekord_listy_plac(nazwa_eksportowanego_pliku, p)
                If a <> 0 Then
                    wynik = -1
                    Exit For
                End If
            Next
        End If

        If wynik <> 0 Then
            Return wynik
        End If



        nazwa_eksportowanego_pliku = katalog_eksportowy_ZACHOD
        If Right(nazwa_eksportowanego_pliku, 1) <> "\" Then
            nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & "\"
        End If
        nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & "Honoraria wsp�pracownik�w_"
        nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & rok
        '        nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & miesiac & ".sym"
        'zmiana z dnia 20 stycznia 2011
        'HARP 2.0.101 - zmieniono rozszerzenie pliku i kodowanie znak�w na WIN 1250
        nazwa_eksportowanego_pliku = nazwa_eksportowanego_pliku & miesiac & ".txt"


        Try
            If System.IO.File.Exists(nazwa_eksportowanego_pliku) Then
                System.IO.File.Delete(nazwa_eksportowanego_pliku)
            End If
        Catch ex As Exception
            msg = "Wyst�pi� problem podczas usuwania wcze�niej zapisanego pliku eksportowego " & nazwa_eksportowanego_pliku
            msg = msg & vbCrLf & ex.Message
            MessageBox.Show(msg)
            Return -1
        End Try

        If colEx_N_Zach_ListaWspolPracownikow.Count > 0 Then
            a = RZ_N_zapisz_naglowek_pliku(nazwa_eksportowanego_pliku, 2)
            If a <> 0 Then Return -1

            For Each p In colEx_N_Zach_ListaWspolPracownikow
                a = RZ_N_zapisz_rekord_listy_plac(nazwa_eksportowanego_pliku, p)
                If a <> 0 Then
                    wynik = -1
                    Exit For
                End If
            Next
        End If

        Return wynik

    End Function

    Public Function RZ_N_zapisz_naglowek_pliku(ByVal plik As String, _
                                        ByVal grupa_osob As Integer) As Integer

        'grupa os�b -   1 - pracownicy
        '               2 - wsp�pracownicy

        Dim str As String
        Dim wynik As Integer = 0
        Dim msg As String = ""
        Dim enc As System.Text.Encoding

        enc = System.Text.Encoding.Default


        If grupa_osob = 1 Then
            str = "Nazwisko Imi�;Pesel;honoraria_1;koszty_1" & vbCrLf
        Else
            str = "Nazwisko Imi�;Pesel;honoraria_2;koszty_1" & vbCrLf
        End If


        Try
            '            System.IO.File.WriteAllText(plik, str)
            'zmiana z dznia 21 stycznia 2011 wersja 2.0.101 
            '
            System.IO.File.WriteAllText(plik, str, enc)

        Catch ex As Exception
            msg = "Wyst�pi� problem podczas zapisu nag��wka pliku eksportowego " & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1
        End Try


        Return wynik


    End Function


    Public Function RZ_N_zapisz_rekord_listy_plac(ByVal plik As String, ByRef p As clsPozycjaEksportowaNZachod) As Integer

        Dim str As String
        Dim n As String
        Dim pesel As String
        Dim k_w As String
        Dim wynik As Integer = 0
        Dim msg As String = ""
        Dim enc As System.Text.Encoding

        enc = System.Text.Encoding.Default

        'przyk��dowy rekord
        'Kowalska Danuta;61120802184;200,00;1

        n = p.imie_nazwisko
        pesel = p.pesel
        k_w = Format(p.kwota_wlasne, "###0.00")



        'tu zamiana ewentualnej kropki na przecinek jako separator dziesi�tny w kwocie 
        k_w = Left(k_w, Len(k_w) - 3) & "," & Right(k_w, 2)


        str = n & ";" & pesel & ";" & k_w & ";" & p.program & vbCrLf


        Try
            '           System.IO.File.AppendAllText(plik, str)
            'zmiana z dnia 21 stycznia 2011 HARP 2.0.101
            System.IO.File.AppendAllText(plik, str, enc)

        Catch ex As Exception
            msg = "Wyst�pi� problem podczas zapisu eksportowanego rekordu listy honorari�w " & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try

        Return wynik


    End Function



End Module
