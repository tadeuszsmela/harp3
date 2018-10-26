Module basBazaDanych
    Public serwer_SQL As String
    Public baza_danych_SQL As String
    Public uzytkownik_SQL As String
    Public haslo_SQL As String
    Public szyfrowanie_sql As Boolean = False

    Public tryb_zapisu_ustawien As Integer = 0 ' 0 - plik serverSQL.inf
    '                                               1 - rejestr systemu 
    Public polaczenie_sql As System.Data.SqlClient.SqlConnection


    Public Function sprawdz_pole_tabeli_db(ByVal tabela As String, ByVal nazwa_pola As String) As Integer

        Dim sql As String
        Dim a As Integer

        sql = "select " & nazwa_pola & " from " & tabela

        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function


    Public Function otworz_polaczenie_sql() As Integer
        Dim kon_str As String
        Dim msg As String
        Dim wynik As Integer = 0

        kon_str = "server=" & serwer_SQL & "; "
        kon_str = kon_str & "uid=" & uzytkownik_SQL & "; "
        kon_str = kon_str & "pwd=" & haslo_SQL & ";"
        kon_str = kon_str & "database=" & baza_danych_SQL & ";"
        If szyfrowanie_sql Then
            kon_str = kon_str & "Encrypt=yes;" ' "Encrypt=yes;"
        End If

        If Not IsNothing(polaczenie_sql) Then
            Try
                polaczenie_sql.Dispose()
            Catch ex As Exception

            End Try
        End If

        Try
            wskaznik_myszy(1)
            polaczenie_sql = New System.Data.SqlClient.SqlConnection(kon_str)
            polaczenie_sql.Open()
        Catch ex As Exception
            msg = "Wystąpił problem podczas nawiązywania połączenia z serwerem " & ex.Message
            MessageBox.Show(msg)
            wynik = -1
        End Try

        wskaznik_myszy(0)

        Return wynik

    End Function


    Public Function zapisz_ustawienia_servera_SQL(ByVal serwer As String, _
                                    ByVal user As String, _
                                    ByVal haslo As String, _
                                    ByVal baza As String, _
                                    ByVal tryb_zapisu As Integer, _
                                    ByVal szyfrowanie As Boolean) As Integer

        Dim dane As String

        Dim nazwa_pliku As String
        Dim tmp_str As String
        Dim msg As String
        Dim kl As Microsoft.Win32.RegistryKey
        Dim a As Integer
        Dim s As Integer = 0

        If szyfrowanie Then
            s = 1
        End If
        tmp_str = serwer & ";"
        tmp_str = tmp_str & user & ";"
        tmp_str = tmp_str & haslo & ";"
        tmp_str = tmp_str & baza
        'to dodane w wersji 3.0.36
        '18 maja 2018 - w związku z wejściem RODO
        'to zaremowane w dniu 21 czerwca 2018- powrót 
        'ostatecznie dodane w dniu 29 sierpnia 2018
        tmp_str = tmp_str & ";" & s


        '        dane = szyrfuj_usytawienia_polaczenia(tmp_str)

        'zmiana z dnia 18 maja 2018 
        'w związku z wejściem RODO 

        'powrót w dniu 21 czerwca 2018- na razie w związku z tym że Opole nie pobrało tej wersji a potrzebuję modyfikacji dla Lublina
        'ostatecznie dodane w dniu 29 sierpnia 2018
        'wersja 3.0.36 
        dane = zaszyrfuj_usytawienia_polaczenia_aes(tmp_str)

        'wg założeń tak zaszyfrowane dane sa zapisywane w pliku o zmienionej nazwie serverSQLa.inf
        'podczas startu programu ładowanuy jest plik z a na końcu ' no chyba że go nie ma to ładowany jest plik "stary"
        If Len(dane) = 0 Then 'wystąpił problem podczas szyfrowania
            Return -1
        End If




        If tryb_zapisu = 0 Then
            nazwa_pliku = Windows.Forms.Application.StartupPath

            If Right(nazwa_pliku, 1) <> "\" Then
                nazwa_pliku = nazwa_pliku & "\"
            End If

            '            nazwa_pliku = nazwa_pliku & "config\serverSQL.inf"
            'zmiana z dnia 29 sierpnia 2018
            'wersja 3.0.36
            'zapisując dane stosowane jest szyfrowanie AES - i tworzony jest przy tym plik o nazwie serverSQLa.inf

            nazwa_pliku = nazwa_pliku & "config\serverSQLa.inf"

            Try
                If System.IO.File.Exists(nazwa_pliku) Then
                    System.IO.File.Delete(nazwa_pliku)
                End If

            Catch ex As Exception
                msg = "Wystąpił problem podczas zapisu pliku ustawień serwera SQL " & ex.Message
                MessageBox.Show(msg)
                Return -1
            End Try

            Try
                System.IO.File.WriteAllText(nazwa_pliku, dane)
            Catch ex As Exception
                msg = "Wystąpił problem podczas zapisu pliku ustawień serwera SQL " & ex.Message
                MessageBox.Show(msg)
                Return -1
            End Try

            Try
                kl = My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\ComRad\HARP", True)
                a = 0
                If kl Is Nothing Then
                    a = utworz_klucze_rejestru()
                End If
                If a = 0 Then
                    kl = My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\ComRad\HARP", True)
                    kl.SetValue("parametry", "")
                    kl.SetValue("zapis_ustawien", "0")
                Else
                    msg = "Wystąpił problem podczas tworzenia klucza rejestru. "
                    msg = msg & vbCrLf & "Zapis ustawień jest możliwy tylko dla użytkowników z upranieniami administratora systemu Windows"
                    'ten komunikat w dniu 30 sierpnia został zaremowany w dniu 30 sierpnia 2018 v 3.0.36
                    '      MessageBox.Show(msg)
                    Return -1
                End If
            Catch ex As Exception
                msg = "Wystąpił problem podczas zapisu ustawień do rejestru"
                msg = msg & vbCrLf & ex.Message
                'ten komunikat w dniu 30 sierpnia został zaremowany w dniu 30 sierpnia 2018 v 3.0.36
                '                MessageBox.Show(msg)
                Return -1
            End Try


        Else 'do rejestru
            Try
                kl = My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\ComRad\HARP", True)
                a = 0
                If kl Is Nothing Then
                    a = utworz_klucze_rejestru()
                End If
                If a = 0 Then
                    kl = My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\ComRad\HARP", True)
                    kl.SetValue("parametry", dane)
                    kl.SetValue("zapis_ustawien", "1")
                Else
                    msg = "Wystąpił problem podczas tworzenia klucza rejestru. "
                    msg = msg & vbCrLf & "Zapis ustawień jest możliwy tylko dla użytkowników z upranieniami administratora systemu Windows"
                    MessageBox.Show(msg)
                    Return -1
                End If


            Catch ex As Exception
                msg = "Wystąpił problem podczas zapisu ustawień do rejestru"
                msg = msg & vbCrLf & ex.Message
                MessageBox.Show(msg)

                Return -1
            End Try

        End If

        Return 0

    End Function


    Public Function ustal_id_nowododanego_rekordu() As Integer
        Dim sql As String
        Dim msg As String
        Dim cmd As System.Data.SqlClient.SqlCommand
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim wynik As Integer = 0

        wskaznik_myszy(1)


        sql = "select @@identity as indent " 'from wnioski"
        '        sql = "select SCOPE_IDENTITY()"

        Try
            cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = cmd.ExecuteReader
            If dr.Read() Then
                wynik = dr.GetValue(0)
            End If
        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli zapisu nowego rekodru:" & vbCrLf & ex.Message
            wskaznik_myszy(0)
            MessageBox.Show(msg)
            wynik = -1
        End Try

        Try
            dr.Close()
            cmd.Dispose()
        Catch ex As Exception

        End Try

        wskaznik_myszy(0)

        Return wynik

    End Function
    Public Function wykonaj_polecenie_SQL_identity(ByVal komenda As String) As Integer
        'ta funkcja po dodaniu rekordu zwraca id rekordu nowo dodanego
        ' na dzień 25 marca 2008 jest nie używana bo informacja o IDENTITY jest wyciągana w oddzielnej funkcji
        Dim sql As String
        Dim msg As String
        Dim cmd As System.Data.SqlClient.SqlCommand
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim wynik As Integer = 0
        Dim id_rekordu As Integer = 0

        wskaznik_myszy(1)


        sql = "SET NOCOUNT ON "
        sql = sql & komenda
        sql = sql & vbCrLf & " SELECT @@IDENTITY"


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            cmd = New System.Data.SqlClient.SqlCommand(komenda, polaczenie_sql)
            dr = cmd.ExecuteReader()
            dr.Read()
            wynik = dr.GetValue(0)

        Catch ex As Exception
            msg = "Wystąpił problem podczas wykonywania polecenia zapisu do bazy danych:" & vbCrLf & ex.Message
            wskaznik_myszy(0)
            MessageBox.Show(msg)
            wynik = -1
        End Try

        Try
            dr.Close()
            cmd.Dispose()
        Catch ex As Exception

        End Try

        wskaznik_myszy(0)

        Return wynik

    End Function
    Public Function wykonaj_polecenie_SQL(ByVal komenda As String) As Integer
        Dim msg As String
        Dim cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0

        wskaznik_myszy(1)


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            cmd = New System.Data.SqlClient.SqlCommand(komenda, polaczenie_sql)
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            msg = "Wystąpił problem podczas wykonywania polecenia zapisu do bazy danych:" & vbCrLf & ex.Message
            wskaznik_myszy(0)
            MessageBox.Show(msg)
            wynik = -1
        End Try

        Try
            cmd.Dispose()
        Catch ex As Exception

        End Try
        wskaznik_myszy(0)

        Return wynik

    End Function

    Public Function sprawdz_tabele_nieobecnosci(ByVal id_pracownika As Integer, ByVal data_emisji As String) As Integer
        'funkcja używana podczas wystawiania wycen we wnioskach
        'sprawdza czy osoba w terminie podanym jako parametr jest zapisana w tabeli nieobecności
        'funkcja zwraca 
        '               0 jeżeli nie jest za[pisany jako nieobecnu
        '               1 jeżeli pracownik w danym dniu jest na urlopie wypocz
        '               2 jeżeli pracownik w danym dniu jest na urlopie okol
        '               3 jeżeli pracownik w danym dniu jest na urlopie bezpłatnum
        '               4 jeżeli pracownik w danym dniu jest na zwolnieniu lekarskim
        '               5 jest nieobecny z innego powodu - to dostępne od wesrji 2.0.54
        '               -1 jeżeli błąd
        Dim str_data As String
        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim msg As String
        Dim data_em As Date
        Dim kontr As Boolean = False

        data_em = CDate(data_emisji)
        str_data = Format(data_em, "yyyy-MM-dd")

        sql = "SET DATEFORMAT YMD "
        sql = sql & " select rodzaj_nieobecnosci from nieobecnosci "
        sql = sql & " where data_rozpoczecia <= '" & str_data & "' and data_zakonczenia >= '" & str_data & "'"
        sql = sql & " and id_pracownika = " & id_pracownika

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                kontr = True
                wynik = dr.GetInt32(0)
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli obecności pracownika:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd = Nothing

        Catch ex As Exception

        End Try

        If kontr = True Then
            wynik = wynik + 1
        End If
        Return wynik


    End Function


    Public Function sprawdz_blokade_wnioskow(ByVal dp As Date, ByVal dk As Date, ByVal wybrany_program As String) As Integer

        'jako wybrany prohgram - nazwa programu
        'funkcja sprawdza czy w okresie podanym jakmo parametr
        'jest co najmniej jeden wniosek z ustawiona blokada edycji
        'jeżeli tak to zwraca 1 jeżeli nie to zwraca 0 jeżeli błąd to zwraca -1
        'uzywana podczas wystawiania nowych wniosków
        'aby zabezpieczyc przed wystaiwieniem nowego wniosku w sytuacji gdy pozostałe wnioski sa zablokowane do edycji 
        'tzn najprawdopodobniej wyeksportowane



        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim a As Integer
        Dim msg As String
        Dim wynik As Integer = 0
        Dim dp_str As String
        Dim dk_str As String
        Dim data_em As Date
        Dim id_wn As Integer

        Dim id_programu As Integer = 0
        Dim p As clsProgram

        If Len(wybrany_program) > 0 Then
            For Each p In colProgramy
                If p.nazwa_programu = wybrany_program Then
                    id_programu = p.id
                    Exit For
                End If
            Next
        End If


        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")



        sql = "SET DATEFORMAT YMD "
        sql = sql & " select id "
        sql = sql & " from wnioski"
        sql = sql & " WHERE data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' "
        sql = sql & " AND blokada_edycji=1"
        If Len(wybrany_program) > 0 Then
            sql = sql & " AND nr_programu =" & id_programu
        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then

                id_wn = dr.GetValue(0)

                wynik = 1
            End If


        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli stanu blokady wniosków:" & vbCrLf & ex.Message
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

    Public Function sprawdz_czy_wystawiono_wyceny(ByVal rok As String, ByVal id_pracownika As Integer, ByVal dp As Date, ByVal dk As Date) As Integer
        'funkcja używana jest podczas zapisu do tabeli nieobecności
        'sprawdza czy w zadanym okresie wystawiono wyceny pracownikowi
        'funkcja sprawdza czy w roczniku ROK
        ' są w okresie podanycm jako dwa odtsatnie parametwry
        'wyceny wystawione dla pracownika
        'zwraca 0 jeżeli nie ma 
        '1 - jeżeli są
        '-1 jeżeli błąd

        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim msg As String
        Dim dp_str As String
        Dim dk_str As String

        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & " SELECT "
        sql = sql & "wnioski.id_redakcji, "
        sql = sql & " pozycje_wniosku.identyfikator "
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' "
        sql = sql & " AND pozycje_wniosku.id_pracownika = " & id_pracownika



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                'jest rekord
                wynik = 1
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli wystawionych wycen pracownika:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd = Nothing

        Catch ex As Exception

        End Try

        Return wynik

    End Function

    Public Function ustal_aktualny_status_pracownika(ByVal id_pracownika As Integer, _
                                                        ByVal data_emisji As String, _
                                                        ByRef indywidualny_wsp_wyceny As Decimal) As Integer
        'funkcja zwraca atatus osoby ustalony na podstawie aktualnej umowy danej osoby
        'zwraca -1 - jeżeli nie ma umowy w danym dniu (umowy obowiązują na czas określony
        '       -2 - jeżeli błąd podczas sprawdzania

        Dim str_data As String
        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = -1
        Dim msg As String
        Dim data_em As Date

        data_em = CDate(data_emisji)

        str_data = Format(data_em, "yyyy-MM-dd")

        sql = "SET DATEFORMAT YMD "
        sql = sql & " select STATUS_PRACOWNIKA,INDYWIDUALNY_WSPOLCZYNNIK_WYCENY from umowy_pracownikow "
        sql = sql & " where data_rozpoczecia <= '" & str_data & "' and data_zakonczenia >= '" & str_data & "'"
        sql = sql & " and id_pracownika = " & id_pracownika

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                wynik = dr.GetInt32(0)
                If indywidualne_wsp_wyceny_dostepne Then
                    indywidualny_wsp_wyceny = dr.GetValue(1)
                Else
                    indywidualny_wsp_wyceny = 1
                End If
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas ustalania statusu pracownika:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -2

        End Try

        Try
            dr.Close()
            Cmd = Nothing

        Catch ex As Exception

        End Try

        Return wynik




    End Function
    Public Function ustal_id_rekordu_nowej_audycji(ByVal ramowka As Integer, _
                                                    ByVal id_redakcji As Integer, _
                                                    ByVal id_audycji As String) As Integer

        'funkcja zwraca id rekodru audycji 
        ' króej odpowiednie pola zawierają to co w parametrach


        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim msg As String

        sql = " select id from audycje "
        sql = sql & "where id_ramowki=" & ramowka
        sql = sql & " AND id_redakcji=" & id_redakcji
        sql = sql & " AND IDENTYFIKATOR_AUDYCJI='" & id_audycji & "'"


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                wynik = dr.GetInt32(0)

            End If

            dr.Close()
            Cmd = Nothing

        Catch ex As Exception
            msg = "Wystąpił problem podczas ustalania wewnętrznego identyfikatora nowej audycji:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try


        Return wynik




    End Function

    Public Function ustal_id_aktualnej_ramowki(ByVal data As Date, _
                                               ByRef nazwa_ramowki As String) As Integer
        'funkcja zwraca id_ramowki obowiązującej w dniu podanym jako data
        'przez referencje przekazywana jest w wyniku nazwa ramowki

        Dim str_data As String
        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim msg As String

        str_data = Format(data, "yyyy-MM-dd")

        sql = "SET DATEFORMAT YMD "
        sql = sql & " select id, nazwa_ramowki from ramowki where data_poczatkowa <= '" & str_data & "' and data_koncowa >= '" & str_data & "'"

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                wynik = dr.GetInt32(0)
                nazwa_ramowki = dr.GetValue(1)
            End If
            dr.Close()
            Cmd = Nothing

        Catch ex As Exception
            msg = "Wystąpił problem podczas ustalania identyfikatora aktualnej ramówki:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try


        Return wynik


    End Function
    Public Function zapisz_pozycje_tabeli_wycen(ByVal id_audycji As Integer, ByRef p As clsPozycjaTabeliWycen) As Integer

        Dim sql As String
        Dim id As Integer
        Dim a As Integer

        Dim msg As String
        Dim wymag As Integer = 0
        Dim wprac As Decimal
        Dim ww As Decimal
        Dim wprod As Decimal
        Dim wprac_str As String
        Dim ww_str As String
        Dim wprod_str As String
        Dim ggww_str As String
        Dim dgww_str As String
        Dim dgww As Single
        Dim ggww As Single
        Dim ku As Integer = 0

        Dim nazwa As String
        ' Dim uwagi As String



        If id_audycji < 1 Then
            msg = "Nieprawidłowo określona audycja."
            MessageBox.Show(msg)
            Return -1
        End If

        id = p.id_rekordu

        wprac = p.wycena_pracownika
        ww = p.wycena_wspolpracownika
        wprod = p.wycena_producenta

        wprac_str = formatuj_wycene(wprac)
        ww_str = formatuj_wycene(ww)
        wprod_str = formatuj_wycene(wprod)

        dgww = p.dgww
        ggww = p.ggww

        If licencjobiorca = "Radio Kielce SA" Then
            dgww_str = formatuj_wspolczynnik(dgww)
            ggww_str = formatuj_wspolczynnik(ggww)
        Else
            dgww_str = formatuj_wycene(dgww)
            ggww_str = formatuj_wycene(ggww)
        End If

        If tryb_obslugi_kosztow_uzysku > 0 Then
            ku = p.koszty_uzysku
            'tu na wszeli wypadek
            If tryb_obslugi_kosztow_uzysku < 2 And ku = 2 Then
                ku = 0
            End If
        Else
            ku = 0
        End If

        nazwa = skoryguj_apostrofy_do_SQL(p.nazwa_pozycji)

        If id = 0 Then
            sql = "INSERT INTO TABELA_WYCEN_AUDYCJI "
            sql = sql & "(nazwa_wyceny, "
            sql = sql & "id_audycji, "
            sql = sql & "id_pozycji, "
            sql = sql & "wycena_pracownika, "
            sql = sql & "wycena_wspolpracownika, "
            sql = sql & "wycena_producenta, "
            If sprawozdania_programowe_dostepne Then
                sql = sql & "WYMAGANIA_SPRAWOZDAN_PROGRAMOWYCH, "
            End If
            sql = sql & "rodzaj_wyceny, "
            If tryb_regulacji_wsp_wycen > 0 Then
                sql = sql & "DOLNA_GRANICA_WSP_WYCENY, "
                sql = sql & "GORNA_GRANICA_WSP_WYCENY, "
            End If
            If tryb_obslugi_kosztow_uzysku > 0 Then
                sql = sql & "koszty_uzysku, "
            End If

            If regulacja_wg_ustalonych_stawek Then
                sql = sql & "stawki_wyceny_p, "
                sql = sql & "stawki_wyceny_w, "
                sql = sql & "stawki_wyceny_pro, "
            End If




            sql = sql & "uwagi) "


            sql = sql & " VALUES('" & nazwa & "',"
            sql = sql & id_audycji & ","
            sql = sql & "'" & p.identyfikator & "',"
            '            sql = sql & o.txtWycenaPracownika.Text & ","
            '            sql = sql & o.txtWycenaWspolpracownika.Text & ","
            '            sql = sql & o.txtWycenaProducenta.Text & ","
            sql = sql & wprac_str & ","
            sql = sql & ww_str & ","
            sql = sql & wprod_str & ","

            If sprawozdania_programowe_dostepne Then
                sql = sql & p.wymagania_sprawozdan_programowych & ","

            End If
            sql = sql & p.rodzaj_wyceny & ","

            If tryb_regulacji_wsp_wycen > 0 Then
                sql = sql & dgww_str & ", "
                sql = sql & ggww_str & ", "
            End If
            If tryb_obslugi_kosztow_uzysku > 0 Then
                sql = sql & ku & ", "
            End If
            If regulacja_wg_ustalonych_stawek Then
                sql = sql & "'" & Trim(p.stawki_p) & "', "
                sql = sql & "'" & Trim(p.stawki_w) & "', "
                sql = sql & "'" & Trim(p.stawki_pr) & "', "
            End If


            sql = sql & "'-')"

        Else
            sql = "UPDATE TABELA_WYCEN_AUDYCJI set "
            sql = sql & " nazwa_wyceny='" & nazwa & "', "
            sql = sql & " id_pozycji='" & p.identyfikator & "', "
            sql = sql & " wycena_pracownika=" & wprac_str & ", " ' o.txtWycenaPracownika.Text & ", "
            sql = sql & " wycena_wspolpracownika=" & ww_str & ", " ' o.txtWycenaWspolpracownika.Text & ", "
            sql = sql & " wycena_producenta=" & wprod_str & "," 'o.txtWycenaProducenta.Text & ", "

            If sprawozdania_programowe_dostepne Then
                sql = sql & " WYMAGANIA_SPRAWOZDAN_PROGRAMOWYCH=" & p.wymagania_sprawozdan_programowych & ", "

            End If
            sql = sql & "rodzaj_wyceny=" & p.rodzaj_wyceny & ","

            If tryb_regulacji_wsp_wycen > 0 Then
                sql = sql & "DOLNA_GRANICA_WSP_WYCENY=" & dgww_str & ", "
                sql = sql & "GORNA_GRANICA_WSP_WYCENY=" & ggww_str & ", "
            End If
            If tryb_obslugi_kosztow_uzysku > 0 Then
                sql = sql & " koszty_uzysku=" & ku & ", "
            End If

            If regulacja_wg_ustalonych_stawek Then
                sql = sql & " stawki_wyceny_p= '" & Trim(p.stawki_p) & "', "
                sql = sql & " stawki_wyceny_w = '" & Trim(p.stawki_w) & "', "
                sql = sql & " stawki_wyceny_pro= '" & Trim(p.stawki_pr) & "', "
            End If

            sql = sql & " uwagi='-'"


            sql = sql & " where id=" & id
        End If



        a = wykonaj_polecenie_SQL(sql)


        Return a


    End Function
    Public Function zapisz_zmienione_stawki_ogolnej_tabeli_wycen(ByRef p As clsPozycjaTabeliWycen) As Integer
        Dim sql As String
        Dim id As Integer
        Dim a As Integer

        Dim wprac As Decimal
        Dim ww As Decimal
        Dim wprod As Decimal
        Dim wprac_str As String
        Dim ww_str As String
        Dim wprod_str As String

        id = p.id_rekordu

        wprac = p.wycena_pracownika
        ww = p.wycena_wspolpracownika
        wprod = p.wycena_producenta

        wprac_str = formatuj_wycene(wprac)
        ww_str = formatuj_wycene(ww)
        wprod_str = formatuj_wycene(wprod)


        sql = "UPDATE TABELA_WYCEN_OGOLNA set "
        sql = sql & " wycena_pracownika=" & wprac_str & ", "
        sql = sql & " wycena_wspolpracownika=" & ww_str & ", "
        sql = sql & " wycena_producenta=" & wprod_str

        sql = sql & " where id=" & id



        a = wykonaj_polecenie_SQL(sql)
        Return a

    End Function

    Public Function zapisz_zmienione_stawki_prywatnej_tabeli_wycen(ByRef p As clsPozycjaTabeliWycen) As Integer
        Dim sql As String
        Dim id As Integer
        Dim a As Integer

        Dim wprac As Decimal
        Dim ww As Decimal
        Dim wprod As Decimal
        Dim wprac_str As String
        Dim ww_str As String
        Dim wprod_str As String

        id = p.id_rekordu

        wprac = p.wycena_pracownika
        ww = p.wycena_wspolpracownika
        wprod = p.wycena_producenta

        wprac_str = formatuj_wycene(wprac)
        ww_str = formatuj_wycene(ww)
        wprod_str = formatuj_wycene(wprod)


        sql = "UPDATE TABELA_WYCEN_PRACOWNIKOW set "
        sql = sql & " wycena_pracownika=" & wprac_str & ", "
        sql = sql & " wycena_wspolpracownika=" & ww_str & ", "
        sql = sql & " wycena_producenta=" & wprod_str

        sql = sql & " where id=" & id



        a = wykonaj_polecenie_SQL(sql)
        Return a


    End Function
    Public Function zapisz_zmienione_stawki_tabeli_wycen(ByRef p As clsPozycjaTabeliWycen) As Integer
        Dim sql As String
        Dim id As Integer
        Dim a As Integer

        Dim wprac As Decimal
        Dim ww As Decimal
        Dim wprod As Decimal
        Dim wprac_str As String
        Dim ww_str As String
        Dim wprod_str As String

        id = p.id_rekordu

        wprac = p.wycena_pracownika
        ww = p.wycena_wspolpracownika
        wprod = p.wycena_producenta

        wprac_str = formatuj_wycene(wprac)
        ww_str = formatuj_wycene(ww)
        wprod_str = formatuj_wycene(wprod)


        sql = "UPDATE TABELA_WYCEN_AUDYCJI set "
        sql = sql & " wycena_pracownika=" & wprac_str & ", "
        sql = sql & " wycena_wspolpracownika=" & ww_str & ", "
        sql = sql & " wycena_producenta=" & wprod_str

        sql = sql & " where id=" & id



        a = wykonaj_polecenie_SQL(sql)
        Return a

    End Function

    Public Function zaladuj_wniosek_honoracyjny(ByVal id As Integer, ByRef opis As clsWniosekHonoracyjny) As Integer
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim k As Integer

        sql = "select "
        sql = sql & "id, "
        sql = sql & "id_redakcji, "
        sql = sql & "id_audycji, "
        sql = sql & "id_audycji_tabeli_wycen, "
        sql = sql & "data_emisji, "
        sql = sql & "godzina_rozpoczecia, "
        sql = sql & "godzina_zakonczenia, "
        sql = sql & "tytul_audycji, "
        sql = sql & "id_autora, "
        sql = sql & "imie_nazwisko_autora, "
        sql = sql & "wystawiony, "
        sql = sql & "imie_nazwisko_osoby_wystawiajacej, "
        sql = sql & "zaakceptowany_przez_kierownika, "
        sql = sql & "imie_nazwisko_kierownika, "
        sql = sql & "zaakceptowany_przez_szefa_programu, "
        sql = sql & "imie_nazwisko_szefa_programu, "
        sql = sql & "zaakceptowany_przez_zarzad, "
        sql = sql & "imie_nazwisko_zarzad, "
        sql = sql & "mpk, "
        sql = sql & "zrodlo_finansowania, "
        sql = sql & "ograniczenie_kosztow_honoracyjnych, "
        sql = sql & "ograniczenie_kosztow_pozahonoracyjnych, "
        sql = sql & "zadanie "

        If rozszerzone_sprawozdania_dostepne Then
            sql = sql & ", info "
            sql = sql & ", rodzaj_muzyki "
            sql = sql & ", rodzaj_dokumentacji "
            sql = sql & ", dlugosc_muzyki "

        End If
        If wstepne_zatwierdzanie_dostepne Then
            sql = sql & ", przygotowany_do_emisji, "
            sql = sql & "imie_nazwisko_autora2, "
            sql = sql & "akceptacja_wstepna_inspektora, "
            sql = sql & "imie_nazwisko_inspektora, "
            sql = sql & "akceptacja_wstepna_kierownika, "
            sql = sql & "imie_nazwisko_kierownika2, "
            sql = sql & "akceptacja_wstepna_szefa, "
            sql = sql & "imie_nazwisko_szefa_programu2 "
        End If
        If rozliczanie_minutowe_audycji_dostepne Then
            sql = sql & ",dlugosc "
        End If
        sql = sql & ", nr_programu,"
        sql = sql & " pasmo, "
        sql = sql & " grupa_docelowa,"
        sql = sql & " sygnatura_archiwalna,"
        sql = sql & " rodzaj_realizacji "

        sql = sql & "from wnioski "
        sql = sql & " where id= " & id


        wskaznik_myszy(1)

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować sie z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                opis.id = dr.GetValue(0)
                opis.identyfikator_redakcji = dr.GetValue(1)
                opis.id_audycji = dr.GetValue(2)
                opis.id_rekordu_w_ramowce = dr.GetValue(3)
                opis.data_emisji = dr.GetValue(4)
                opis.godzina_rozpoczecia = dr.GetValue(5)
                opis.godzina_zakonczenia = dr.GetValue(6)
                opis.tytul_audycji = dr.GetValue(7)
                opis.id_autora = dr.GetValue(8)
                opis.imie_nazwisko_autora = dr.GetValue(9)
                opis.wystawiony = dr.GetValue(10)
                opis.imie_nazwisko_osoby_wystawiajacej = dr.GetValue(11)
                opis.akceptacja_kierownika = dr.GetValue(12)
                opis.imie_nazwisko_kierownika = dr.GetValue(13)
                opis.akceptacja_szefa_programu = dr.GetValue(14)
                opis.imie_nazwisko_szefa_programu = dr.GetValue(15)
                opis.akceptacja_zarzadu = dr.GetValue(16)
                opis.imie_nazwisko_zarzadu = dr.GetValue(17)
                opis.kod_mpk = dr.GetValue(18)
                opis.zrodlo_finansowania = dr.GetValue(19)
                opis.gg_kosztow_honoracyjnych = dr.GetValue(20)
                opis.gg_kosztow_pozahonoracyjnych = dr.GetValue(21)
                opis.kod_zadania = dr.GetValue(22)
                k = 22
                If rozszerzone_sprawozdania_dostepne Then
                    k = k + 1
                    opis.info = dr.GetValue(k)
                    k = k + 1
                    opis.rodzaj_muzyki = dr.GetValue(k)
                    k = k + 1
                    opis.rodzaj_dokumentacji = dr.GetValue(k)
                    k = k + 1
                    opis.dlugosc_muzyki = dr.GetValue(k)
                End If
                If wstepne_zatwierdzanie_dostepne Then
                    k = k + 1
                    opis.przygotowany_do_emisji = dr.GetValue(k)
                    k = k + 1
                    opis.imie_nazwisko_autora2 = dr.GetValue(k)
                    k = k + 1
                    opis.akceptacja_inspektora = dr.GetValue(k)
                    k = k + 1
                    opis.imie_nazwisko_inspektora = dr.GetValue(k)
                    k = k + 1
                    opis.wstepna_akceptacja_kierownika = dr.GetValue(k)
                    k = k + 1
                    opis.imie_nazwisko_kierownika2 = dr.GetValue(k)
                    k = k + 1
                    opis.wstepna_akceptacja_szefa_programu = dr.GetValue(k)
                    k = k + 1
                    opis.imie_nazwisko_szefa_programu2 = dr.GetValue(k)
                End If
                If rozliczanie_minutowe_audycji_dostepne Then
                    k += 1
                    opis.dlugosc = dr.GetValue(k)
                End If
                k += 1
                opis.nr_programu = dr.GetValue(k)
                k += 1
                opis.id_pasma = dr.GetValue(k)
                k += 1
                Try
                    opis.grupa_docelowa = Trim(dr.GetValue(k))

                Catch ex13 As Exception
                    opis.grupa_docelowa = "-"
                End Try
                Try
                    k += 1
                    opis.sygnatura_archiwalna = dr.GetValue(k)

                Catch ex As Exception

                End Try
                k += 1
                opis.rodzaj_realizacji = dr.GetValue(k)

                odczytano = True
            Else
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception

            msg = "Wystąpił problem podczas ładowania pełnego opisu wniosku: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If odczytano = True Then
            wynik = 0
        Else
            wynik = -1
        End If

        Return wynik


    End Function

    Public Function zapisz_nowe_pasmo_wniosku(ByVal id_wniosku As Integer, ByVal pasmo As Integer) As Integer

        Dim sql As String
        Dim a As Integer

        sql = "update wnioski "
        sql = sql & " set pasmo=" & pasmo
        sql = sql & " where id=" & id_wniosku

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function

    Public Function spec_zapisz_info_o_zamknieciu_wniosku(ByVal id As Integer) As Integer
        'funkcja wywoływana podczas zamykania programu przez system windows
        'np przy wylogowywaniu
        'gdy wyświetlone jest niemodalne okno wniosku i system zamyka okno wniosku
        'nie wykonywana jest procedura zapisu informacji o zamknieciu wniosku
        'ta funkcja jest wykonywana za kazdym razem przy zamykaniu programu
        'ale jest zapis do bazy danych gdy ID>0

        Dim sql As String
        Dim a As Integer = 0
        Dim uzytk As String

        If id > 0 Then
            sql = "update wnioski"

            sql = sql & " set otwarty='False', "
            sql = sql & " kto_otworzyl=''"

            sql = sql & " where id=" & id

            a = wykonaj_polecenie_SQL(sql)

        End If

        Return a

    End Function


    Public Function zapisz_informacje_o_zamknieciu_wniosku_specjal(ByVal id As Integer) As Integer
        'funkcja służy tylko do odblokowywania oznaczonych jako edytowane
        'wniosków honoracyjnych
        Dim sql As String
        Dim a As Integer


        sql = "update wnioski"

        sql = sql & " set otwarty='False', "
        sql = sql & " kto_otworzyl=''"

        sql = sql & " where id=" & id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function


    Public Function zapisz_informacje_o_otwarciu_wniosku(ByRef opis As clsWniosekHonoracyjny, ByVal otwieranie As Boolean) As Integer
        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim uzytk As String

        uzytk = skoryguj_apostrofy_do_SQL(aktualny_uzytkownik.imie_nazwisko)

        rok = Year(opis.data_emisji)
        sql = "update wnioski"

        If otwieranie Then
            sql = sql & " set otwarty='True', "
            sql = sql & " kto_otworzyl='" & uzytk & "' "
        Else
            sql = sql & " set otwarty='False', "
            sql = sql & " kto_otworzyl=''"
        End If
        sql = sql & " where id=" & opis.id

        a = wykonaj_polecenie_SQL(sql)
        Return a

    End Function

    Public Function zapisz_nowa_granice_kosztow(ByVal tryb As Integer, ByRef opis As clsWniosekHonoracyjny, ByVal wartosc As Decimal) As Integer
        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim wartosc_str As String

        wartosc_str = formatuj_wycene(wartosc)

        rok = Year(opis.data_emisji)
        sql = "update wnioski "

        If tryb = 0 Then
            sql = sql & " set ograniczenie_kosztow_honoracyjnych= " & wartosc_str
        Else
            sql = sql & " set ograniczenie_kosztow_pozahonoracyjnych= " & wartosc_str
        End If
        sql = sql & " where id=" & opis.id

        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function
    Public Function zapisz_nowe_haslo(ByVal haslo As Integer, ByVal id_prac As Integer) As Integer
        Dim sql As String
        Dim a As Integer


        sql = "UPDATE pracownicy"
        sql = sql & " set password = " & haslo
        sql = sql & " where id =" & id_prac

        a = wykonaj_polecenie_SQL(sql)
        Return a

    End Function
    Public Function zapisz_nowe_haslo2(ByVal haslo As Integer, ByVal id_prac As Integer) As Integer
        'ta funkcja ustawia pole WYMAGANA_ZMIANA_HASLA na false

        Dim sql As String
        Dim a As Integer


        sql = "UPDATE pracownicy"
        sql = sql & " set password = " & haslo & ", "
        sql = sql & " wymagana_zmiana_hasla = 0 "

        sql = sql & " where id =" & id_prac


        a = wykonaj_polecenie_SQL(sql)
        Return a

    End Function


    Public Function zaladuj_uprawnienia_w_programach(ByVal id_pracownika As Integer, ByVal tryb As Integer) As Integer

        'tryb   0 - aktualnie zalogowany użytkownik
        '       1 - edytowany pracownik

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim upr As clsUprawnienie_w_programie
        If tryb = 0 Then
            If colUprawnieniaWProgramach.Count > 0 Then
                Do
                    colUprawnieniaWProgramach.Remove(1)
                Loop While colUprawnieniaWProgramach.Count > 0
            End If
        Else

            If colUprawnieniaProgrWybranego_pracownika.Count > 0 Then
                Do
                    colUprawnieniaProgrWybranego_pracownika.Remove(1)
                Loop While colUprawnieniaProgrWybranego_pracownika.Count > 0
            End If
        End If



        sql = "select "
        sql = sql & " id,"
        sql = sql & " id_programu,"
        sql = sql & " poziom_uprawnien "
        sql = sql & " from uprawnienia_w_programach where id_pracownika = " & id_pracownika

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    upr = New clsUprawnienie_w_programie
                    upr.id_rekordu = dr.GetValue(0)
                    upr.id_programu = dr.GetValue(1)
                    upr.poziom_uprawnien = dr.GetValue(2)
                    If tryb = 0 Then
                        colUprawnieniaWProgramach.Add(upr, upr.id_rekordu)
                    Else
                        colUprawnieniaProgrWybranego_pracownika.Add(upr, upr.id_rekordu)
                    End If
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu uprawnień pracownika w programach: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik



    End Function

    Public Function zaladuj_uprawnienia_w_redakcji(ByVal id_prac As Integer, ByVal tryb As Integer) As Integer
        'tryb   0 - aktualnie zalogowany użytkownik
        '       1 - edytowany pracownik

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim upr As clsUprawnienie_w_redakcji

        If tryb = 0 Then
            If colUprawnieniaWRedakcjach.Count > 0 Then
                Do
                    colUprawnieniaWRedakcjach.Remove(1)
                Loop While colUprawnieniaWRedakcjach.Count > 0
            End If
        Else

            If colUprawnieniaWRedWybranegoPracownika.Count > 0 Then
                Do
                    colUprawnieniaWRedWybranegoPracownika.Remove(1)
                Loop While colUprawnieniaWRedWybranegoPracownika.Count > 0
            End If
        End If



        sql = "select "
        sql = sql & " id,"
        sql = sql & " id_redakcji,"
        sql = sql & " poziom_uprawnien "
        sql = sql & " from uprawnienia_w_redakcjach where id_pracownika = " & id_prac

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    upr = New clsUprawnienie_w_redakcji
                    upr.id_rekordu = dr.GetValue(0)
                    upr.id_redakcji = dr.GetValue(1)
                    upr.poziom_uprawnien = dr.GetValue(2)
                    'od wersji 3 nie ma poziomu 3 - szef programu
                    'dlatego tu degradacja podczas odczytu

                    If upr.poziom_uprawnien > 2 Then
                        upr.poziom_uprawnien = 0
                    End If

                    If tryb = 0 Then
                        colUprawnieniaWRedakcjach.Add(upr, upr.id_rekordu)
                    Else
                        colUprawnieniaWRedWybranegoPracownika.Add(upr, upr.id_rekordu)
                    End If
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu uprawnień pracownika w redakcjach: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik


    End Function

    Public Function zaladuj_liste_usunietych_pracownikow() As Integer


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False

        'ta funkcja wywooływana jest na starcie programu
        'ale też przed każdym eksportem po to żeby wczytać aktualne pesele
        'dlatego wyczyszczenie colekcji
        If colUsunieciPracownicy.Count > 0 Then
            Do
                colUsunieciPracownicy.Remove(1)
            Loop While colUsunieciPracownicy.Count > 0
        End If

        sql = "select "
        sql = sql & " id, "
        sql = sql & " nazwisko, "
        sql = sql & " imie_nazwisko, "
        sql = sql & " rodzaj_pracownika,"
        sql = sql & " pesel,"
        sql = sql & " identyfikator_zewnetrzny,"
        sql = sql & " uprawnienia "
        sql = sql & " from pracownicy "
        sql = sql & " WHERE usuniety=1 "
        sql = sql & " order by nazwisko "

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    akt_prac = New clsPracownik
                    akt_prac.id = dr.GetValue(0)
                    akt_prac.nazwisko = dr.GetValue(1)
                    akt_prac.imie_nazwisko = dr.GetValue(2)
                    akt_prac.rodzaj_pracownika = dr.GetValue(3)
                    akt_prac.pesel = dr.GetValue(4)
                    akt_prac.identyfikator_zewnetrzny = dr.GetValue(5)
                    akt_prac.uprawnienia = dr.GetValue(6)
                    akt_prac.ustalono_status_pracownika = False
                    akt_prac.status_pracownika = 0
                    colUsunieciPracownicy.Add(akt_prac, akt_prac.id)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania rekordu pracownika: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik
    End Function

    Public Function sprawdz_czy_wpolpracownik_wewnetrzny(ByVal id_pracownika As Integer) As Integer
        'funkcja zwraca 0  - jeżeli nie współpracownik wewnetrzny
        '               1  - jeżeli współpracownik wewnętrzny
        '              -1  - jeżeli błąd podczas kontroli

        'funkcja używana tylko w Radiu rzeszów

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim kontr As Boolean = False

        sql = "select wspolpracownik_wewnetrzny from pracownicy"
        sql = sql & " where id=" & id_pracownika



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                kontr = dr.GetValue(0)
            End If
            If kontr Then
                wynik = 1 'jeżeli współpracownik wewnętrzny to wynik=1
            End If

            dr.Close()
            Cmd = Nothing

        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli statusu współpracownika:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try


        Return wynik



    End Function
    Public Function zaladuj_liste_pracownikow() As Integer


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim k As Integer

        'ta funkcja wywooływana jest na starcie programu
        'ale też przed każdym eksportem po to żeby wczytać aktualne pesele i identyfikatory zewnetrzne
        'dlatego wyczyszczenie colekcji
        If colPracownicy.Count > 0 Then
            Do
                colPracownicy.Remove(1)
            Loop While colPracownicy.Count > 0
        End If

        sql = "select "
        sql = sql & " id, "
        sql = sql & " nazwisko, "
        sql = sql & " imie_nazwisko, "
        sql = sql & " rodzaj_pracownika,"
        sql = sql & " pesel,"
        sql = sql & " identyfikator_zewnetrzny,"
        sql = sql & " uprawnienia "
        If wymuszanie_zmiany_hasla_dostepne Then
            sql = sql & " , wymagana_zmiana_hasla "
        End If

        If licencjobiorca = "Radio Dla Ciebie SA" Then
            sql = sql & " , blokada_eksportu_simple "
        End If

        If oznaczanie_wspolpracownikow_wewnetrznych_dostepne Then
            sql = sql & " , wspolpracownik_wewnetrzny "
        End If
        If tryb_sortowania_listy_osob > 0 Then
            sql = sql & " , nazwisko_do_sortowania "
        End If

        If oznaczanie_pracownikow_ryczaltowych_dostepne Then
            sql = sql & ", ryczalt "
        End If

        sql = sql & " from pracownicy "
        sql = sql & " WHERE usuniety=0 "

        If tryb_sortowania_listy_osob = 0 Then
            sql = sql & " order by nazwisko "
        Else
            sql = sql & " order by nazwisko_do_sortowania "
        End If

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    akt_prac = New clsPracownik
                    akt_prac.id = dr.GetValue(0)
                    akt_prac.nazwisko = dr.GetValue(1)
                    akt_prac.imie_nazwisko = dr.GetValue(2)
                    akt_prac.rodzaj_pracownika = dr.GetValue(3)
                    akt_prac.pesel = dr.GetValue(4)
                    akt_prac.identyfikator_zewnetrzny = dr.GetValue(5)
                    akt_prac.uprawnienia = dr.GetValue(6)
                    k = 6
                    If wymuszanie_zmiany_hasla_dostepne Then
                        k += 1
                        akt_prac.wymuszenie_zmiany_hasla = dr.GetValue(k)
                    End If
                    If licencjobiorca = "Radio Dla Ciebie SA" Then
                        k += 1
                        akt_prac.blokada_eksportu_SIMPLE = dr.GetValue(k)
                    End If
                    If oznaczanie_wspolpracownikow_wewnetrznych_dostepne Then
                        k += 1
                        akt_prac.wspolpracownik_wewnetrzny = dr.GetValue(k)
                    End If
                    If tryb_sortowania_listy_osob > 0 Then
                        k += 1
                        akt_prac.nazwisko_do_sortowania = dr.GetValue(k)
                    Else
                        akt_prac.nazwisko_do_sortowania = akt_prac.nazwisko
                    End If
                    If oznaczanie_pracownikow_ryczaltowych_dostepne Then
                        k += 1
                        akt_prac.ryczalt = dr.GetValue(k)
                    Else
                        akt_prac.ryczalt = False
                    End If
                    akt_prac.ustalono_status_pracownika = False
                    akt_prac.status_pracownika = 0
                    colPracownicy.Add(akt_prac, akt_prac.id)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania rekordu pracownika: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik
    End Function

    Public Function zaladuj_prywatna_tabele_wycen(ByVal id_pracownika As Integer) As Integer
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim poz As clsPozycjaTabeliWycen
        ' Dim aa As Integer


        If colPrywatnaTabelaWycen.Count > 0 Then
            Do
                colPrywatnaTabelaWycen.Remove(1)
            Loop While colPrywatnaTabelaWycen.Count > 0
        End If



        sql = "select "
        sql = sql & " id,"
        sql = sql & " id_pozycji,"
        sql = sql & " nazwa_wyceny,"
        sql = sql & " rodzaj_wyceny,"
        sql = sql & " wycena_pracownika,"
        sql = sql & " wycena_wspolpracownika,"
        sql = sql & " wycena_producenta, "
        sql = sql & " DOLNA_GRANICA_WSP_WYCENY, "
        sql = sql & " GORNA_GRANICA_WSP_WYCENY, "
        sql = sql & " WYMAGANIA_SPRAWOZDAN_PROGRAMOWYCH, "
        sql = sql & "koszty_uzysku "
        sql = sql & " from TABELA_WYCEN_PRACOWNIKOW where id_pracownika = " & id_pracownika
        sql = sql & " order by id_pozycji"

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    poz = New clsPozycjaTabeliWycen
                    poz.rodzaj_tabeli = 2 ' tabela prywatna pracownika
                    poz.id_rekordu = dr.GetValue(0)
                    poz.identyfikator = dr.GetValue(1)
                    poz.nazwa_pozycji = dr.GetValue(2)
                    poz.rodzaj_wyceny = dr.GetValue(3)
                    poz.wycena_pracownika = dr.GetValue(4)
                    poz.wycena_wspolpracownika = dr.GetValue(5)
                    poz.wycena_producenta = dr.GetValue(6)
                    poz.dgww = dr.GetValue(7)
                    poz.ggww = dr.GetValue(8)
                    poz.wymagania_sprawozdan_programowych = dr.GetValue(9)
                    If tryb_obslugi_kosztow_uzysku > 0 Then
                        poz.koszty_uzysku = dr.GetValue(10)
                    Else
                        poz.koszty_uzysku = 0 '50%
                    End If
                    colPrywatnaTabelaWycen.Add(poz, poz.identyfikator & "1")
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu tabeli wycen pracownika: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik




    End Function

    Public Function zaladuj_stale_pozycje_audycji(ByVal id As Integer) As Integer
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim poz As clsStalaPozycjaAudycji
        Dim aa As Integer


        If colStalePozycjeAktualnejAudycji.Count > 0 Then
            Do
                colStalePozycjeAktualnejAudycji.Remove(1)
            Loop While colStalePozycjeAktualnejAudycji.Count > 0
        End If



        sql = "select "
        sql = sql & " id,"
        sql = sql & " id_audycji,"
        sql = sql & " tytul,"
        sql = sql & " godzina_emisji,"
        sql = sql & " rodzaj,"
        sql = sql & " podrodzaj,"
        sql = sql & " dlugosc, "
        sql = sql & " forma_radiowa, "
        sql = sql & " rodzaj_licencji, "
        sql = sql & " region, "
        sql = sql & "powtorka, "
        sql = sql & "id_autora, "
        sql = sql & "imie_nazwisko_autora, "
        sql = sql & " rodzaj_audycji, "
        sql = sql & " rodzaj_produkcji "
       
        sql = sql & ", grupa_docelowa "
        If MPK_dostepne And tryb_obslugi_MPK = 1 Then
            sql = sql & ", mpk "
        End If

        sql = sql & " from stale_pozycje_audycji where id_audycji = " & id
        sql = sql & " order by godzina_emisji"

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    poz = New clsStalaPozycjaAudycji
                    poz.id = dr.GetValue(0)
                    poz.id_audycji = dr.GetValue(1)
                    poz.tytul = dr.GetValue(2)
                    poz.godzina_emisji = dr.GetValue(3)
                    poz.rodzaj = dr.GetValue(4)
                    poz.podrodzaj = dr.GetValue(5)
                    poz.dlugosc = dr.GetValue(6)
                    poz.forma_radiowa = dr.GetValue(7)
                    poz.rodzaj_licencji = dr.GetValue(8)
                    poz.region = dr.GetValue(9)
                    poz.powtorka = dr.GetValue(10)
                    poz.id_pracownika = dr.GetValue(11)
                    poz.imie_nazwisko_pracownika = dr.GetValue(12)
                    poz.rodzaj_audycji = dr.GetValue(13)
                    poz.rodzaj_produkcji = dr.GetValue(14)
                    Try

                        poz.grupa_docelowa = dr.GetValue(15)
                    Catch ex13 As Exception

                    End Try
                    If MPK_dostepne And tryb_obslugi_MPK = 1 Then
                        poz.mpk = dr.GetValue(16)
                    End If
                   

                    colStalePozycjeAktualnejAudycji.Add(poz, poz.id & "1")
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu stałych pozycji aktualnej audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik
    End Function
    Public Function zaladuj_ogolna_tabele_wycen_ramowki(ByVal id_ramowki As Integer) As Integer

        'funkcja używana przy zmianie stawek honoracyjnych ogólnej tabeli wycen w ramówce
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim poz As clsPozycjaTabeliWycen
        Dim aa As Integer
        Dim k As Integer


        sql = "select "
        sql = sql & " id,"
        sql = sql & " id_pozycji,"
        sql = sql & " nazwa_wyceny,"
        sql = sql & " rodzaj_wyceny,"
        sql = sql & " wycena_pracownika,"
        sql = sql & " wycena_wspolpracownika,"
        sql = sql & " wycena_producenta, "
        sql = sql & " DOLNA_GRANICA_WSP_WYCENY, "
        sql = sql & " GORNA_GRANICA_WSP_WYCENY, "
        sql = sql & " WYMAGANIA_SPRAWOZDAN_PROGRAMOWYCH, "
        sql = sql & "koszty_uzysku "

        sql = sql & " from TABELA_WYCEN_OGOLNA where id_ramowki = " & id_ramowki

        If colTabelaWycenCalejRamowki.Count > 0 Then
            Do
                colTabelaWycenCalejRamowki.Remove(1)
            Loop While colTabelaWycenCalejRamowki.Count > 0
        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    k += 1
                    poz = New clsPozycjaTabeliWycen
                    poz.rodzaj_tabeli = 1 ' tabela audycji ramówkowej
                    poz.id_rekordu = dr.GetValue(0)
                    poz.identyfikator = dr.GetValue(1)
                    poz.nazwa_pozycji = dr.GetValue(2)
                    poz.rodzaj_wyceny = dr.GetValue(3)
                    poz.wycena_pracownika = dr.GetValue(4)
                    poz.wycena_wspolpracownika = dr.GetValue(5)
                    poz.wycena_producenta = dr.GetValue(6)
                    poz.dgww = dr.GetValue(7)
                    poz.ggww = dr.GetValue(8)
                    poz.wymagania_sprawozdan_programowych = dr.GetValue(9)
                    If tryb_obslugi_kosztow_uzysku > 0 Then
                        poz.koszty_uzysku = dr.GetValue(10)
                    Else
                        poz.koszty_uzysku = 0 '50%
                    End If
                    colTabelaWycenCalejRamowki.Add(poz, k & "_")
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu ogólnej tabeli wycen z ramówki: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik



    End Function
    Public Function zaladuj_tabele_wycen_calej_ramowki(ByVal id_ramowki As Integer, ByVal id_redakcji As Integer) As Integer
        'funkcja używana przy zmianie stawek honoracyjnych tabel wycen audycji w całej ramówce
        'ładuje pozycje tabeli wycen wszystkuich audycji z ramóki o id podanym jako pierwszy parametr
        ' i dodatkowo tylko z tej redakcji która jes podana jako parametr
        ' ieżeli id redakcji =0 to ładowane są pozycje z wszystkich redakcji
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim poz As clsPozycjaTabeliWycen
        Dim aa As Integer
        Dim k As Integer


        sql = "select "
        sql = sql & " id,"
        sql = sql & " id_pozycji,"
        sql = sql & " nazwa_wyceny,"
        sql = sql & " rodzaj_wyceny,"
        sql = sql & " wycena_pracownika,"
        sql = sql & " wycena_wspolpracownika,"
        sql = sql & " wycena_producenta, "
        sql = sql & " DOLNA_GRANICA_WSP_WYCENY, "
        sql = sql & " GORNA_GRANICA_WSP_WYCENY, "
        sql = sql & " WYMAGANIA_SPRAWOZDAN_PROGRAMOWYCH, "
        sql = sql & "koszty_uzysku "
        sql = sql & " from TABELA_WYCEN_AUDYCJI where id_audycji IN  ("
        sql = sql & " select id from audycje where id_ramowki=" & id_ramowki
        If id_redakcji > 0 Then
            sql = sql & " AND id_redakcji= " & id_redakcji
        End If
        sql = sql & ")"

        If colTabelaWycenCalejRamowki.Count > 0 Then
            Do
                colTabelaWycenCalejRamowki.Remove(1)
            Loop While colTabelaWycenCalejRamowki.Count > 0
        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    k += 1
                    poz = New clsPozycjaTabeliWycen
                    poz.rodzaj_tabeli = 1 ' tabela audycji ramówkowej
                    poz.id_rekordu = dr.GetValue(0)
                    poz.identyfikator = dr.GetValue(1)
                    poz.nazwa_pozycji = dr.GetValue(2)
                    poz.rodzaj_wyceny = dr.GetValue(3)
                    poz.wycena_pracownika = dr.GetValue(4)
                    poz.wycena_wspolpracownika = dr.GetValue(5)
                    poz.wycena_producenta = dr.GetValue(6)
                    poz.dgww = dr.GetValue(7)
                    poz.ggww = dr.GetValue(8)
                    poz.wymagania_sprawozdan_programowych = dr.GetValue(9)
                    If tryb_obslugi_kosztow_uzysku > 0 Then
                        poz.koszty_uzysku = dr.GetValue(10)
                    Else
                        poz.koszty_uzysku = 0 '50%
                    End If
                    colTabelaWycenCalejRamowki.Add(poz, k & "_")
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu tabeli wycen z ramówki: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik

    End Function


    Public Function zaladuj_gorna_granice_kosztow_honor_audycji(ByVal id_aud As Integer) As Decimal
        Dim wynik As Decimal = 0

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim sql As String
        Dim kw As Decimal = 0
        Dim msg As String




        Try
            sql = "select ograniczenie_kosztow_honoracyjnych from audycje where id= " & id_aud

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader


            If dr.Read Then
                kw = dr.GetValue(0)
            End If
            wynik = kw

        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu górnej granicy kosztów honoracyjnych: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = 1 'tu celowo żeby nie mozna było zapisac kwoty większej niż 1 zł
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik


    End Function

    Public Function zaladuj_tabele_wycen_wniosku(ByVal id_audycji As Integer, ByVal data_emisji As Date, ByVal tryb As Integer) As Integer
        'tryb   0 bez łądowania ogólnej tabeli wycen (potrzebne podczas powielania tabeli wycen przy edycji ram owki audycji)
        '       1 - z łądowaniem ogólnej tablei wycen

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim poz As clsPozycjaTabeliWycen
        Dim aa As Integer


        If colTabelaWycenAktualnegoWNiosku.Count > 0 Then
            Do
                colTabelaWycenAktualnegoWNiosku.Remove(1)
            Loop While colTabelaWycenAktualnegoWNiosku.Count > 0
        End If



        sql = "select "
        sql = sql & " id,"
        sql = sql & " id_pozycji,"
        sql = sql & " nazwa_wyceny,"
        sql = sql & " rodzaj_wyceny,"
        sql = sql & " wycena_pracownika,"
        sql = sql & " wycena_wspolpracownika,"
        sql = sql & " wycena_producenta, "
        sql = sql & " DOLNA_GRANICA_WSP_WYCENY, "
        sql = sql & " GORNA_GRANICA_WSP_WYCENY, "
        sql = sql & " WYMAGANIA_SPRAWOZDAN_PROGRAMOWYCH, "
        sql = sql & "koszty_uzysku "
        If regulacja_wg_ustalonych_stawek Then
            sql = sql & ", stawki_wyceny_p "
            sql = sql & ", stawki_wyceny_w "
            sql = sql & ", stawki_wyceny_pro "
        End If

        sql = sql & " from TABELA_WYCEN_AUDYCJI where id_audycji = " & id_audycji
        sql = sql & " order by id_pozycji"

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    poz = New clsPozycjaTabeliWycen
                    poz.rodzaj_tabeli = 1 ' tabela audycji ramówkowej
                    poz.id_rekordu = dr.GetValue(0)
                    poz.identyfikator = dr.GetValue(1)
                    poz.nazwa_pozycji = dr.GetValue(2)
                    poz.rodzaj_wyceny = dr.GetValue(3)
                    poz.wycena_pracownika = dr.GetValue(4)
                    poz.wycena_wspolpracownika = dr.GetValue(5)
                    poz.wycena_producenta = dr.GetValue(6)
                    poz.dgww = dr.GetValue(7)
                    poz.ggww = dr.GetValue(8)
                    poz.wymagania_sprawozdan_programowych = dr.GetValue(9)
                    If tryb_obslugi_kosztow_uzysku > 0 Then
                        poz.koszty_uzysku = dr.GetValue(10)
                    Else
                        poz.koszty_uzysku = 0 '50%
                    End If
                    If regulacja_wg_ustalonych_stawek Then
                        poz.stawki_p = dr.GetValue(11)
                        poz.stawki_w = dr.GetValue(12)
                        poz.stawki_pr = dr.GetValue(13)
                    End If

                    colTabelaWycenAktualnegoWNiosku.Add(poz, poz.identyfikator & "1")
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu tabeli wycen aktualnej audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try
        If tryb > 0 Then
            'ogólna tabele wycen trzeba odświeżac bo w różnych datach emisji mogą obowiązywać różne ramówki
            aa = zaladuj_ogolna_tabele_wycen(data_emisji)
            'dodanie do listy pozycji z ogólnej tabeli wycen
            For Each poz In colOgolnaTabelaWycen
                colTabelaWycenAktualnegoWNiosku.Add(poz, poz.identyfikator & "0")
            Next
        End If

        Return wynik


    End Function

    Public Function zaladuj_ogolna_tabele_wycen(ByVal data_emisji As Date) As Integer
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim poz As clsPozycjaTabeliWycen
        Dim nr As String = "nazwa ramowki"
        Dim id_ramowki As Integer = 0
        If colOgolnaTabelaWycen.Count > 0 Then
            Do
                colOgolnaTabelaWycen.Remove(1)
            Loop While colOgolnaTabelaWycen.Count > 0
        End If

        id_ramowki = ustal_id_aktualnej_ramowki(data_emisji, nr)

        If id_ramowki < 1 Then
            Return 0
        End If

        sql = "select "
        sql = sql & " id,"
        sql = sql & " id_pozycji,"
        sql = sql & " nazwa_wyceny,"
        sql = sql & " rodzaj_wyceny,"
        sql = sql & " wycena_pracownika,"
        sql = sql & " wycena_wspolpracownika,"
        sql = sql & " wycena_producenta, "
        sql = sql & " DOLNA_GRANICA_WSP_WYCENY, "
        sql = sql & " GORNA_GRANICA_WSP_WYCENY, "
        sql = sql & " WYMAGANIA_SPRAWOZDAN_PROGRAMOWYCH, "
        sql = sql & "koszty_uzysku "
        If regulacja_wg_ustalonych_stawek Then
            sql = sql & ", stawki_wyceny_p "
            sql = sql & ", stawki_wyceny_w "
            sql = sql & ", stawki_wyceny_pro "
        End If

        sql = sql & " from TABELA_WYCEN_OGOLNA where id_ramowki = " & id_ramowki
        sql = sql & " order by id_pozycji"

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    poz = New clsPozycjaTabeliWycen
                    poz.rodzaj_tabeli = 0 ' ogólna tabela wycen
                    poz.id_rekordu = dr.GetValue(0)
                    poz.identyfikator = dr.GetValue(1)
                    poz.nazwa_pozycji = dr.GetValue(2)
                    poz.rodzaj_wyceny = dr.GetValue(3)
                    poz.wycena_pracownika = dr.GetValue(4)
                    poz.wycena_wspolpracownika = dr.GetValue(5)
                    poz.wycena_producenta = dr.GetValue(6)
                    poz.dgww = dr.GetValue(7)
                    poz.ggww = dr.GetValue(8)
                    poz.wymagania_sprawozdan_programowych = dr.GetValue(9)
                    If tryb_obslugi_kosztow_uzysku > 0 Then
                        poz.koszty_uzysku = dr.GetValue(10)
                    Else
                        poz.koszty_uzysku = 0 '50%
                    End If
                    If regulacja_wg_ustalonych_stawek Then
                        poz.stawki_p = dr.GetValue(11)
                        poz.stawki_w = dr.GetValue(12)
                        poz.stawki_pr = dr.GetValue(13)
                    End If

                    colTabelaWycenAktualnegoWNiosku.Add(poz, poz.identyfikator & "0")
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu ogólnej tabeli wycen: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik




    End Function
    Public Function zaladuj_nieobecnosci_pracownika(ByVal id As Integer) As Integer
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        'Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim np As clsNieobecnoscPracownika


        If colNIeobecnosciPracownika.Count > 0 Then
            Do
                colNIeobecnosciPracownika.Remove(1)
            Loop While colNIeobecnosciPracownika.Count > 0
        End If



        sql = "select "
        sql = sql & " id,"
        sql = sql & " id_pracownika,"
        sql = sql & " data_rozpoczecia,"
        sql = sql & " data_zakonczenia,"
        sql = sql & " rodzaj_nieobecnosci,"
        sql = sql & " uwagi "
        sql = sql & " from nieobecnosci where id_pracownika = " & id
        sql = sql & " order by data_rozpoczecia desc"

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    np = New clsNieobecnoscPracownika
                    np.id_rekordu = dr.GetValue(0)
                    np.id_pracownika = dr.GetValue(1)
                    np.data_poczatkowa = dr.GetValue(2)
                    np.data_koncowa = dr.GetValue(3)
                    np.rodzaj_nieobecnosci = dr.GetValue(4)
                    np.uwagi = dr.GetValue(5)
                    colNIeobecnosciPracownika.Add(np, np.id_rekordu)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu spisu nieobecności pracownika: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik

    End Function

    Public Function zaladuj_umowy_pracownika(ByVal id As Integer) As Integer
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        'Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim um As clsUmowa


        If colUmowyPracownika.Count > 0 Then
            Do
                colUmowyPracownika.Remove(1)
            Loop While colUmowyPracownika.Count > 0
        End If



        sql = "select "
        sql = sql & " id_umowy,"
        sql = sql & " id_pracownika,"
        sql = sql & " data_rozpoczecia,"
        sql = sql & " data_zakonczenia,"
        sql = sql & " status_pracownika,"
        sql = sql & " indywidualny_wspolczynnik_wyceny,"
        sql = sql & " tresc_umowy, "
        sql = sql & " uwagi "
        sql = sql & " from UMOWY_PRACOWNIKOW where id_pracownika = " & id
        sql = sql & " order by data_rozpoczecia desc"

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    um = New clsUmowa
                    um.id_rekordu = dr.GetValue(0)
                    um.id_pracownika = dr.GetValue(1)
                    um.data_poczatkowa = dr.GetValue(2)
                    um.data_koncowa = dr.GetValue(3)
                    um.status = dr.GetValue(4)
                    um.indywidualny_wspolczynnik_wyceny = dr.GetValue(5)
                    um.tresc = dr.GetValue(6)
                    um.uwagi = dr.GetValue(7)
                    colUmowyPracownika.Add(um, um.id_rekordu)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu spisu umów pracownika: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik

    End Function
    Public Function sprawdz_czy_wniosek_jest_juz_otwarty(ByRef opis As clsWniosekHonoracyjny, ByRef kto_edytuje As String) As Integer
        'funkcja zwraca wartość 0 jeżeli nikt nie edytuje wnioisku
        '                       1 jeżeli edytuje
        ' podstawia tez referencyjna zmienną KTO_EDYTUJE
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik_bol As Boolean = False
        Dim wynik As Integer
        Dim odczytano As Boolean = False
        Dim rok As String
        Dim nazwisko As String = ""

        rok = Year(opis.data_emisji)

        sql = "select "
        sql = sql & "id, "
        sql = sql & "otwarty, "
        sql = sql & "kto_otworzyl "

        sql = sql & "from wnioski "
        sql = sql & " where id= " & opis.id

        'MessageBox.Show(sql)

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                wynik_bol = dr.GetValue(1)
                nazwisko = dr.GetValue(2)
                odczytano = True
            Else
                odczytano = False
            End If
        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli stanu otwarcia wniosku: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If odczytano = False Then
            wynik = -1
        Else
            If wynik_bol = True Then
                kto_edytuje = nazwisko
                wynik = 1
            Else
                wynik = 0
                kto_edytuje = ""
            End If
        End If

        Return wynik


    End Function
    Public Function sprawdz_czy_wniosek_jest_juz_otwarty(ByVal id As Integer, ByVal data_emisji As Date, ByRef kto_edytuje As String) As Integer

        'funkcja używana przed założeniem blokady edycji
        ' różni się sposoem przekazania parametru id wniosku
        'funkcja zwraca wartość 0 jeżeli nikt nie edytuje wnioisku
        '                       1 jeżeli edytuje
        ' podstawia tez referencyjna zmienną KTO_EDYTUJE

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik_bol As Boolean = False
        Dim wynik As Integer
        Dim odczytano As Boolean = False
        Dim rok As String
        Dim nazwisko As String = ""

        rok = Year(data_emisji)

        sql = "select "
        sql = sql & "id, "
        sql = sql & "otwarty, "
        sql = sql & "kto_otworzyl "

        sql = sql & "from wnioski"
        sql = sql & " where id= " & id

        'MessageBox.Show(sql)

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                wynik_bol = dr.GetValue(1)
                nazwisko = dr.GetValue(2)
                odczytano = True
            Else
                odczytano = False
            End If
        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli stanu otwarcia wniosku: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If odczytano = False Then
            wynik = -1
        Else
            If wynik_bol = True Then
                kto_edytuje = nazwisko
                wynik = 1
            Else
                wynik = 0
                kto_edytuje = ""
            End If
        End If

        Return wynik


    End Function
    Public Function odswiez_pola_zatwierdzenia_wniosku(ByRef opis As clsWniosekHonoracyjny) As Integer

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim rok As String

        rok = Year(opis.data_emisji)

        sql = "select "
        sql = sql & "id, "
        sql = sql & "wystawiony, "
        sql = sql & "zaakceptowany_przez_kierownika, "
        sql = sql & "zaakceptowany_przez_szefa_programu, "
        sql = sql & "zaakceptowany_przez_zarzad, "
        sql = sql & "blokada_edycji "

        sql = sql & "from wnioski"
        sql = sql & " where id= " & opis.id

        'MessageBox.Show(sql)

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                opis.wystawiony = dr.GetValue(1)
                opis.akceptacja_kierownika = dr.GetValue(2)
                opis.akceptacja_szefa_programu = dr.GetValue(3)
                opis.akceptacja_zarzadu = dr.GetValue(4)
                opis.blokada_edycji = dr.GetValue(5)
                odczytano = True
            Else
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli stanu zatwierdzenia wniosku: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If odczytano = True Then
            wynik = 0
        Else
            opis.wystawiony = True
            opis.akceptacja_kierownika = True
            opis.akceptacja_szefa_programu = True
            opis.akceptacja_zarzadu = True
            wynik = -1
        End If

        Return wynik

    End Function

    Public Function wczytaj_rekord_pracownika(ByVal login_name As String, _
                                                ByRef haslo As Integer, _
                                                ByRef id_rekordu As Integer, _
                                                ByRef imie_nazwisko As String, _
                                                ByRef uprawnienia As Integer, _
                                                ByRef wymag_zmiana_hasla As Boolean) As Integer


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim usuniety As Boolean = False

        sql = "select "
        sql = sql & "id, "
        sql = sql & "imie_nazwisko, "
        sql = sql & "password, "
        sql = sql & "uprawnienia, "
        sql = sql & "usuniety "
        If wymuszanie_zmiany_hasla_dostepne Then
            sql = sql & ", wymagana_zmiana_hasla "
        End If

        sql = sql & " from pracownicy "
        sql = sql & " where nazwisko= '" & login_name & "'"


        Try
            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                id_rekordu = dr.GetValue(0)
                imie_nazwisko = dr.GetValue(1)
                haslo = dr.GetValue(2)
                uprawnienia = dr.GetValue(3)
                usuniety = dr.GetValue(4)
                If wymuszanie_zmiany_hasla_dostepne Then
                    wymag_zmiana_hasla = dr.GetValue(5)
                End If
                odczytano = True
            Else
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu danych pracownika: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        If odczytano = True Then
            If usuniety = False Then
                wynik = 0
            Else
                wynik = -1
            End If
        Else
            wynik = -1
        End If

        Return wynik

    End Function
    Public Function db_wczytaj_koszty_zadania(ByVal id_zadania As Integer, _
                                                ByVal data_poczatkowa As Date, _
                                                ByVal data_koncowa As Date) As Decimal


        Dim sql As String
        Dim r1 As String
        Dim r2 As String
        Dim msg As String
        Dim dp As String
        Dim dk As String
        Dim aa As Decimal

        r1 = Year(data_poczatkowa)
        r2 = Year(data_koncowa)

        dp = Format(data_poczatkowa, "yyyy-MM-dd")
        dk = Format(data_koncowa, "yyyy-MM-dd")

        'If r1 <> r2 Then
        ' msg = "Możliwe jest wyznaczenie kosztów redakcji tylko w ramach jednego roku"
        ' MessageBox.Show(msg, naglowek_komunikatow)
        ' Return -1
        ' End If

        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
        sql = sql & " * pozycje_wniosku.ilosc "
        sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
        sql = sql & " AS Suma_wydatkow"
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE pozycje_wniosku.zadanie =" & id_zadania
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & dp & "' AND '" & dk & "'"

        aa = wczytaj_sume_kosztow(sql)

        Return aa

    End Function

    Public Function db_wczytaj_koszty_wg_zrodla(ByVal zrodlo As Integer, _
                                                     ByVal dp As Date, _
                                                     ByVal dk As Date) As Decimal

        Dim sql As String
        Dim r1 As String
        Dim r2 As String
        Dim msg As String
        Dim dp_str As String
        Dim dk_str As String
        Dim aa As Decimal

        r1 = Year(dp)
        r2 = Year(dk)

        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")

        ' If r1 <> r2 Then
        'msg = "Możliwe jest wyznaczenie kosztów tylko w ramach jednego roku"
        'MessageBox.Show(msg, naglowek_komunikatow)
        'Return -1
        'End If

        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
        sql = sql & " * pozycje_wniosku.ilosc "
        sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
        sql = sql & " AS Suma_wydatkow"
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE  wnioski.zrodlo_finansowania = " & zrodlo
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "'"

        aa = wczytaj_sume_kosztow(sql)

        Return aa



    End Function

    Public Function db_wczytaj_niehonoracyjne_koszty_redakcji(ByVal id_redakcji As String, _
                                                                ByVal data_poczatkowa As Date, _
                                                                ByVal data_koncowa As Date) As Decimal

        Dim sql As String
        '        Dim r1 As String
        '        Dim r2 As String
        '       Dim msg As String
        Dim dp As String
        Dim dk As String
        Dim aa As Decimal

        '        r1 = Year(data_poczatkowa)
        '        r2 = Year(data_koncowa)

        dp = Format(data_poczatkowa, "yyyy-MM-dd")
        dk = Format(data_koncowa, "yyyy-MM-dd")

        '       If r1 <> r2 Then
        ' msg = "Możliwe jest wyznaczenie kosztów redakcji tylko w ramach jednego roku"
        ' MessageBox.Show(msg, naglowek_komunikatow)
        ' Return -1
        'End If

        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT Sum(niehonoracyjne_pozycje_wniosku.stawka_podstawowa "
        sql = sql & " * niehonoracyjne_pozycje_wniosku.ilosc) "
        sql = sql & " AS Suma_wydatkow"
        sql = sql & " FROM wnioski RIGHT JOIN niehonoracyjne_pozycje_wniosku"
        sql = sql & " ON wnioski.id = niehonoracyjne_pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE wnioski.id_redakcji = '" & id_redakcji & "' "
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & dp & "' AND '" & dk & "'"

        aa = wczytaj_sume_kosztow(sql)

        Return aa

    End Function



    Public Function db_wczytaj_koszty_redakcji2(ByVal identyfikator As String, _
                                            ByVal data_poczatkowa As Date, _
                                            ByVal data_koncowa As Date) As Decimal
        'ta funkcja wywoływana jest podczas otwierania pozycji wniosku
        'do edycji - służy do budżetowanoa redakcji
        ' wywoływana jest też podczas wyznaczania kosztów redakcji 
        ' chodzi o to że zrealizowane koszty redakcji mają pokazywac te wydatki 
        '   któe odejmują się od budżetów
        ' a to zależy od trybu kontroli budżetów

        '        Public tryb_kontroli_budzetow_redakcji As Integer '0 - kontrola wyłączona
        '                                                 '1 kontrola wydatków finansowanych ze środków włąsnych
        '                                                 '2- kontrola wydatków na zadania o id < niż 10
        '                                                  3 - kontrola na zad <10 finansoa=wane ze srodków własnych 
        '                                                  4 - kontrola wszystkich  

        Dim sql As String
        Dim dp As String
        Dim dk As String
        Dim aa As Decimal


        dp = Format(data_poczatkowa, "yyyy-MM-dd")
        dk = Format(data_koncowa, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
        sql = sql & " * pozycje_wniosku.ilosc "
        sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
        sql = sql & " AS Suma_wydatkow"
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE pozycje_wniosku.identyfikator Like '" & identyfikator & "' "
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & dp & "' AND '" & dk & "' "
        If tryb_kontroli_budzetow_redakcji = 3 Then
            sql = sql & " and pozycje_wniosku.zadanie<10 and wnioski.zrodlo_finansowania=1 "
        ElseIf tryb_kontroli_budzetow_redakcji = 2 Then
            sql = sql & " and pozycje_wniosku.zadanie<10"
        ElseIf tryb_kontroli_budzetow_redakcji = 1 Then
            sql = sql & " and wnioski.zrodlo_finansowania = 1"
        End If

        aa = wczytaj_sume_kosztow(sql)

        Return aa

    End Function


    Public Function wyznacz_dlugosc_muzyki_wg_mpk(ByVal kod_mpk As String, _
                                                    ByVal data_poczatkowa As Date, _
                                                    ByVal data_koncowa As Date, _
                                                    ByVal nr_programu As Integer) As Integer
        'jeżeli nr_programu=0 to znaczy ze wyliczac dla wszystkich proramów

        Dim sql As String
        Dim r1 As String
        Dim r2 As String
        Dim msg As String
        Dim dp As String
        Dim dk As String
        Dim aa As Decimal
        Dim dl As Integer


        r1 = Year(data_poczatkowa)
        r2 = Year(data_koncowa)

        dp = Format(data_poczatkowa, "yyyy-MM-dd")
        dk = Format(data_koncowa, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & " SELECT DISTINCT Sum(dlugosc_muzyki) "
        sql = sql & " AS dlugosc"
        sql = sql & " FROM wnioski "
        sql = sql & " WHERE mpk='" & Trim(kod_mpk) & "' "
        sql = sql & " AND data_emisji BETWEEN '" & dp & "' AND '" & dk & "'"

        If nr_programu > 0 Then
            sql = sql & " AND wnioski.nr_programu=" & nr_programu
        End If


        aa = wczytaj_sume_kosztow(sql)

        dl = CInt(aa)

        Return dl



    End Function


    Public Function wyznacz_dlugosc_muzyki(ByVal rodzaj As String, ByVal data_poczatkowa As Date, ByVal data_koncowa As Date, ByVal nr_programu As Integer, ByVal id_redakcji As String) As Integer
        'jeżeli nr_programu=0 to znaczy ze wyliczac dla wszystkich proramów
        'jeżeli długość id_redakcji=0 to wyznaczac dla wszystkucj redakcnji

        Dim sql As String
        Dim r1 As String
        Dim r2 As String
        Dim msg As String
        Dim dp As String
        Dim dk As String
        Dim aa As Decimal
        Dim dl As Integer
        Dim tmp_id_redakcji As String = ""

        tmp_id_redakcji = Trim(id_redakcji)


        r1 = Year(data_poczatkowa)
        r2 = Year(data_koncowa)

        dp = Format(data_poczatkowa, "yyyy-MM-dd")
        dk = Format(data_koncowa, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & " SELECT DISTINCT Sum(dlugosc_muzyki) "
        sql = sql & " AS dlugosc"
        sql = sql & " FROM wnioski "
        sql = sql & " WHERE rodzaj_muzyki='" & rodzaj & "' "
        sql = sql & " AND data_emisji BETWEEN '" & dp & "' AND '" & dk & "'"

        If nr_programu > 0 Then
            sql = sql & " AND wnioski.nr_programu=" & nr_programu
        End If
        If Len(tmp_id_redakcji) > 0 Then
            sql = sql & " AND wnioski.id_redakcji='" & tmp_id_redakcji & "'"
        End If


        aa = wczytaj_sume_kosztow(sql)

        dl = CInt(aa)

        Return dl



    End Function

    Public Function wyznacz_dlugosc_muzyki2(ByVal d_em As Date, ByVal nr_programu As Integer)

        Dim sql As String
        Dim d_em_str As String
        Dim aa As Decimal
        Dim dl As Integer



        d_em_str = Format(d_em, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & " SELECT DISTINCT Sum(dlugosc_muzyki) "
        sql = sql & " AS dlugosc"
        sql = sql & " FROM wnioski "
        sql = sql & " WHERE data_emisji = '" & d_em_str & "'"
        If nr_programu > 0 Then
            sql = sql & " AND nr_programu=" & nr_programu
        End If

        aa = wczytaj_sume_kosztow(sql)

        dl = CInt(aa)

        Return dl


    End Function

    Public Function db_wczytaj_koszt_pasma(ByVal id_pasma As Integer, _
                                           ByVal id_programu As Integer, _
                                           ByVal data_poczatkowa As Date, _
                                           ByVal data_koncowa As Date) As Decimal

        Dim sql As String
        Dim r1 As String
        Dim r2 As String
        Dim msg As String
        Dim dp As String
        Dim dk As String
        Dim aa As Decimal

        r1 = Year(data_poczatkowa)
        r2 = Year(data_koncowa)

        dp = Format(data_poczatkowa, "yyyy-MM-dd")
        dk = Format(data_koncowa, "yyyy-MM-dd")

        '        If r1 <> r2 Then
        ' msg = "Możliwe jest wyznaczenie kosztów redakcji tylko w ramach jednego roku"
        ' MessageBox.Show(msg, naglowek_komunikatow)
        'Return -1
        'End If

        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
        sql = sql & " * pozycje_wniosku.ilosc "
        sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
        sql = sql & " AS Suma_wydatkow"
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE wnioski.pasmo=" & id_pasma
        If id_programu > 0 Then
            sql = sql & " AND wnioski.nr_programu = " & id_programu
        End If
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & dp & "' AND '" & dk & "'"

        aa = wczytaj_sume_kosztow(sql)

        Return aa


    End Function

    Public Function db_wczytaj_koszt_programu(ByVal id_programu As Integer, _
                                              ByVal data_poczatkowa As Date, _
                                              ByVal data_koncowa As Date) As Decimal


        Dim sql As String
        Dim r1 As String
        Dim r2 As String
        Dim msg As String
        Dim dp As String
        Dim dk As String
        Dim aa As Decimal

        r1 = Year(data_poczatkowa)
        r2 = Year(data_koncowa)

        dp = Format(data_poczatkowa, "yyyy-MM-dd")
        dk = Format(data_koncowa, "yyyy-MM-dd")

        '        If r1 <> r2 Then
        ' msg = "Możliwe jest wyznaczenie kosztów redakcji tylko w ramach jednego roku"
        ' MessageBox.Show(msg, naglowek_komunikatow)
        'Return -1
        'End If

        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
        sql = sql & " * pozycje_wniosku.ilosc "
        sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
        sql = sql & " AS Suma_wydatkow"
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE wnioski.nr_programu = " & id_programu
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & dp & "' AND '" & dk & "'"

        aa = wczytaj_sume_kosztow(sql)

        Return aa




    End Function


    Public Function db_wczytaj_koszty_redakcji(ByVal identyfikator As String, _
                                                ByVal data_poczatkowa As Date, _
                                                ByVal data_koncowa As Date) As Decimal

        ' ta funcja wywołuywana jest podcza kontroli kosztó audycji

        Dim sql As String
        Dim r1 As String
        Dim r2 As String
        Dim msg As String
        Dim dp As String
        Dim dk As String
        Dim aa As Decimal

        r1 = Year(data_poczatkowa)
        r2 = Year(data_koncowa)

        dp = Format(data_poczatkowa, "yyyy-MM-dd")
        dk = Format(data_koncowa, "yyyy-MM-dd")

        '        If r1 <> r2 Then
        ' msg = "Możliwe jest wyznaczenie kosztów redakcji tylko w ramach jednego roku"
        ' MessageBox.Show(msg, naglowek_komunikatow)
        'Return -1
        'End If

        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
        sql = sql & " * pozycje_wniosku.ilosc "
        sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
        sql = sql & " AS Suma_wydatkow"
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE pozycje_wniosku.identyfikator Like '" & identyfikator & "' "
        sql = sql & " AND wnioski.data_emisji BETWEEN '" & dp & "' AND '" & dk & "'"

        aa = wczytaj_sume_kosztow(sql)

        Return aa

    End Function

    Public Function wczytaj_sume_kosztow(ByVal sql As String) As Decimal
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim wynik As Decimal = 0
        Dim odczytano As Boolean = False



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                If dr.GetValue(0) IsNot System.DBNull.Value Then
                    wynik = dr.GetValue(0)
                Else
                    wynik = 0
                End If
            Else
                odczytano = False
            End If
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu sumy kosztów z bazy danych: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If odczytano = False Then
            wynik = -1
        End If

        Return wynik

    End Function

    Public Function zaladuj_rodzaje_programowe()
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        'Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim rodz As clsRodzajProgramowy



        If colRodzajeProgramowe.Count > 0 Then
            Do
                colRodzajeProgramowe.Remove(1)
            Loop While colRodzajeProgramowe.Count > 0
        End If

        rodz = New clsRodzajProgramowy
        rodz.id = 0
        rodz.rodzaj = "-"
        colRodzajeProgramowe.Add(rodz, rodz.id)


        sql = "select "
        sql = sql & " id,"
        sql = sql & " rodzaj, skrot "
        sql = sql & " from rodzaje_programowe"

        If licencjobiorca = "Radio Wrocław SA" Then
            sql = sql & " order by skrot"
        Else
            sql = sql & " order by rodzaj"
        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    rodz = New clsRodzajProgramowy
                    rodz.id = dr.GetValue(0)
                    rodz.rodzaj = dr.GetValue(1)
                    rodz.skrot = dr.GetValue(2)
                    colRodzajeProgramowe.Add(rodz, rodz.id)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu spisu rodzajów programowych: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik



    End Function


    Public Function ustal_rodzaj_programowy(ByVal podrodzaj As String) As String
        Dim wynik As String = "-"
        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim msg As String
        Dim odczytano As Boolean = False
        Dim rodz As String

        sql = "select rodzaj from rodzaje_programowe "
        sql = sql & " where id = (select id_rodzaju from podrodzaje_programowe where podrodzaj='" & podrodzaj & "')"

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True

                    rodz = dr.GetValue(0)

                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = rodz
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu rodzaju programowego: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = "-"
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik


    End Function
    Public Function zaladuj_wszystkie_podrodzaje_programowe() As Integer
        'ta funkkcja jest potrzebna jak na razie tylko podczas improtu danych z daccord
        'pierwsze 5 znaków z pola Program Area ma siepokrywać z pięcoma znakami z podrodzaju
        'na tej podstawie zostanie ustalony rodzaj


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        'Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim podrodz As String
        Dim rodz As clsRodzajProgramowy
        Dim id As Integer = 0
        Dim k As Integer

        wskaznik_myszy(1)

        colWszystkiepodrodzaje.Clear()



        sql = "select "
        sql = sql & " id,"
        sql = sql & " podrodzaj "
        sql = sql & " from podrodzaje_programowe "

        sql = sql & " order by podrodzaj"


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    k += 1
                    podrodz = dr.GetValue(1)
                    colWszystkiepodrodzaje.Add(podrodz, k & "_")

                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu spisu rodzajów programowych: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik


    End Function

    Public Sub zaladuj_podrodzaje_programowe(ByVal rodzaj As String, ByRef cbo As ComboBox)
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        'Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim podrodz As String
        Dim rodz As clsRodzajProgramowy
        Dim id As Integer = 0


        wskaznik_myszy(1)

        cbo.Items.Clear()

        cbo.Items.Add("-")

        For Each rodz In colRodzajeProgramowe
            If UCase(rodz.rodzaj) = UCase(rodzaj) Then
                id = rodz.id
                Exit For
            End If
        Next

        If id = 0 Then Return

        sql = "select "
        sql = sql & " id,"
        sql = sql & " podrodzaj "
        sql = sql & " from podrodzaje_programowe "

        sql = sql & " where id_rodzaju= " & id

        If licencjobiorca = "Radio Wrocław SA" Then
            sql = sql & " order by skrot"
        Else
            sql = sql & " order by podrodzaj"
        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    podrodz = dr.GetValue(1)
                    cbo.Items.Add(podrodz)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu spisu rodzajów programowych: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


    End Sub

    Public Function ustal_skrot_podrodzaju(ByVal rodzaj As String, ByVal podrodzaj As String, ByRef skrot As String) As Integer
        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim odczytano As Boolean = False
        Dim tmp_skrot As String = "-"
        Dim wynik As Integer
        Dim msg As String

        If Len(podrodzaj) > 2 Then
            tmp_skrot = Left(podrodzaj, 3)
        Else
            tmp_skrot = podrodzaj
        End If

        sql = sql & "   SELECT DISTINCT podrodzaje_programowe.skrot "
        sql = sql & " FROM rodzaje_programowe RIGHT JOIN podrodzaje_programowe"
        sql = sql & " ON rodzaje_programowe.id = podrodzaje_programowe.id_rodzaju"
        sql = sql & " WHERE  rodzaje_programowe.rodzaj ='" & rodzaj & "' AND "
        sql = sql & " podrodzaje_programowe.podrodzaj ='" & podrodzaj & "'"

        Try
            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    tmp_skrot = dr.GetValue(0)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli podrodzaju programowego: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        skrot = tmp_skrot
        Return wynik

    End Function

    Public Function zapisz_kod_zadania(ByRef opis_wniosku As clsWniosekHonoracyjny, ByVal nowy_kod_zadania As Integer) As Integer
        Dim sql As String
        Dim rok As String
        Dim a As Integer

        rok = Year(opis_wniosku.data_emisji)


        sql = "UPDATE wnioski"
        sql = sql & " set zadanie=" & nowy_kod_zadania
        sql = sql & " WHERE id=" & opis_wniosku.id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function

    Public Function zapisz_nowy_kod_mpk(ByRef opis_wniosku As clsWniosekHonoracyjny, ByVal nowy_kod_mpk As String) As Integer
        Dim sql As String
        Dim rok As String
        Dim a As Integer

        rok = Year(opis_wniosku.data_emisji)


        sql = "UPDATE wnioski"
        sql = sql & " set mpk='" & nowy_kod_mpk & "'"
        sql = sql & " WHERE id=" & opis_wniosku.id

        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function


    Public Function zapisz_nowa_grupe_docelowa_wniosku(ByRef opis_wniosku As clsWniosekHonoracyjny, ByVal grupa_docelowa As String) As Integer
        Dim sql As String
        Dim a As Integer
        Dim grp_doc As String = ""

        grp_doc = Trim(skoryguj_apostrofy_do_SQL(grupa_docelowa))


        sql = "UPDATE wnioski"
        sql = sql & " set grupa_docelowa= '" & grp_doc & "'"
        sql = sql & " WHERE id=" & opis_wniosku.id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function


    Public Function zapisz_nowy_rodzaj_realizacji_wniosku(ByRef opis_wniosku As clsWniosekHonoracyjny, ByVal rr As String) As Integer
        Dim sql As String
        Dim a As Integer
        Dim grp_doc As String = ""
        Dim r As clsRodzajrealizacji
        Dim id_rr As Integer = 1

        For Each r In colRodzajeRealizacji
            If r.rodzaj = rr Then
                id_rr = r.id
                Exit For
            End If

        Next


        sql = "UPDATE wnioski"
        sql = sql & " set rodzaj_realizacji= " & id_rr
        sql = sql & " WHERE id=" & opis_wniosku.id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function



    Public Function zapisz_nowy_rodzaj_dokumentacji(ByRef opis_wniosku As clsWniosekHonoracyjny, ByVal dok As String) As Integer
        Dim sql As String
        Dim a As Integer



        sql = "UPDATE wnioski"
        sql = sql & " set rodzaj_dokumentacji= '" & dok & "'"
        sql = sql & " WHERE id=" & opis_wniosku.id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function

    Public Function zapisz_nowy_rodzaj_muzyki(ByRef opis_wniosku As clsWniosekHonoracyjny, ByVal rodzaj As String) As Integer
        Dim sql As String
        Dim a As Integer



        sql = "UPDATE wnioski"
        sql = sql & " set rodzaj_muzyki= '" & rodzaj & "'"
        sql = sql & " WHERE id=" & opis_wniosku.id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function
    Public Function zapisz_nowe_zrodlo_finansowania(ByRef opis_wniosku As clsWniosekHonoracyjny, ByVal id_zrodla As Integer) As Integer

        Dim sql As String
        Dim rok As String
        Dim a As Integer

        rok = Year(opis_wniosku.data_emisji)


        sql = "UPDATE wnioski"
        sql = sql & " set zrodlo_finansowania=" & id_zrodla
        sql = sql & " WHERE id=" & opis_wniosku.id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function

    Public Function zatwierdz_ostatecznie_wniosek(ByVal dp As Date, ByVal dk As Date, ByVal wartosc As Boolean) As Integer
        ' ta funkcja zatwoerdza wnioski w zadanym okresie
        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim d_em As String
        Dim miesiac As Integer

        Dim l_dni_miesiaca As Integer
        Dim d_pocz As String = ""
        Dim d_konc As String = ""
        Dim msg As String
        Dim aa As Integer

        d_pocz = Format(dp, "yyyy-MM-dd")
        d_konc = Format(dk, "yyyy-MM-dd")



        If wartosc = True Then 'zatwierdzanie
            a = sprawdz_zatwierdzenie_wnioskow2(1, dp, dk) ' zatwierdzenie na poziomie szefa
            If a > 0 Then
                msg = "W wybranym okresie są wnioski nie zatwierdzone na poziomie Szefa Programu"
                msg = msg & vbCrLf & "Czy przerwać procedurę zatwierdzania ?"
                aa = MessageBox.Show(msg, naglowek_komunikatow, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If aa = vbYes Then
                    Return -1
                End If
            ElseIf a < 0 Then
                Return -1
            End If
            a = sprawdz_zatwierdzenie_wnioskow2(2, dp, dk) ' zatwierdzenie na poziomie Kierownika
            If a > 0 Then
                msg = "W wybranym okresie są wnioski nie zatwierdzone na poziomie Kierownika Redakcji"
                msg = msg & vbCrLf & "Czy przerwać procedurę zatwierdzania ?"
                aa = MessageBox.Show(msg, naglowek_komunikatow, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If aa = vbYes Then
                    Return -1
                End If
            ElseIf a < 0 Then
                Return -1
            End If
        End If





        sql = "SET DATEFORMAT YMD "
        sql = sql & " UPDATE wnioski"
        sql = sql & " set zaakceptowany_przez_zarzad = '" & wartosc & "', "
        sql = sql & " imie_nazwisko_zarzad =  '" & aktualny_uzytkownik.imie_nazwisko & " '"

        sql = sql & " WHERE data_emisji BETWEEN '" & d_pocz & "' AND '" & d_konc & "' AND otwarty='FALSE'"

        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function
    Public Function zatwierdz_ostatecznie_wniosek(ByVal tryb As Integer, _
                                            ByVal data_emisji As Date, _
                                            ByVal id As Integer, _
                                            ByVal wartosc As Boolean) As Integer
        ' tryb 1 - aktualny wniosek (id podany jako parametr
        '       2 - cały dzień - data podana jako parametr
        '       3 - cały miesiąc - data miesiąca podana jako parametr
        '       id - id wniosku 
        ' wartość   true - zatwierdzanie
        '           false - zdejmowanie zatwierdzenia 

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim d_em As String
        Dim miesiac As Integer

        Dim l_dni_miesiaca As Integer
        Dim d_pocz As String = ""
        Dim d_konc As String = ""
        Dim dp As Date
        Dim dk As Date
        Dim msg As String
        Dim aa As Integer


        rok = Year(data_emisji)

        d_em = Format(data_emisji, "yyyy-MM-dd")

        If tryb = 3 Then
            miesiac = Month(data_emisji)
            l_dni_miesiaca = ustal_liczbe_dni_w_miesiacu(miesiac, Year(data_emisji))
            d_pocz = rok & "-" & miesiac & "-01"
            d_konc = rok & "-" & miesiac & "-" & l_dni_miesiaca
            dp = CDate(d_pocz)
            dk = CDate(d_konc)
        ElseIf tryb = 2 Then
            dp = data_emisji
            dk = data_emisji
        End If

        If wartosc = True Then 'zatwierdzanie
            If tryb > 1 Then
                a = sprawdz_zatwierdzenie_wnioskow2(1, dp, dk) ' zatwierdzenie na poziomie szefa
                If a > 0 Then
                    msg = "W wybranym okresie są wnioski nie zatwierdzone na poziomie Szefa Programu"
                    msg = msg & vbCrLf & "Czy przerwać procedurę zatwierdzania ?"
                    aa = MessageBox.Show(msg, naglowek_komunikatow, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If aa = vbYes Then
                        Return -1
                    End If
                ElseIf a < 0 Then
                    Return -1
                End If
                a = sprawdz_zatwierdzenie_wnioskow2(2, dp, dk) ' zatwierdzenie na poziomie Kierownika
                If a > 0 Then
                    msg = "W wybranym okresie są wnioski nie zatwierdzone na poziomie Kierownika Redakcji"
                    msg = msg & vbCrLf & "Czy przerwać procedurę zatwierdzania ?"
                    aa = MessageBox.Show(msg, naglowek_komunikatow, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If aa = vbYes Then
                        Return -1
                    End If
                ElseIf a < 0 Then
                    Return -1
                End If
            End If
        End If



        sql = "SET DATEFORMAT YMD "
        sql = sql & " UPDATE wnioski"
        sql = sql & " set zaakceptowany_przez_zarzad = '" & wartosc & "', "
        sql = sql & " imie_nazwisko_zarzad =  '" & aktualny_uzytkownik.imie_nazwisko & " '"

        If tryb = 1 Then
            sql = sql & " WHERE id=" & id
        ElseIf tryb = 2 Then
            sql = sql & " WHERE data_emisji='" & d_em & "' AND otwarty = 'FALSE'"
        ElseIf tryb = 3 Then
            sql = sql & " WHERE data_emisji BETWEEN '" & d_pocz & "' AND '" & d_konc & "' AND otwarty='FALSE'"
        End If

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function

    Public Function ustaw_blokade_wnioskow_w_wybr_okresie(ByVal dp As Date, ByVal dk As Date) As Integer

        Dim sql As String
        Dim a As Integer
        Dim d_pocz As String = ""
        Dim d_konc As String = ""


        d_pocz = Format(dp, "yyyy-MM-dd")
        d_konc = Format(dk, "yyyy-MM-dd")

        sql = "SET DATEFORMAT YMD "
        sql = sql & " UPDATE wnioski"
        sql = sql & " set blokada_edycji = 1"
        sql = sql & " WHERE data_emisji BETWEEN '" & d_pocz & "' AND '" & d_konc & "' AND otwarty='FALSE'"

        a = wykonaj_polecenie_SQL(sql)

        Return a



    End Function
    Public Function ustaw_blokade_wnioskow(ByVal dp As Date, ByVal dk As Date, ByVal wartosc As Boolean, ByVal wybrany_program As String) As Integer
        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim d_em As String
        Dim miesiac As Integer

        Dim l_dni_miesiaca As Integer
        Dim d_pocz As String = ""
        Dim d_konc As String = ""
        Dim p As clsProgram
        Dim id_programu As Integer = 0


        If Len(wybrany_program) > 0 Then
            For Each p In colProgramy
                If p.nazwa_programu = wybrany_program Then
                    id_programu = p.id
                    Exit For
                End If
            Next
        End If



        d_pocz = Format(dp, "yyyy-MM-dd")
        d_konc = Format(dk, "yyyy-MM-dd")

        sql = "SET DATEFORMAT YMD "
        sql = sql & " UPDATE wnioski"
        sql = sql & " set blokada_edycji = '" & wartosc & "'"
        sql = sql & " WHERE data_emisji BETWEEN '" & d_pocz & "' AND '" & d_konc & "' AND otwarty='FALSE'"
        If id_programu > 0 Then
            sql = sql & " AND nr_programu=" & id_programu
        End If
        a = wykonaj_polecenie_SQL(sql)

        Return a





    End Function

    Public Function ustaw_blokade_wnioskow(ByVal tryb As Integer, _
                                            ByVal data_emisji As Date, _
                                            ByVal id As Integer, _
                                            ByVal wartosc As Boolean, _
                                            ByVal wybrany_program As String) As Integer
        ' tryb 1 - aktualny wniosek (id podany jako parametr
        '       2 - cały dzień - data podana jako parametr
        '       3 - cały miesiąc - data miesiąca podana jako parametr
        '       id - id wniosku 
        ' wartość   true - blokowanie
        '           false - zdejmowanie blokady 
        'wybrany_program - nazwa rogramu - jeżeli pusty ciąg znaków to wszystkie programy

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim d_em As String
        Dim miesiac As Integer

        Dim l_dni_miesiaca As Integer
        Dim d_pocz As String = ""
        Dim d_konc As String = ""
        Dim p As clsProgram
        Dim id_programu As Integer = 0


        If Len(wybrany_program) > 0 Then
            For Each p In colProgramy
                If p.nazwa_programu = wybrany_program Then
                    id_programu = p.id
                    Exit For
                End If
            Next
        End If

        rok = Year(data_emisji)

        d_em = Format(data_emisji, "yyyy-MM-dd")

        If tryb = 3 Then
            miesiac = Month(data_emisji)
            l_dni_miesiaca = ustal_liczbe_dni_w_miesiacu(miesiac, Year(data_emisji))
            d_pocz = rok & "-" & miesiac & "-01"
            d_konc = rok & "-" & miesiac & "-" & l_dni_miesiaca
        End If

        sql = "SET DATEFORMAT YMD "
        sql = sql & " UPDATE wnioski"
        sql = sql & " set blokada_edycji = '" & wartosc & "'"
        If tryb = 1 Then
            sql = sql & " WHERE id=" & id
        ElseIf tryb = 2 Then
            sql = sql & " WHERE data_emisji='" & d_em & "' AND otwarty = 'FALSE'"
        ElseIf tryb = 3 Then
            sql = sql & " WHERE data_emisji BETWEEN '" & d_pocz & "' AND '" & d_konc & "' AND otwarty='FALSE'"
        End If
        If tryb > 1 Then
            If id_programu > 0 Then
                sql = sql & " AND nr_programu=" & id_programu
            End If
        End If

        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function

    Public Function sprawdz_skutecznosc_zalozenia_blokady(ByVal dp As Date, ByVal dk As Date, ByVal wybrany_program As String) As Integer
        'funkcja sprawdza czy w podanym zakresie dat
        ' są wnioski z nie założoną blokadą edycji
        ' jeżeli nie ma to zwraca 0
        ' jeżeli są to zwraca 1
        ' jeżeli błąd to zwraca -1

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim d_em As String
        Dim odczytano As Boolean = False
        Dim msg As String
        Dim wynik As Integer = 0
        Dim d_pocz As String = ""
        Dim d_konc As String = ""
        Dim miesiac As Integer
        Dim l_dni_miesiaca As Integer



        Dim p As clsProgram
        Dim id_programu As Integer = 0


        If Len(wybrany_program) > 0 Then
            For Each p In colProgramy
                If p.nazwa_programu = wybrany_program Then
                    id_programu = p.id
                    Exit For
                End If
            Next
        End If



        d_pocz = Format(dp, "yyyy-MM-dd")
        d_konc = Format(dk, "yyyy-MM-dd")

        sql = "SET DATEFORMAT YMD "
        sql = sql & " select blokada_edycji from  wnioski "
        sql = sql & " WHERE data_emisji BETWEEN '" & d_pocz & "' AND '" & d_konc & "' AND blokada_edycji='FALSE'"
        If id_programu > 0 Then
            sql = sql & " AND nr_programu=" & id_programu
        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                wynik = 1
            End If


        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli wykonania polecenia założenia blokady edycji wniosków:" & vbCrLf & ex.Message
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


    Public Function zaladuj_dostepne_rodzaje_programowe(ByRef colLista As Collection, ByVal id_audycji As Integer) As Integer
        'funkcja wczytuje z tabeli audycji liste rodzajów programowych dostępnych w danej audycji

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim id_aud As String = ""
        Dim k As Integer = 0
        Dim tmp_lista As String = ""
        Dim poz() As String



        If colLista.Count > 0 Then
            Do
                colLista.Remove(1)
            Loop While colLista.Count > 0
        End If


        sql = "select dostepne_rodzaje_programowe  from audycje where id=" & id_audycji


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True

                tmp_lista = dr.GetValue(0)
            Else
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu listy dostepnych rodzajów programowych w audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If wynik = 0 Then
            If Len(tmp_lista) > 0 Then
                poz = Split(tmp_lista, ";")

                For k = 0 To poz.Length - 1
                    colLista.Add(poz(k), "_" & k)
                Next

            End If
        End If


        Return wynik


    End Function


    Public Function zaladuj_liste_identyfikatoriw_audycji(ByRef colLista As Collection, ByVal id_ramowki As Integer, ByVal id_redakcji As Integer) As Integer
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim id_aud As String = ""
        Dim k As Integer = 0





        If colLista.Count > 0 Then
            Do
                colLista.Remove(1)
            Loop While colLista.Count > 0
        End If


        sql = "select identyfikator_audycji from audycje where id_ramowki=" & id_ramowki
        sql = sql & " and id_redakcji =" & id_redakcji
        sql = sql & " order by identyfikator_audycji"



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True

                    id_aud = dr.GetValue(0)
                    k += 1
                    colLista.Add(id_aud, k)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania listy używanych identyfikatorów audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik




    End Function

    Public Function sprawdz_skutecznosc_zalozenia_blokady(ByVal tryb As Integer, ByVal data_emisji As Date, ByVal wybrany_program As String) As Integer

        ' tryb  2 - cały dzień - data podana jako parametr
        '       3 - cały miesiąc - data miesiąca podana jako parametr

        'funkcja sprawdza czy w podanym zakresie dat lub w podanej dacie emisji
        ' są wnioski z nie założoną blokadą edycji
        ' jeżeli nie ma to zwraca 0
        ' jeżeli są to zwraca 1
        ' jeżeli błąd to zwraca -1


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim d_em As String
        Dim odczytano As Boolean = False
        Dim msg As String
        Dim wynik As Integer = 0
        Dim d_pocz As String = ""
        Dim d_konc As String = ""
        Dim miesiac As Integer
        Dim l_dni_miesiaca As Integer
        Dim p As clsProgram
        Dim id_programu As Integer = 0


        If Len(wybrany_program) > 0 Then
            For Each p In colProgramy
                If p.nazwa_programu = wybrany_program Then
                    id_programu = p.id
                    Exit For
                End If
            Next
        End If

        rok = Year(data_emisji)
        d_em = Format(data_emisji, "yyyy-MM-dd")

        If tryb = 3 Then
            miesiac = Month(data_emisji)
            l_dni_miesiaca = ustal_liczbe_dni_w_miesiacu(miesiac, Year(data_emisji))
            d_pocz = rok & "-" & miesiac & "-01"
            d_konc = rok & "-" & miesiac & "-" & l_dni_miesiaca
        End If


        sql = "SET DATEFORMAT YMD "
        sql = sql & " select blokada_edycji from  wnioski "
        If tryb = 2 Then
            sql = sql & " WHERE data_emisji='" & d_em & "' AND blokada_edycji='FALSE'"
        ElseIf tryb = 3 Then
            sql = sql & " WHERE data_emisji BETWEEN '" & d_pocz & "' AND '" & d_konc & "' AND blokada_edycji='FALSE'"
        End If

        If id_programu > 0 Then
            sql = sql & " AND nr_programu=" & id_programu
        End If

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                wynik = 1
            End If


        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli wykonania polecenia założenia blokady edycji wniosków:" & vbCrLf & ex.Message
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


    Public Function zapisz_informacje_o_akceptacji(ByVal tryb As Integer, _
                                                    ByVal wartosc As Boolean, _
                                                    ByVal imie_nazwisko As String, _
                                                    ByRef wn As clsWniosekHonoracyjny) As Integer

        'tryb:
        '   1 pole wystawiony
        '   2 akceptacja koerownika
        '   3 akceptacja szefa
        '   4 akceptacja zarzadu
        '   5 przygotowany do emisji
        '   6 zatwierdzony do emisji przez Inspektora programu
        '   7 zatwierdzony do emisji przez Kierownika redakcji
        '   8 zatwierdzony do emisji przez Szefa programu

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        rok = Year(wn.data_emisji)

        Dim nazwisko As String
        nazwisko = skoryguj_apostrofy_do_SQL(imie_nazwisko)

        sql = "UPDATE wnioski "
        If tryb = 1 Then
            sql = sql & " set wystawiony= '" & wartosc & "', "
            If wartosc = True Then
                sql = sql & " imie_nazwisko_osoby_wystawiajacej= '" & nazwisko & "'"
            Else
                sql = sql & " imie_nazwisko_osoby_wystawiajacej= ' '"
            End If
        ElseIf tryb = 2 Then
            sql = sql & " set zaakceptowany_przez_kierownika= '" & wartosc & "', "
            If wartosc = True Then
                sql = sql & " imie_nazwisko_kierownika= '" & nazwisko & "'"
            Else
                sql = sql & " imie_nazwisko_kierownika= ' '"
            End If
        ElseIf tryb = 3 Then
            sql = sql & " set zaakceptowany_przez_szefa_programu= '" & wartosc & "', "
            If wartosc = True Then
                sql = sql & " imie_nazwisko_szefa_programu= '" & nazwisko & "'"
            Else
                sql = sql & " imie_nazwisko_szefa_programu= ' '"
            End If
        ElseIf tryb = 4 Then
            sql = sql & " set zaakceptowany_przez_zarzad= '" & wartosc & "', "
            If wartosc = True Then
                sql = sql & " imie_nazwisko_zarzad= '" & nazwisko & "'"
            Else
                sql = sql & " imie_nazwisko_zarzad= ' '"
            End If
        ElseIf tryb = 5 Then
            sql = sql & " set przygotowany_do_emisji= '" & wartosc & "', "
            If wartosc = True Then
                sql = sql & " imie_nazwisko_autora2= '" & nazwisko & "'"
            Else
                sql = sql & " imie_nazwisko_autora2= ' '"
            End If
        ElseIf tryb = 6 Then
            sql = sql & " set akceptacja_wstepna_inspektora= '" & wartosc & "', "
            If wartosc = True Then
                sql = sql & " imie_nazwisko_inspektora= '" & nazwisko & "'"
            Else
                sql = sql & " imie_nazwisko_inspektora= ' '"
            End If
        ElseIf tryb = 7 Then
            sql = sql & " set akceptacja_wstepna_kierownika= '" & wartosc & "', "
            If wartosc = True Then
                sql = sql & " imie_nazwisko_kierownika2= '" & nazwisko & "'"
            Else
                sql = sql & " imie_nazwisko_kierownika2= ' '"
            End If
        ElseIf tryb = 8 Then
            sql = sql & " set akceptacja_wstepna_szefa= '" & wartosc & "', "
            If wartosc = True Then
                sql = sql & " imie_nazwisko_szefa_programu2= '" & nazwisko & "'"
            Else
                sql = sql & " imie_nazwisko_szefa_programu2= ' '"
            End If

        End If

        sql = sql & " WHERE id=" & wn.id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function
    Public Function zapisz_log_pozycji_wniosku(ByVal id_pozycji As Integer, ByVal tresc As String, ByVal komentarz As String) As Integer
        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim tresc1 As String
        Dim nazw As String
        Dim kom As String

        tresc1 = skoryguj_apostrofy_do_SQL(tresc)
        If Len(tresc1) > 299 Then
            tresc1 = Left(tresc1, 299)
        End If


        nazw = skoryguj_apostrofy_do_SQL(aktualny_uzytkownik.imie_nazwisko)
        kom = skoryguj_apostrofy_do_SQL(komentarz)


        sql = "insert into pozycje_wniosku_log "
        sql = sql & " (id_pozycji,"
        sql = sql & " imie_nazwisko_pracownika,"
        sql = sql & " nazwa_komputera, "
        sql = sql & " info"
        If rejestracja_kometarzy_pozycji_wniosku Then
            sql = sql & ", komentarz "
        End If
        sql = sql & " ) "
        sql = sql & " VALUES ("
        sql = sql & id_pozycji & ", "
        sql = sql & " '" & nazw & "', "
        sql = sql & " '" & nazwa_stacji_komputerowej & "', "
        sql = sql & " '" & tresc1 & "'"
        If rejestracja_kometarzy_pozycji_wniosku Then
            sql = sql & ", '" & kom & "' "
        End If

        sql = sql & ")"

        a = wykonaj_polecenie_SQL(sql)

        Return a



    End Function


    Public Function zapisz_log_wniosku_2(ByVal id As Integer, ByVal tresc As String) As Integer

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim tresc1 As String
        Dim nazw As String

        tresc1 = skoryguj_apostrofy_do_SQL(tresc)
        If Len(tresc1) > 299 Then
            tresc1 = Left(tresc1, 299)
        End If


        nazw = skoryguj_apostrofy_do_SQL(aktualny_uzytkownik.imie_nazwisko)



        sql = "insert into wnioski_log "
        sql = sql & " (id_wniosku,"
        sql = sql & " imie_nazwisko_pracownika,"
        sql = sql & " nazwa_komputera, "
        sql = sql & " info) "
        sql = sql & " VALUES ("
        sql = sql & id & ", "
        sql = sql & " '" & nazw & "', "
        sql = sql & " '" & nazwa_stacji_komputerowej & "', "
        sql = sql & " '" & tresc1 & "')"

        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function


    Public Function zapisz_log_wniosku(ByVal id_wniosku As Integer, ByVal tresc As String) As Integer

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim tresc1 As String
        Dim nazw As String

        tresc1 = skoryguj_apostrofy_do_SQL(tresc)
        If Len(tresc1) > 299 Then
            tresc1 = Left(tresc1, 299)
        End If


        nazw = skoryguj_apostrofy_do_SQL(aktualny_uzytkownik.imie_nazwisko)

        '   rok = Year(opis.data_emisji)


        sql = "insert into wnioski_log "
        sql = sql & " (id_wniosku,"
        sql = sql & " imie_nazwisko_pracownika,"
        sql = sql & " nazwa_komputera, "
        sql = sql & " info) "
        sql = sql & " VALUES ("
        sql = sql & id_wniosku & ", "
        sql = sql & " '" & nazw & "', "
        sql = sql & " '" & nazwa_stacji_komputerowej & "', "
        sql = sql & " '" & tresc1 & "')"

        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function

    Public Function zaladuj_dlugosc_audycji_MKA(ByRef poz As clsPozycjaZestawieniaMinutowego, ByVal dp As Date, ByVal dk As Date) As Integer

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim sql As String
        Dim rok As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim id_red As String
        Dim id_aud As String

        id_red = Left(poz.id_audycji, 2)
        id_aud = Right(poz.id_audycji, 2)


        poz.dlugosc = 1

        d_p_str = Format(dp, "yyyy-MM-dd")
        d_k_str = Format(dk, "yyyy-MM-dd")


        sql = "SELECT sum(dlugosc) "
        sql = sql & " FROM wnioski "
        sql = sql & " WHERE id_redakcji='" & id_red & "' AND id_audycji ='" & id_aud & "' AND data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                poz.dlugosc = dr.GetValue(0)
                odczytano = True
            End If
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu łącznej długości audycji do zestawienia minutowych kosztów audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try



        Return wynik




    End Function

    Public Function zaladuj_tytul_audycji_MKA(ByRef k_a As clsPozycjaZestawieniaMinutowego, ByVal dp As Date, ByVal dk As Date) As Integer
        'funkcja używana podczas łądowania zestawienia minutowych kosztów audycji
        'klauzula DISTINKT powodowała że jeżeli z jakiegość powodu zmienionu tytuł audycji ramówkowej
        'to w liście audycji powielał by się wiersz zestawienia

        'funkcja wczytuje tytuł pierwszej audycji
        'ze spisu wniosków


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim sql As String
        Dim rok As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim id_red As String
        Dim id_aud As String

        id_red = Left(k_a.id_audycji, 2)
        id_aud = Right(k_a.id_audycji, 2)

        If id_aud = "99" Then
            k_a.tytul_audycji = "Audycje pozaramówkowe"
            Return 0
        End If

        rok = Year(dp)
        d_p_str = Format(dp, "yyyy-MM-dd")
        d_k_str = Format(dk, "yyyy-MM-dd")


        sql = "SELECT tytul_audycji "
        sql = sql & " FROM wnioski "
        sql = sql & " WHERE id_redakcji='" & id_red & "' AND id_audycji ='" & id_aud & "' AND data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                k_a.tytul_audycji = dr.GetValue(0)
                odczytano = True
            Else
                k_a.tytul_audycji = "Nie określona"
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu tytułu audycji do zestawienia minutowych kosztów audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik

    End Function



    Public Function zaladuj_tytul_audycji_ZKA(ByRef k_a As clsKosztyAudycji, ByVal dp As Date, ByVal dk As Date) As Integer
        'funkcja używana podczas łądowania zestawienia kosztó audycji
        'klauzula DISTINKT powodowała że jeżeli z jakiegość powodu zmienionu tytuł audycji ramówkowej
        'to w liście audycji powielał by się wiersz zestawienia

        'funkcja wczytuje tytuł pierwszej audycji
        'ze spisu wniosków


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim sql As String
        Dim rok As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim id_red As String
        Dim id_aud As String

        id_red = Left(k_a.id_audycji, 2)
        id_aud = Right(k_a.id_audycji, 2)

        If id_aud = "99" Then
            k_a.tytul_audycji = "Audycje pozaramówkowe"
            Return 0
        End If

        rok = Year(dp)
        d_p_str = Format(dp, "yyyy-MM-dd")
        d_k_str = Format(dk, "yyyy-MM-dd")


        sql = "SELECT tytul_audycji "
        sql = sql & " FROM wnioski "
        sql = sql & " WHERE id_redakcji='" & id_red & "' AND id_audycji ='" & id_aud & "' AND data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                k_a.tytul_audycji = dr.GetValue(0)
                odczytano = True
            Else
                k_a.tytul_audycji = "Nie określona"
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu tytułu audycji do zestawienia kosztó audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik

    End Function

    Public Function zaladuj_liczbe_audycji_ZKA(ByRef k_a As clsKosztyAudycji, ByVal dp As Date, ByVal dk As Date) As Integer
        'funkcja używana podczas łądowania zestawienia kosztó audycji
        'klauzula DISTINKT powodowała że jeżeli z jakiegość powodu zmienionu tytuł audycji ramówkowej
        'to w liście audycji powielał by się wiersz zestawienia

        'funkcja wczytuje tytuł pierwszej audycji
        'ze spisu wniosków


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim sql As String
        Dim rok As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim id_red As String
        Dim id_aud As String

        id_red = Left(k_a.id_audycji, 2)
        id_aud = Right(k_a.id_audycji, 2)

        If id_aud = "99" Then
            k_a.tytul_audycji = "Audycje pozaramówkowe"
            Return 0
        End If

        rok = Year(dp)
        d_p_str = Format(dp, "yyyy-MM-dd")
        d_k_str = Format(dk, "yyyy-MM-dd")



        sql = "SELECT count(*) "
        sql = sql & " FROM wnioski "
        sql = sql & " WHERE id_redakcji='" & id_red & "' AND id_audycji ='" & id_aud & "' AND data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                k_a.liczba_audycji = dr.GetValue(0)
                odczytano = True
            Else
                k_a.liczba_audycji = 1 ' żeby później nie dzieliło sięprzez 0 (ta liczba jest potrzebna do wyliczenia średniej
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu liczby audycji do zestawienia kosztów audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik

    End Function

    Public Function db_ustal_nazwe_zadania(ByVal id As Integer, ByRef nazwa As String) As Integer
        Dim sql As String
        Dim tmp_str As String
        Dim odczytano As Boolean = False

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim msg As String


        sql = "select nazwa_zadania from zadania where id=" & id
        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then

                tmp_str = dr.GetValue(0)
                odczytano = True
            Else
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu nazwy zadania: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        If odczytano Then
            nazwa = tmp_str
        Else
            nazwa = "Zadanie nr " & id
        End If

        Return 0


    End Function

    Public Function zaladuj_spis_zrodel_do_zestawienia(ByVal dp As Date, ByVal dk As Date) As Integer
        'funkcja wykorzystywana jest przy wyliczaniu zestawienia zrealizowanych wydatków w rozbiciu na żródła finansowania
        'z tabeli wniosków honoracyjnych wczytywany jest zestaw żródeł (bez powtórzeń)
        'za któe wystawiono wnioski w okresie podanym jako parametry d_p - d_k 
        'po to aby w następnym kroku wczytać zrealizowane wydatki na te źródła

        'funkcja zwraca 0 jueżeli wykonanie OK
        '               -1 jeżeli wystąpił problem


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim sql As String
        Dim rok As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0

        Dim k_zr As clsKosztyZrodel

        rok = Year(dp)
        d_p_str = Format(dp, "yyyy-MM-dd")
        d_k_str = Format(dk, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT zrodlo_finansowania "
        sql = sql & " FROM wnioski "
        sql = sql & " WHERE  data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then

                    k_zr = New clsKosztyZrodel
                    k_zr.id = dr.GetValue(0)

                    colZestawienieKosztowWgZrodlaFinansowania.Add(k_zr, k_zr.id)

                    odczytano = True
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu listy źródeł finansowania w wystawionych wnioskach honoracyjnych: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If colZestawienieKosztowWgZrodlaFinansowania.Count > 0 Then
            For Each k_zr In colZestawienieKosztowWgZrodlaFinansowania
                k_zr.nazwa_zrodla = ustal_nazwe_zrodla(k_zr.id)
            Next
        End If

        Return wynik


    End Function

    Public Function zaladuj_spis_zadan_do_zestawienia(ByVal dp As Date, ByVal dk As Date) As Integer

        'funkcja wykorzystywana jest przy wyliczaniu zestawienia zrealizowanych wydatków w rozbiciu na zadania
        'z tabeli wniosków honoracyjnych wczytywany jest zestaw zadań (bez powtórzeń)
        'za któe wystawiono wnioski w okresie podanym jako parametry d_p - d_k 
        'po to aby w następnym kroku wczytać zrealizowane wydatki na te zadania

        'funkcja zwraca 0 jueżeli wykonanie OK
        '               -1 jeżeli wystąpił problem


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim sql As String
        Dim rok As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0

        Dim k_zad As clsKosztyZadan

        rok = Year(dp)
        d_p_str = Format(dp, "yyyy-MM-dd")
        d_k_str = Format(dk, "yyyy-MM-dd")


        sql = "SET DATEFORMAT YMD "
        sql = sql & "   SELECT DISTINCT pozycje_wniosku.zadanie "
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
        sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
        sql = sql & " WHERE  wnioski.data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then

                    k_zad = New clsKosztyZadan
                    k_zad.id = dr.GetValue(0)

                    colZestawienieKOsztowZadan.Add(k_zad, k_zad.id)

                    odczytano = True
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu listy zrealizowanych zadań: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik

    End Function


    Public Function zaladuj_spis_audycji_do_zestawienia_minutowego(ByRef r As clsRedakcja, ByVal d_p As Date, ByVal d_k As Date) As Integer
        'funkcja wykorzystywana jest przy wyliczaniu zestawienia minutowych kosdztów audycji
        'z tabeli wniosków honoracyjnych wczytywany jest zestaw audycji (bez powtórzeń)
        'za któe wystawiono wnioski w okresie podanym jako parametry d_p - d_k 
        'po to aby w następnym kroku wczytać zrealizowane wydatki na te audycje

        'funkcja zwraca 0 jueżeli wykonanie OK
        '               -1 jeżeli wystąpił problem


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim id_audycji As String

        Dim msg As String
        Dim sql As String
        Dim rok As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim k_a As clsPozycjaZestawieniaMinutowego
        Dim ii As Integer = 0
        Dim id_red As String

        id_red = r.identyfikator_redakcji

        rok = Year(d_p)
        d_p_str = Format(d_p, "yyyy-MM-dd")
        d_k_str = Format(d_k, "yyyy-MM-dd")


        sql = "SELECT  DISTINCT  id_audycji "
        sql = sql & " FROM wnioski"
        sql = sql & " WHERE id_redakcji='" & id_red & "' AND data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"
        sql = sql & " ORDER BY id_audycji "




        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    id_audycji = dr.GetValue(0)
                    k_a = New clsPozycjaZestawieniaMinutowego
                    k_a.id_audycji = id_red & "-" & id_audycji
                    k_a.redakcja = r.nazwa_redakcji
                    colZestawienieMinutoweAudycji.Add(k_a, k_a.id_audycji)

                    odczytano = True
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu listy zapisanych audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik


    End Function


    Public Function zaladuj_spis_audycji_za_miesiac_dla_szefa_programu(ByRef pr As clsProgram, ByVal d_p As Date, ByVal d_k As Date) As Integer
        'funkcja wykorzystywana jest przy wyliczaniu zestawienia zrealizowanych wydatków redakcji w rozbiciu na audycje
        'z tabeli wniosków honoracyjnych wczytywany jest zestaw audycji (bez powtórzeń)
        'za któe wystawiono wnioski w okresie podanym jako parametry d_p - d_k 
        'po to aby w następnym kroku wczytać zrealizowane wydatki na te audycje

        'funkcja zwraca 0 jueżeli wykonanie OK
        '               -1 jeżeli wystąpił problem


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim id_audycji As String

        Dim msg As String
        Dim sql As String
        Dim rok As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim k_a As clsKosztyAudycji
        Dim ii As Integer = 0
        Dim id_red As String


        rok = Year(d_p)
        d_p_str = Format(d_p, "yyyy-MM-dd")
        d_k_str = Format(d_k, "yyyy-MM-dd")


        sql = "SELECT  DISTINCT  id_audycji, id_redakcji "
        sql = sql & " FROM wnioski"
        sql = sql & " WHERE nr_programu=" & pr.id & " AND data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"
        sql = sql & " ORDER BY id_audycji "




        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    id_audycji = dr.GetValue(0)
                    id_red = dr.GetValue(1)
                    k_a = New clsKosztyAudycji
                    k_a.id_audycji = id_red & "-" & id_audycji


                    k_a.nazwa_redakcji = ustal_nazwe_redakcji(id_red)

                    colZestawienieKosztowAudycji.Add(k_a, k_a.id_audycji)

                    odczytano = True
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu listy zapisanych audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try



        Return wynik


    End Function


    Public Function zaladuj_spis_audycji_za_miesiac(ByRef r As clsRedakcja, ByVal d_p As Date, ByVal d_k As Date) As Integer
        'funkcja wykorzystywana jest przy wyliczaniu zestawienia zrealizowanych wydatków redakcji w rozbiciu na audycje
        'z tabeli wniosków honoracyjnych wczytywany jest zestaw audycji (bez powtórzeń)
        'za któe wystawiono wnioski w okresie podanym jako parametry d_p - d_k 
        'po to aby w następnym kroku wczytać zrealizowane wydatki na te audycje

        'funkcja zwraca 0 jueżeli wykonanie OK
        '               -1 jeżeli wystąpił problem


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim id_audycji As String

        Dim msg As String
        Dim sql As String
        Dim rok As String
        Dim d_p_str As String
        Dim d_k_str As String
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim k_a As clsKosztyAudycji
        Dim ii As Integer = 0
        Dim id_red As String

        id_red = r.identyfikator_redakcji

        rok = Year(d_p)
        d_p_str = Format(d_p, "yyyy-MM-dd")
        d_k_str = Format(d_k, "yyyy-MM-dd")


        sql = "SELECT  DISTINCT  id_audycji "
        sql = sql & " FROM wnioski"
        sql = sql & " WHERE id_redakcji='" & id_red & "' AND data_emisji BETWEEN '" & d_p_str & "' AND '" & d_k_str & "'"
        sql = sql & " ORDER BY id_audycji "




        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    id_audycji = dr.GetValue(0)
                    k_a = New clsKosztyAudycji
                    k_a.id_audycji = id_red & "-" & id_audycji
                    k_a.nazwa_redakcji = r.nazwa_redakcji
                    colZestawienieKosztowAudycji.Add(k_a, k_a.id_audycji)

                    odczytano = True
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu listy zapisanych audycji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try



        Return wynik


    End Function


    Public Function wyznacz_kwote_laczna_wniosku(ByVal id As Integer) As Decimal


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim msg As String
        Dim dp_str As String
        Dim dk_str As String
        Dim ident As String
        Dim odczytano As Boolean = False
        Dim wynik As Decimal = 0

        Dim kw As Decimal = 0





        sql = "SELECT "
        sql = sql & " Sum(pozycje_wniosku.stawka_podstawowa * pozycje_wniosku.ilosc "
        sql = sql & " * pozycje_wniosku.wspolczynnik_wyceny) AS suma "
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku ON wnioski.id = pozycje_wniosku.id_wniosku "
        sql = sql & " WHERE wnioski.id = " & id




        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then

                kw = dr.GetValue(0)

                odczytano = True
            Else
                odczytano = False
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu sumy honorariów wniosku: " & vbCrLf & ex.Message
            If InStr(UCase(msg), "DBNULL") = 0 Then
                wskaznik_myszy(0)
                Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
                wynik = -1
            End If
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If wynik = 0 Then
            'nie było błędu podczas odczytu
            wynik = kw
        End If

        Return wynik




    End Function


    Public Function wyznacz_kwote_wg_kosztow_uzysku(ByRef plp As clsPozycjaListyPlac, ByVal grupa As Integer, ByVal koszt As Integer, ByVal dp As Date, ByVal dk As Date) As Decimal
        'grupa 1- pracownicy
        '       2 współpracownicy
        '       3 - producenci


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim msg As String
        Dim dp_str As String
        Dim dk_str As String
        Dim ident As String
        Dim odczytano As Boolean = False
        Dim wynik As Decimal = 0

        Dim kw As Decimal = 0

        If grupa > 0 Then
            If koszt = 0 Then
                ident = "%2_-_" & grupa.ToString ' koszty 0%
            ElseIf koszt = 20 Then
                ident = "%1_-_" & grupa.ToString 'koszty 20%
            Else
                ident = "%0_-_" & grupa.ToString 'koszty 50%
            End If
        Else
            If koszt = 0 Then
                ident = "%2_-__"  ' koszty 0%
            ElseIf koszt = 20 Then
                ident = "%1_-__"  'koszty 20%
            Else
                ident = "%0_-__" 'koszty 50%
            End If

        End If

        rok = Year(dp)
        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")



        sql = "SELECT "
        sql = sql & " Sum(pozycje_wniosku.stawka_podstawowa * pozycje_wniosku.ilosc "
        sql = sql & " * pozycje_wniosku.wspolczynnik_wyceny) AS suma "
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku ON wnioski.id = pozycje_wniosku.id_wniosku "
        sql = sql & " WHERE wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' "
        sql = sql & " AND pozycje_wniosku.id_pracownika=" & plp.id_pracownika
        ' If grupa > 0 Then
        sql = sql & " AND pozycje_wniosku.identyfikator LIKE '" & ident & "'"
        'End If



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then

                kw = dr.GetValue(0)

                odczytano = True
            Else
                odczytano = False
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu zsumowanych honorariów pracowników: " & vbCrLf & ex.Message
            If InStr(UCase(msg), "DBNULL") = 0 Then
                wskaznik_myszy(0)
                Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
                wynik = -1
            End If
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If wynik = 0 Then
            'nie było błędu podczas odczytu
            wynik = kw
        End If

        Return wynik








    End Function


    Public Function wyznacz_kwote_ze_srodkow_wlasnych(ByRef plp As clsPozycjaListyPlac, ByVal grupa As Integer, ByVal dp As Date, ByVal dk As Date) As Decimal
        'grupa 1- pracownicy
        '       2 współpracownicy
        '       3 - producenci
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim msg As String
        Dim dp_str As String
        Dim dk_str As String
        Dim ident As String
        Dim odczytano As Boolean = False
        Dim wynik As Decimal = 0

        Dim kw As Decimal = 0


        ident = "%" & grupa.ToString

        rok = Year(dp)
        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")



        sql = "SELECT "
        sql = sql & " Sum(pozycje_wniosku.stawka_podstawowa * pozycje_wniosku.ilosc "
        sql = sql & " * pozycje_wniosku.wspolczynnik_wyceny) AS suma "
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku ON wnioski.id = pozycje_wniosku.id_wniosku "
        sql = sql & " WHERE wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' "
        sql = sql & " AND wnioski.zrodlo_finansowania=1 "
        sql = sql & " AND pozycje_wniosku.id_pracownika=" & plp.id_pracownika
        If grupa > 0 Then
            sql = sql & " AND pozycje_wniosku.identyfikator LIKE '" & ident & "'"
        End If



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then

                kw = dr.GetValue(0)

                odczytano = True
            Else
                odczytano = False
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu zsumowanych honorariów pracowników: " & vbCrLf & ex.Message
            If InStr(UCase(msg), "DBNULL") = 0 Then
                wskaznik_myszy(0)
                Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
                wynik = -1
            End If
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If wynik = 0 Then
            'nie było błędu podczas odczytu
            wynik = kw
        End If

        Return wynik


    End Function

    Public Function db_zaladuj_liste_plac(ByVal grupa As Integer, ByVal dp As Date, ByVal dk As Date) As Integer
        'grupa 1- pracownicy
        '       2 współpracownicy
        '       3 - producenci
        'od wersji 2.0.87 jest mozliwośc załadowania wszystkich - grupa 0


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim msg As String
        Dim wynik As Integer = 0
        Dim dp_str As String
        Dim dk_str As String
        Dim ident As String
        Dim odczytano As Boolean = False
        Dim plp As clsPozycjaListyPlac


        ident = "%" & grupa.ToString

        rok = Year(dp)
        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")



        sql = "SELECT pozycje_wniosku.id_pracownika, "
        '        sql = sql & " pozycje_wniosku_" & rok & ".imie_nazwisko_pracownika, "
        sql = sql & " Sum(pozycje_wniosku.stawka_podstawowa * pozycje_wniosku.ilosc "
        sql = sql & " * pozycje_wniosku.wspolczynnik_wyceny) AS suma "
        sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku ON wnioski.id = pozycje_wniosku.id_wniosku "
        sql = sql & " WHERE wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' "
        If grupa > 0 Then
            ' to na życzenie Radia Wrocław - dostepne od wersji 2.0.87
            sql = sql & " AND pozycje_wniosku.identyfikator LIKE '" & ident & "'"
        End If

        sql = sql & " GROUP BY pozycje_wniosku.id_pracownika "
        '        sql = sql & " ORDER BY pozycje_wniosku_" & rok & ".imie_nazwisko_pracownika "



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    plp = New clsPozycjaListyPlac
                    plp.id_pracownika = dr.GetValue(0)
                    '                    plp.imie_nazwisko = dr.GetValue(1)
                    plp.kwota = dr.GetValue(1)

                    colListaPlac.Add(plp, plp.id_pracownika & "_")

                    odczytano = True
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu zsumowanych honorariów pracowników: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik

    End Function

    Public Function sprawdz_czy_wnioski_sa_edytowane(ByVal dp As Date, ByVal dk As Date, ByRef wynik_str As String) As Integer

        'funkcja sprqawdza czy wnioski w okresiepodanym jako parametry otwarte do edycji
        'zwraca 0 jeżeli nie ma wniosków otwartych do edycji
        '       1 jeżeli są wnioski właśnie teraz edytowane
        '       -1 jeżeli błąd

        'używana jest podczas eksportu wniosków 

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim msg As String
        Dim wynik As Integer = 0
        Dim dp_str As String
        Dim dk_str As String
        Dim data_em As Date
        Dim tytul As String = ""
        Dim tmp_str As String = ""


        rok = Year(dp)
        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")



        sql = "SET DATEFORMAT YMD "
        sql = sql & " select id, "
        sql = sql & "data_emisji,"
        sql = sql & "tytul_audycji "

        sql = sql & " from wnioski"
        sql = sql & " WHERE data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' "
        sql = sql & " AND otwarty=1"



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                data_em = dr.GetValue(1)
                tytul = dr.GetValue(2)

                wynik = 1
            End If


        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli stanu wniosków:" & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            Cmd.Dispose()
            dr.Close()

        Catch ex As Exception

        End Try
        If wynik > 0 Then
            tmp_str = tytul & " z dnia " & Format(data_em, "yyyy-MM-dd")
        End If

        wynik_str = tmp_str
        Return wynik

    End Function

    Public Function sprawdz_zatwierdzenie_wnioskow2(ByVal tryb As Integer, ByVal dp As Date, ByVal dk As Date) As Integer

        'funkcja sprqawdza czy wnioski w okresie podanym jako parametry są zatwierdzone:
        ' jeżeli tryb   1 - przez szefa programu
        '               2 - przez kierowników redakcji

        'zwraca 0 jeżeli wszystkiesą zatwierdzone
        '       1 jeżeli są wnioski niezatwoerdzone
        '       -1 jeżeli błąd

        'funkcja jest używana przed grupowym zatwierdzeniem wniosków przez zarząd

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim msg As String
        Dim wynik As Integer = 0
        Dim dp_str As String
        Dim dk_str As String

        rok = Year(dp)
        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")



        sql = "SET DATEFORMAT YMD "
        sql = sql & " select id from wnioski"
        sql = sql & " WHERE data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' AND "
        If tryb = 1 Then
            sql = sql & " zaakceptowany_przez_szefa_programu='FALSE'"
        Else
            sql = sql & " zaakceptowany_przez_kierownika='FALSE'"

        End If



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                wynik = 1
            End If


        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli stanu zatwierdzenia wnioskó:" & vbCrLf & ex.Message
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

    Public Function sprawdz_zatwierdzenie_wnioskow(ByVal dp As Date, ByVal dk As Date) As Integer

        'funkcja sprqawdza czy wnioski w okresiepodanym jako parametry są zatwierdzone przez zarząd
        'zwraca 0 jeżeli wszystkiesą zatwierdzone
        '       1 jeżeli są wnioski niezatwoerdzone
        '       -1 jeżeli błąd


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim rok As String
        Dim a As Integer
        Dim msg As String
        Dim wynik As Integer = 0
        Dim dp_str As String
        Dim dk_str As String

        rok = Year(dp)
        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")



        sql = "SET DATEFORMAT YMD "
        sql = sql & " select id from wnioski"
        sql = sql & " WHERE data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' AND zaakceptowany_przez_zarzad='FALSE'"



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then
                wynik = 1
            End If


        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli stanu zatwierdzenia wnioskó:" & vbCrLf & ex.Message
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
    Public Function ustal_nazwisko_pracownika(ByVal id As Integer, ByRef imie_nazwisko As String, ByRef nazwisko As String) As Integer
        'funkcja używana podczas generowania listy płac
        'wczytuke imie nazwisko ze spisu pracownikó
        ' potrzeba ta wynika z tego że wczytanie listy płac może odbywać się z wczytaniem tylko identyfikatora
        'ze względu na sposób działąnia funkcji agregujących

        'nazwisko dodano w dniu 10 sierpnia 2009
        'ze względu na sortowanie listy 

        ' zwraca 0 jeżeli OK i wtedy podstawia nazwisko
        '       - 1 jeżeli błąd

        Dim sql As String
        Dim odczytano As Boolean = False
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim msg As String
        Dim nazw As String = ""
        Dim im_nazw As String

        sql = "select nazwisko, imie_nazwisko from pracownicy where id=" & id

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        wskaznik_myszy(0)
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            If dr.Read Then

                nazw = dr.GetValue(0)
                im_nazw = dr.GetValue(1)

                odczytano = True
            Else
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu danych pracownika: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        If odczytano Then
            nazwisko = nazw
            imie_nazwisko = im_nazw
            wynik = 0
        Else
            nazwisko = "Nie ustalono !!!"
            imie_nazwisko = "Nie ustalono !!!"
            wynik = -1
        End If

        Return wynik

    End Function
    Public Function sprawdz_identyfikator_zewnetrzny(ByRef opis_osoby As clsPracownik, ByRef nazwiska_osob As String) As Integer
        'funkcja sprawdza czy w bazie danych jest pracownik o id_zewnetrznym podanym jako parametr
        'jeżeli jest to przez referencję zwraca nazwisko osoby (osób)
        'zwtraca 0 jeżeli nie ma innych osób,
        '- 1 jeżeli błąd 
        ' liczbę osóób które mają identyfikator zewnętrzny podany jako parametr
        Dim sql As String
        Dim k As Integer = 0
        Dim odczytano As Boolean = False
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = -1
        Dim msg As String
        Dim tmp_id_rekordu As Integer
        Dim tmp_nazw As String
        Dim wynik_nazwiska As String = ""


        If opis_osoby.identyfikator_zewnetrzny = 0 Then
            Return 0
        End If

        sql = "select id, imie_nazwisko from pracownicy "
        sql = sql & "where identyfikator_zewnetrzny=" & opis_osoby.identyfikator_zewnetrzny


        Try
            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    k += 1
                    tmp_id_rekordu = dr.GetValue(0)
                    tmp_nazw = dr.GetValue(1)
                    If tmp_id_rekordu <> opis_osoby.id Then
                        wynik_nazwiska = wynik_nazwiska & " " & tmp_nazw
                    Else
                        k -= 1 'jeżeli to jest ta sama osoba to odejmowanie licznika
                    End If
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = k

        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania kontroli pola identyfikatora zewnętrznego w tabeli spisu pracowników: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()
        Catch ex As Exception

        End Try

        If wynik > 0 Then
            nazwiska_osob = wynik_nazwiska
        Else
            nazwiska_osob = ""
        End If


        Return wynik





    End Function

    Public Function zapisz_nowy_budzet_redakcji(ByVal id_red As String, ByVal miesiac As Integer, ByVal rok As Integer) As Integer
        Dim sql As String
        Dim a As Integer

        sql = "INSERT INTO BUDZETY_REDAKCJI "
        sql = sql & "(id_redakcji, "
        sql = sql & "miesiac, "
        sql = sql & "rok) "
        sql = sql & " VALUES('" & id_red & "',"
        sql = sql & miesiac & ","
        sql = sql & rok & ")"

        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function

    Public Function zapisz_budzet_redakcji(ByVal id_rekordu As Integer, ByRef br As clsBudzetRedakcji) As Integer
        Dim sql As String
        Dim a As Integer
        Dim ap_str As String
        Dim rp_str As String
        Dim aw_str As String
        Dim rw_str As String
        Dim p_str As String

        ap_str = formatuj_wycene(br.ap)
        rp_str = formatuj_wycene(br.rp)
        aw_str = formatuj_wycene(br.aw)
        rw_str = formatuj_wycene(br.rw)
        p_str = formatuj_wycene(br.p)

        sql = "UPDATE BUDZETY_REDAKCJI "
        sql = sql & "set "
        sql = sql & "plan_ap= " & ap_str & ","
        sql = sql & "plan_rp= " & rp_str & ","
        sql = sql & "plan_aw= " & aw_str & ","
        sql = sql & "plan_rw= " & rw_str & ","
        sql = sql & "plan_p= " & p_str
        sql = sql & " where id = " & id_rekordu


        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function

    Public Function db_zaladuj_budzet_redakcji(ByVal id_red As String, _
                                            ByRef br As clsBudzetRedakcji, _
                                            ByVal miesiac As Integer, _
                                            ByVal rok As Integer) As Integer


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False


        br.ap = 0
        br.rp = 0
        br.aw = 0
        br.rw = 0
        br.p = 0

        sql = "Select plan_ap, "
        sql = sql & "plan_rp, "
        sql = sql & "plan_aw, "
        sql = sql & "plan_rw, "
        sql = sql & "plan_p  "
        sql = sql & "from budzety_redakcji"
        sql = sql & " where "
        sql = sql & " id_redakcji ='" & id_red & "'"
        sql = sql & " AND miesiac =" & miesiac
        sql = sql & " AND rok =" & rok
        wskaznik_myszy(1)
        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować sie z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                br.ap = dr.GetValue(0)
                br.rp = dr.GetValue(1)
                br.aw = dr.GetValue(2)
                br.rw = dr.GetValue(3)
                br.p = dr.GetValue(4)
            End If
            wynik = 0
        Catch ex As Exception
            wskaznik_myszy(0)
            msg = "Wystąpił problem podczas ładowania rekordu planu wydatków: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        wskaznik_myszy(0)

        Return wynik

    End Function

    Public Function zapisz_dlugosc_muzyki(ByRef opis As clsWniosekHonoracyjny)
        Dim sql As String
        Dim a As Integer


        sql = "update wnioski "
        sql = sql & " set dlugosc_muzyki = " & opis.dlugosc_muzyki
        sql = sql & " where id= " & opis.id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function
    Public Function zapisz_info_wniosku(ByVal id_wniosku As Integer, ByVal info As String) As Integer

        Dim sql As String
        Dim a As Integer

        Dim tmp_str As String

        tmp_str = skoryguj_apostrofy_do_SQL(info)
        If Len(tmp_str) > 4000 Then
            tmp_str = Left(tmp_str, 3999)
        End If

        sql = "update wnioski "
        sql = sql & " set info = '" & tmp_str & "' "
        sql = sql & " where id= " & id_wniosku

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function
    Public Function zapisz_nowy_tytul_wniosku(ByVal id_wniosku As Integer, ByVal tytul As String) As Integer
        Dim sql As String
        Dim a As Integer

        Dim tyt As String
        tyt = skoryguj_apostrofy_do_SQL(tytul)

        sql = "update wnioski "
        sql = sql & " set tytul_audycji = '" & tyt & "' "
        sql = sql & " where id= " & id_wniosku

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function

    Public Function zapisz_nowy_numer_programu(ByVal id_wniosku As Integer, ByVal nr_programu As Integer) As Integer
        Dim sql As String
        Dim a As Integer

        sql = "update wnioski "
        sql = sql & " set nr_programu = " & nr_programu
        sql = sql & " where id= " & id_wniosku

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function
    Public Function zapisz_nowego_autora_wniosku(ByVal id_wniosku As Integer, ByVal id_autora As Integer, ByVal nazw_autora As String) As Integer

        Dim sql As String
        Dim a As Integer
        Dim aut As String
        aut = skoryguj_apostrofy_do_SQL(nazw_autora)

        sql = "update wnioski "
        sql = sql & " set id_autora = " & id_autora & ", "
        sql = sql & " imie_nazwisko_autora = '" & aut & "' "
        sql = sql & " where id= " & id_wniosku

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function
    Public Function wyznacz_dlugosc_slowa2(ByVal tabela As Integer, ByVal d_em As Date, ByVal nr_programu As Integer) As Integer

        ' funkcja zwraca długość słowa w sekundach

        'tabela 0 - z tabeli pozycje_wniosku
        '       1 -z tabeli  pozycje sprawozdania programowego 

        'numer programu 0 - wszystkie programy
        '               1,2,3 - numer programu

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim d_em_str As String

        d_em_str = Format(d_em, "yyyy-MM-dd")

        If tabela = 0 Then
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.dlugosc) "
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku "
            sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku "
            sql = sql & " WHERE  wnioski.data_emisji = '" & d_em_str & "'"
            If nr_programu > 0 Then
                sql = sql & " AND wnioski.nr_programu=" & nr_programu
            End If
        Else
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_sprawozdania_programowego.dlugosc) "
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_sprawozdania_programowego "
            sql = sql & " ON wnioski.id = pozycje_sprawozdania_programowego.id_wniosku "
            sql = sql & " WHERE  wnioski.data_emisji ='" & d_em_str & "' "
            If nr_programu > 0 Then
                sql = sql & " AND wnioski.nr_programu=" & nr_programu
            End If

        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                If dr.GetValue(0) IsNot System.DBNull.Value Then
                    wynik = dr.GetValue(0)
                Else
                    wynik = 0
                End If
            Else
                odczytano = False
            End If
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu długości słowa z bazy danych: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik




    End Function

    Public Function wyznacz_dlugosc_slowa_wyfiltrowanego(ByVal tryb As Integer, _
                                            ByVal tabela As Integer, _
                                            ByVal dp As Date, _
                                            ByVal dk As Date, _
                                            ByVal kryt As String) As Integer

        ' funkcja zwraca długość słowa w sekundach

        'tryb 0 - całe słowo
        '     1 - słowo regionalne
        '
        'tabela 0 - z tabeli pozycje_wniosku
        '       1 -z tabeli  pozycje sprawozdania programowego 

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim dp_str As String
        Dim dk_str As String

        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")

        If tabela = 0 Then
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.dlugosc) "
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku "
            sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku "
            sql = sql & " WHERE  wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' "
            If tryb = 1 Then
                sql = sql & " AND pozycje_wniosku.region = 1 "
            End If
            sql = sql & " AND " & kryt
        Else
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_sprawozdania_programowego.dlugosc) "
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_sprawozdania_programowego "
            sql = sql & " ON wnioski.id = pozycje_sprawozdania_programowego.id_wniosku "
            sql = sql & " WHERE  wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "' "
            If tryb = 1 Then
                sql = sql & " AND pozycje_sprawozdania_programowego.region = 1 "
            End If
            sql = sql & " AND " & kryt
        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                If dr.GetValue(0) IsNot System.DBNull.Value Then
                    wynik = dr.GetValue(0)
                Else
                    wynik = 0
                End If
            Else
                odczytano = False
            End If
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu długości słowa do sprawozdania: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik

    End Function
    Public Function wyznacz_dlugosc_slowa(ByVal tryb As Integer, _
                                            ByVal tabela As Integer, _
                                            ByVal dp As Date, _
                                            ByVal dk As Date) As Integer

        ' funkcja zwraca długość słowa w sekundach

        'tryb 0 - całe słowo
        '     1 - słowo regionalne
        '
        'tabela 0 - z tabeli pozycje_wniosku
        '       1 -z tabeli  pozycje sprawozdania programowego 

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim dp_str As String
        Dim dk_str As String

        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")

        If tabela = 0 Then
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.dlugosc) "
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku "
            sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku "
            sql = sql & " WHERE  wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "'"
            If tryb = 1 Then
                sql = sql & " AND pozycje_wniosku.region = 1"
            End If
        Else
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_sprawozdania_programowego.dlugosc) "
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_sprawozdania_programowego "
            sql = sql & " ON wnioski.id = pozycje_sprawozdania_programowego.id_wniosku "
            sql = sql & " WHERE  wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "'"
            If tryb = 1 Then
                sql = sql & " AND pozycje_sprawozdania_programowego.region = 1"
            End If


        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True
                If dr.GetValue(0) IsNot System.DBNull.Value Then
                    wynik = dr.GetValue(0)
                Else
                    wynik = 0
                End If
            Else
                odczytano = False
            End If
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu długości słowa do sprawozdania " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try


        Return wynik



    End Function

    Public Function sprawdz_zapisane_pozycje(ByVal id_wniosku As Integer, ByVal tryb As Integer) As Integer

        'funkcja sprawdza czy wniosek ma zapisane jakiekolwiek pozycje
        'uruchamiana jest przed usunięciem wniosku

        'tryb   0 - w tabeli pozycji wniosku
        '       1 -w tabeli pozycji pozahonoracyjnych
        '       2  w tabeli pozycji sprawozdania programowego
        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim msg As String
        Dim wynik As Integer = 0


        If tryb = 0 Then
            sql = "select id from pozycje_wniosku "
        ElseIf tryb = 1 Then
            sql = "select id from niehonoracyjne_pozycje_wniosku "
        Else
            sql = "select id from pozycje_sprawozdania_programowego "

        End If


        sql = sql & " where id_wniosku = " & id_wniosku

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                'jest rekord
                wynik = 1
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli wystawionych wycen we wniosku:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd = Nothing

        Catch ex As Exception

        End Try


        Return wynik





    End Function


    Public Function zaladuj_pozycje_wniosku_do_wydruku(ByVal id_wniosku As Integer, ByVal ukryj_kwoty As Boolean) As Integer
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False
        Dim poz As clsPozycjaWniosku
        Dim aa As Integer
        Dim k As Integer = 0


        If colPozycjeWnioskuDoWYdruku.Count > 0 Then
            Do
                colPozycjeWnioskuDoWYdruku.Remove(1)
            Loop While colPozycjeWnioskuDoWYdruku.Count > 0
        End If



        sql = "select "
        sql = sql & "id,"
        sql = sql & " imie_nazwisko_pracownika,"
        sql = sql & " nazwa_pozycji,"
        sql = sql & " tytul,"
        sql = sql & " identyfikator,"
        sql = sql & " stawka_podstawowa,"
        sql = sql & " ilosc, "
        sql = sql & " wspolczynnik_wyceny, "
        sql = sql & " godzina_emisji, "
        sql = sql & " rodzaj, "
        sql = sql & " podrodzaj, "
        sql = sql & " dlugosc, "
        sql = sql & " region,"
        sql = sql & "komentarz "

        sql = sql & " from pozycje_wniosku where id_wniosku = " & id_wniosku

        If rozszerzone_sprawozdania_dostepne Then
            sql = sql & " union select "
            sql = sql & "id,"
            sql = sql & " imie_nazwisko_pracownika,"
            sql = sql & " '-'  as nazwa_pozycji,"
            sql = sql & " tytul,"
            sql = sql & " '-' as identyfikator,"
            sql = sql & " 0 as stawka_podstawowa,"
            sql = sql & " 1 as ilosc, "
            sql = sql & " 1 as wspolczynnik_wyceny, "
            sql = sql & " godzina_emisji, "
            sql = sql & " rodzaj, "
            sql = sql & " podrodzaj, "
            sql = sql & " dlugosc, "
            sql = sql & " region, "
            sql = sql & "komentarz "
            sql = sql & " from pozycje_sprawozdania_programowego where id_wniosku = " & id_wniosku

        End If

        sql = sql & " order by godzina_emisji"

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    poz = New clsPozycjaWniosku
                    poz.id = dr.GetValue(0)

                    poz.imie_nazwisko_pracownika = dr.GetValue(1)
                    poz.nazwa_pozycji = dr.GetValue(2)
                    poz.tytul = dr.GetValue(3)
                    poz.pelny_identyfikator = dr.GetValue(4)
                    If ukryj_kwoty Then
                        poz.stawka_podstawowa = 0
                    Else
                        poz.stawka_podstawowa = dr.GetValue(5)
                    End If
                    poz.ilosc = dr.GetValue(6)
                    poz.godzina_emisji = dr.GetValue(8)
                    poz.wspolczynnik_wyceny = dr.GetValue(7)
                    poz.rodzaj = dr.GetValue(9)
                    poz.podrodzaj = dr.GetValue(10)
                    poz.dlugosc = dr.GetValue(11)
                    poz.region = dr.GetValue(12)
                    Try
                        poz.komentarz = dr.GetValue(13)
                        If poz.komentarz = "-" Then
                            poz.komentarz = ""
                        End If
                    Catch ex13 As Exception

                    End Try
                    k += 1
                    colPozycjeWnioskuDoWYdruku.Add(poz, poz.id & "_" & k)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu  pozycji wniosku do wydruku: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try
        Return wynik



    End Function

    Public Function sprawdz_czy_redakcja_posiada_audycje(ByVal id As Integer) As Integer

        'funkcja używana jest podczas usuwania redakcji

        'sprawdza czy do redakcji przypisano jakiekolwiek audycje
        '1 - jeżeli są
        '-1 jeżeli błąd

        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim msg As String


        sql = sql & " SELECT tytul_audycji from audycje "
        sql = sql & " where id_redakcji=" & id



        Try

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                'jest rekord
                wynik = 1
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli audycji przypisanych do redakcji:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd = Nothing

        Catch ex As Exception

        End Try

        Return wynik

    End Function

    Public Function sprawdz_czy_wniosek_juz_zapisany(ByVal data_emisji As Date, ByVal ramowkowy_id_audycji As Integer) As Integer
        'funkcja używana jest podczas tworzenia nowego wniosku

        'sprawdza czy w danym dniu już wystawiono wniosek za audycje o podanym id ramówkowym

        'zwraca 0 jeżeli nie ma 
        '1 - jeżeli jest juz wniosek
        '-1 jeżeli błąd

        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim msg As String
        Dim d_em_str As String
        Dim dk_str As String

        d_em_str = Format(data_emisji, "yyyy-MM-dd")



        sql = "SET DATEFORMAT YMD "
        sql = sql & " SELECT "
        sql = sql & "id  "
        sql = sql & " FROM wnioski "
        sql = sql & " WHERE id_audycji_tabeli_wycen=" & ramowkowy_id_audycji
        sql = sql & " AND data_emisji='" & d_em_str & "'"



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                'jest rekord
                wynik = 1
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli już wystawionych wniosków:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd = Nothing

        Catch ex As Exception

        End Try


        Return wynik

    End Function

    Public Function sprawdz_nazwisko_nowego_pracownika(ByVal nazwisko As String, ByVal tryb As Integer) As Integer

        'funcka sprawdza czy w spisie pracownikó jest juz osoba o podanym imieniu i nazwisku
        'lub login namie - (tryb 0)

        'tryb 0 - login name
        '       'tryb 1 - imie nazwisko

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim msg As String
        Dim wynik As Integer = 0
        Dim sql As String = ""

        Dim tmp_nazw As String
        tmp_nazw = skoryguj_apostrofy_do_SQL(nazwisko)


        sql = "select id from pracownicy "
        If tryb = 0 Then
            sql = sql & " where nazwisko='" & nazwisko & "'"
        Else
            sql = sql & " where imie_nazwisko='" & tmp_nazw & "'"
        End If

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader
            If dr.Read() Then
                'jest rekord
                wynik = 1
            End If

        Catch ex As Exception
            msg = "Wystąpił problem podczas kontroli wystawionych wycen we wniosku:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd = Nothing

        Catch ex As Exception

        End Try

        Return wynik

    End Function

    Public Function zapisz_nowy_identyfikator_kosztow(ByVal id As Integer, ByVal identyfikator As String) As Integer


        Dim sql As String
        Dim a As Integer

        sql = "update pozycje_wniosku "
        sql = sql & " set identyfikator = '" & identyfikator & "' "
        sql = sql & " where id= " & id

        a = wykonaj_polecenie_SQL(sql)

        Return a


    End Function

    Public Function zapisz_stala_pozycje_wniosku(ByVal id_wniosku As Integer, _
                                                    ByRef p As clsStalaPozycjaAudycji) As Integer
        Dim sql As String
        Dim a As Integer
        Dim rok As String
        Dim spr_str As String = "-"
        Dim spr_str2 As String = "-"
        Dim min As Integer
        Dim sek As Integer
        Dim tytul As String
        ' Dim nazwa As String
        Dim imie_nazwisko As String
        Dim komentarz As String
        Dim godz_em As String
        Dim grupa_docelowa As String

        wskaznik_myszy(1)

        Dim r As clsRodzajProgramowy
        For Each r In colRodzajeProgramowe
            If r.rodzaj = p.rodzaj Then
                spr_str = r.skrot
                Exit For
            End If
        Next
        a = ustal_skrot_podrodzaju(p.rodzaj, p.podrodzaj, spr_str2)
        spr_str = spr_str & " / " & spr_str2
        min = Int(p.dlugosc / 60)
        sek = p.dlugosc - (min * 60)
        If sek > 9 Then
            spr_str = spr_str & " " & min & "'' " & sek & """ "
        Else
            spr_str = spr_str & " " & min & "'' 0" & sek & """ "
        End If

        If p.id_pracownika = 0 Then
            p.imie_nazwisko_pracownika = "-"
        End If
        imie_nazwisko = skoryguj_apostrofy_do_SQL(p.imie_nazwisko_pracownika)
        tytul = skoryguj_apostrofy_do_SQL(p.tytul)
        grupa_docelowa = Trim(skoryguj_apostrofy_do_SQL(p.grupa_docelowa))


        godz_em = Trim(p.godzina_emisji)
        If Len(godz_em) < 5 Then
            godz_em = "0" & godz_em
        End If

        If Len(p.rodzaj_produkcji) < 1 Then
            p.rodzaj_produkcji = "-"
        End If
        If Len(p.rodzaj_audycji) < 1 Then
            p.rodzaj_audycji = "-"
        End If

        sql = "INSERT INTO POZYCJE_SPRAWOZDANIA_PROGRAMOWEGO "
        sql = sql & " (id_wniosku, "
        sql = sql & "tytul, "
        sql = sql & "id_pracownika, "
        sql = sql & "imie_nazwisko_pracownika, "
        sql = sql & "rodzaj, "
        sql = sql & "podrodzaj, "
        sql = sql & "dlugosc, "
        sql = sql & "sprawozdanie, "
        sql = sql & "godzina_emisji, "
        sql = sql & "region, "
        sql = sql & "powtorka, "
        sql = sql & "rodzaj_licencji, "
        sql = sql & "rodzaj_audycji, "
        sql = sql & "rodzaj_produkcji, "
        sql = sql & "forma_radiowa "
        If MPK_dostepne And tryb_obslugi_MPK = 1 Then
            sql = sql & ", mpk "
        End If
        sql = sql & ", grupa_docelowa "
        sql = sql & ")"
        sql = sql & " VALUES(" & id_wniosku & ","
        sql = sql & "'" & tytul & "',"
        sql = sql & p.id_pracownika & ","
        sql = sql & "'" & imie_nazwisko & "',"
        sql = sql & "'" & p.rodzaj & "', "
        sql = sql & "'" & p.podrodzaj & "', "
        sql = sql & p.dlugosc & ", "
        sql = sql & "'" & spr_str & "', "
        sql = sql & "'" & godz_em & "', "
        sql = sql & "'" & p.region & "', "
        sql = sql & "'" & p.powtorka & "', "
        sql = sql & "'" & p.rodzaj_licencji & "', "
        sql = sql & "'" & p.rodzaj_audycji & "', "
        sql = sql & "'" & p.rodzaj_produkcji & "', "
        sql = sql & p.forma_radiowa
        If MPK_dostepne And tryb_obslugi_MPK = 1 Then
            sql = sql & ", '" & p.mpk & "' "
        End If
        sql = sql & ", '" & grupa_docelowa & "'"
        sql = sql & ") "


        a = wykonaj_polecenie_SQL(sql)

        wskaznik_myszy(0)
        Return a


    End Function

    Public Function zapisz_stala_pozycje_audycji(ByRef p As clsStalaPozycjaAudycji) As Integer
        Dim sql As String
        Dim a As Integer
        Dim rok As String
        Dim spr_str As String = "-"
        Dim spr_str2 As String = "-"
        Dim min As Integer
        Dim sek As Integer
        Dim tytul As String
        ' Dim nazwa As String
        Dim imie_nazwisko As String
        Dim komentarz As String
        Dim grupa_docelowa As String = ""

        wskaznik_myszy(1)

        Dim r As clsRodzajProgramowy
        For Each r In colRodzajeProgramowe
            If r.rodzaj = p.rodzaj Then
                spr_str = r.skrot
                Exit For
            End If
        Next
        a = ustal_skrot_podrodzaju(p.rodzaj, p.podrodzaj, spr_str2)

        If p.id_pracownika = 0 Then
            p.imie_nazwisko_pracownika = "-"
        End If
        imie_nazwisko = skoryguj_apostrofy_do_SQL(p.imie_nazwisko_pracownika)
        tytul = skoryguj_apostrofy_do_SQL(p.tytul)
        grupa_docelowa = Trim(skoryguj_apostrofy_do_SQL(p.grupa_docelowa))


        If p.id = 0 Then
            sql = "INSERT INTO STALE_POZYCJE_AUDYCJI "
            sql = sql & " (id_audycji, "
            sql = sql & "tytul, "
            sql = sql & "id_autora, "
            sql = sql & "imie_nazwisko_autora, "
            sql = sql & "rodzaj, "
            sql = sql & "podrodzaj, "
            sql = sql & "dlugosc, "
            sql = sql & "godzina_emisji, "
            sql = sql & "region, "
            sql = sql & "powtorka, "
            sql = sql & "rodzaj_licencji, "
            sql = sql & "rodzaj_audycji, "
            sql = sql & "rodzaj_produkcji, "
            sql = sql & "forma_radiowa "
            If MPK_dostepne And tryb_obslugi_MPK = 1 Then
                sql = sql & ", mpk"
            End If
            sql = sql & ", grupa_docelowa"
            sql = sql & ") "

            sql = sql & " VALUES(" & p.id_audycji & ","
            '            sql = sql & "'" & p.tytul & "',"
            sql = sql & "'" & tytul & "',"
            sql = sql & p.id_pracownika & ","
            sql = sql & "'" & imie_nazwisko & "',"
            sql = sql & "'" & p.rodzaj & "', "
            sql = sql & "'" & p.podrodzaj & "', "
            sql = sql & p.dlugosc & ", "
            sql = sql & "'" & p.godzina_emisji & "', "
            sql = sql & "'" & p.region & "', "
            sql = sql & "'" & p.powtorka & "', "
            sql = sql & "'" & p.rodzaj_licencji & "', "
            sql = sql & "'" & p.rodzaj_audycji & "', "
            sql = sql & "'" & p.rodzaj_produkcji & "', "

            sql = sql & p.forma_radiowa
            If MPK_dostepne And tryb_obslugi_MPK = 1 Then
                sql = sql & ", '" & p.mpk & "' "
            End If
            sql = sql & ",'" & grupa_docelowa & "'"
            sql = sql & ") "
        Else
            sql = "UPDATE STALE_POZYCJE_AUDYCJI set "
            sql = sql & " tytul='" & tytul & "', "
            sql = sql & " id_autora=" & p.id_pracownika & ", "
            sql = sql & " imie_nazwisko_autora='" & imie_nazwisko & "', "
            sql = sql & " rodzaj='" & p.rodzaj & "',"
            sql = sql & " podrodzaj='" & p.podrodzaj & "',"
            sql = sql & " dlugosc=" & p.dlugosc & ","
            sql = sql & "godzina_emisji= '" & p.godzina_emisji & "', "
            sql = sql & "region= '" & p.region & "', "
            sql = sql & "powtorka= '" & p.powtorka & "', "
            sql = sql & "rodzaj_licencji= '" & p.rodzaj_licencji & "', "
            sql = sql & "rodzaj_audycji= '" & p.rodzaj_audycji & "', "
            sql = sql & "rodzaj_produkcji= '" & p.rodzaj_produkcji & "', "
            sql = sql & "forma_radiowa = " & p.forma_radiowa

            If MPK_dostepne And tryb_obslugi_MPK = 1 Then
                sql = sql & ", mpk= '" & p.mpk & "' "
            End If

            sql = sql & ", grupa_docelowa= '" & grupa_docelowa & "' "

            sql = sql & " where id=" & p.id
        End If



        a = wykonaj_polecenie_SQL(sql)

        wskaznik_myszy(0)
        Return a

    End Function


    Public Function zapisz_dlugosc_audycji(ByVal id As Integer, ByVal dlugosc As Integer) As Integer
        Dim sql As String
        Dim a As Integer


        sql = "update wnioski set "
        sql = sql & " dlugosc = " & dlugosc
        sql = sql & " where id= " & id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function
    Public Function zapisz_godzine_emisji(ByVal id As Integer, _
                                            ByVal godzina_rozpoczecia As String, _
                                            ByVal godzina_zakonczenia As String) As Integer
        Dim sql As String
        Dim a As Integer


        sql = "update wnioski set "
        sql = sql & " godzina_rozpoczecia = '" & godzina_rozpoczecia & "', "
        sql = sql & " godzina_zakonczenia = '" & godzina_zakonczenia & "' "
        sql = sql & " where id= " & id

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function

    Public Function zaladuj_opis_audycji(ByVal id As Integer, ByRef opis As String) As Integer
        Dim sql As String


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim k As Integer

        sql = "select info from wnioski where id=" & id


        wskaznik_myszy(1)

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować sie z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True

                opis = dr.GetValue(0)

                odczytano = True
            Else
                odczytano = False
            End If
            wynik = 0
        Catch ex As Exception

            msg = "Wystąpił problem podczas ładowania pełnego opisu audycji: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If odczytano = True Then
            wynik = 0
        Else
            wynik = -1
        End If

        Return wynik


    End Function

    Public Function DoDAJ_rekord_IDENTITY(ByVal komenda As String) As Integer
        'funkcja wykorzystywana jest przy imporcie pracowników z systemu IMPULS
        'po dodaniu nowego pracownika zwraca id rekordu dodanego pracownika

        Dim msg As String
        Dim cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim aa As Integer
        wskaznik_myszy(1)



        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            cmd = New System.Data.SqlClient.SqlCommand(komenda, polaczenie_sql)
            aa = cmd.ExecuteScalar


            wynik = aa
        Catch ex As Exception
            msg = "Wystąpił problem podczas wykonywania polecenia zapisu do bazy danych:" & vbCrLf & ex.Message
            wskaznik_myszy(0)
            MessageBox.Show(msg)
            wynik = -1
        End Try

        Try
            cmd.Dispose()
        Catch ex As Exception

        End Try
        wskaznik_myszy(0)

        Return wynik

    End Function


    Public Function db_zaladuj_szczegoly_wg_mpk(ByVal id_pracownika As Integer, ByVal dp As Date, ByVal dk As Date, ByRef colPozycjeSzczegolow As Collection) As Integer
        Dim a As Integer = 0
        Dim kod_mpk As clsKodMPK
        Dim p As clsPozycjaSzczegolowListyPlac
        Dim wynik As Integer = 0

        Dim sql As String
        Dim dp_str As String
        Dim dk_str As String
        Dim aa As Decimal = 0
        Dim kwota As Decimal = 0

        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")

        For Each kod_mpk In colKodyMPK
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
            sql = sql & " * pozycje_wniosku.ilosc "
            sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
            sql = sql & " AS Suma_wydatkow"
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
            sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
            sql = sql & " WHERE  wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "'"
            sql = sql & " AND pozycje_wniosku.id_pracownika=" & id_pracownika
            If tryb_obslugi_MPK = 0 Then
                sql = sql & " AND  wnioski.mpk = '" & kod_mpk.kod_MPK & "' "
            Else
                sql = sql & " AND  pozycje_wniosku.mpk = '" & kod_mpk.kod_MPK & "' "
            End If


            aa = wczytaj_sume_kosztow(sql)

            If aa < 0 Then
                wynik = -1
                Exit For
            ElseIf aa > 0 Then
                p = New clsPozycjaSzczegolowListyPlac
                p.nazwa = kod_mpk.opis
                p.kwota = aa
                colPozycjeSzczegolow.Add(p, p.nazwa)
            End If
        Next

        Return wynik


    End Function


    Public Function db_zaladuj_szczegoly_wg_zadan(ByVal id_pracownika As Integer, ByVal dp As Date, ByVal dk As Date, ByRef colPozycjeSzczegolow As Collection) As Integer
        Dim a As Integer = 0
        Dim zadanie As clsZadanie
        Dim p As clsPozycjaSzczegolowListyPlac
        Dim sql As String
        Dim dp_str As String
        Dim dk_str As String
        Dim aa As Decimal = 0
        Dim kwota As Decimal = 0
        Dim wynik As Integer = 0

        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")


        For Each zadanie In colZadania
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
            sql = sql & " * pozycje_wniosku.ilosc "
            sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
            sql = sql & " AS Suma_wydatkow"
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
            sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
            sql = sql & " WHERE  wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "'"
            sql = sql & " AND pozycje_wniosku.id_pracownika=" & id_pracownika
            sql = sql & " AND pozycje_wniosku.zadanie = " & zadanie.id

            aa = wczytaj_sume_kosztow(sql)
            If aa < 0 Then
                wynik = -1
                Exit For
            ElseIf aa > 0 Then
                p = New clsPozycjaSzczegolowListyPlac

                p.nazwa = zadanie.nazwa
                p.kwota = aa
                colPozycjeSzczegolow.Add(p, p.nazwa)
            End If

        Next

        Return 0


    End Function


    Public Function db_zaladuj_szczegoly_wg_programow(ByVal id_pracownika As Integer, ByVal dp As Date, ByVal dk As Date, ByRef colPozycjeSzczegolow As Collection) As Integer
        Dim a As Integer = 0
        Dim program As clsProgram
        Dim p As clsPozycjaSzczegolowListyPlac
        Dim sql As String
        Dim dp_str As String
        Dim dk_str As String
        Dim aa As Decimal = 0
        Dim kwota As Decimal = 0
        Dim wynik As Integer = 0

        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")


        For Each program In colProgramy
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
            sql = sql & " * pozycje_wniosku.ilosc "
            sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
            sql = sql & " AS Suma_wydatkow"
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
            sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
            sql = sql & " WHERE  wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "'"
            sql = sql & " AND pozycje_wniosku.id_pracownika=" & id_pracownika
            sql = sql & " AND wnioski.nr_programu = " & program.id

            aa = wczytaj_sume_kosztow(sql)
            If aa < 0 Then
                wynik = -1
                Exit For
            ElseIf aa > 0 Then
                p = New clsPozycjaSzczegolowListyPlac

                p.nazwa = program.nazwa_programu
                p.kwota = aa
                colPozycjeSzczegolow.Add(p, p.nazwa)
            End If

        Next

        Return 0


    End Function


    Public Function db_zaladuj_szczegoly_wg_zrodel_finansowania(ByVal id_pracownika As Integer, ByVal dp As Date, ByVal dk As Date, ByRef colPozycjeSzczegolow As Collection) As Integer
        Dim a As Integer = 0
        Dim zrodlo As clsZrodloFinansowania
        Dim p As clsPozycjaSzczegolowListyPlac
        Dim sql As String
        Dim dp_str As String
        Dim dk_str As String
        Dim aa As Decimal = 0
        Dim kwota As Decimal = 0
        Dim wynik As Integer = 0

        dp_str = Format(dp, "yyyy-MM-dd")
        dk_str = Format(dk, "yyyy-MM-dd")


        For Each zrodlo In colZrodlaFinansowania
            sql = "SET DATEFORMAT YMD "
            sql = sql & "   SELECT DISTINCT Sum(pozycje_wniosku.stawka_podstawowa "
            sql = sql & " * pozycje_wniosku.ilosc "
            sql = sql & "* pozycje_wniosku.wspolczynnik_wyceny) "
            sql = sql & " AS Suma_wydatkow"
            sql = sql & " FROM wnioski RIGHT JOIN pozycje_wniosku"
            sql = sql & " ON wnioski.id = pozycje_wniosku.id_wniosku"
            sql = sql & " WHERE  wnioski.data_emisji BETWEEN '" & dp_str & "' AND '" & dk_str & "'"
            sql = sql & " AND pozycje_wniosku.id_pracownika=" & id_pracownika
            sql = sql & " AND wnioski.zrodlo_finansowania = " & zrodlo.id

            aa = wczytaj_sume_kosztow(sql)
            If aa < 0 Then
                wynik = -1
                Exit For
            ElseIf aa > 0 Then
                p = New clsPozycjaSzczegolowListyPlac

                p.nazwa = zrodlo.nazwa_zrodla
                p.kwota = aa
                colPozycjeSzczegolow.Add(p, p.nazwa)
            End If

        Next

        Return 0


    End Function

    Public Function zaladuj_rodzaje_realizacji()
        Dim p As clsRodzajrealizacji
        Dim k As Integer = 0
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False

        If colRodzajeRealizacji.Count > 0 Then
            Do
                colRodzajeRealizacji.Remove(1)
            Loop While colRodzajeRealizacji.Count > 0
        End If

        Try
            colRodzajeRealizacji.Clear()
        Catch ex As Exception

        End Try


        sql = "select "
        sql = sql & " id, "
        sql = sql & " rodzaj "


        sql = sql & " from rodzaje_realizacji where usuniety=0 "
        sql = sql & " order by id "

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    p = New clsRodzajrealizacji
                    p.id = dr.GetValue(0)
                    p.rodzaj = dr.GetValue(1)
                    colRodzajeRealizacji.Add(p, "-" & p.id)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania rekordu opisującego rodzaj realizacji: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

       

        If colRodzajeRealizacji.Count = 0 Then
            p = New clsRodzajrealizacji
            p.id = 1
            p.rodzaj = "Na żywo"
            colRodzajeRealizacji.Add(p, "-" & p.id)
        End If

        Return wynik




    End Function


    Public Function zaladuj_liste_grup_docelowych()

        Dim p As clsGrupaDocelowa
        Dim k As Integer = 0
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False

        If colListaGrupDocelowych.Count > 0 Then
            Do
                colListaGrupDocelowych.Remove(1)
            Loop While colListaGrupDocelowych.Count > 0
        End If

        Try
            colListaGrupDocelowych.Clear()
        Catch ex As Exception

        End Try


        sql = "select "
        sql = sql & " id, "
        sql = sql & " nazwa_grupy "


        sql = sql & " from grupy_docelowe "
        sql = sql & " order by nazwa_grupy "

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    p = New clsGrupaDocelowa
                    p.id = dr.GetValue(0)
                    p.nazwa = dr.GetValue(1)
                    colListaGrupDocelowych.Add(p, "-" & p.id)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania rekordu opisującego grupę docelową: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If wynik = -1 Then
            If colListaGrupDocelowych.Count = 0 Then
                p = New clsGrupaDocelowa
                p.id = 1
                p.nazwa = "-"
                colListaGrupDocelowych.Add(p, "-" & p.id)
            End If
        End If

        Return wynik



    End Function
    Public Function zaladuj_spis_pasm() As Integer

        Dim p As clsPasmo
        Dim k As Integer = 0
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False

        If colPasma.Count > 0 Then
            Do
                colPasma.Remove(1)
            Loop While colPasma.Count > 0
        End If

        Try
            colPasma.Clear()
        Catch ex As Exception

        End Try


        sql = "select "
        sql = sql & " id, "
        sql = sql & " nazwa_pasma "


        sql = sql & " from pasma "
        sql = sql & " WHERE usuniety=0 "
        sql = sql & " order by nazwa_pasma "

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    p = New clsPasmo
                    p.id = dr.GetValue(0)
                    p.nazwa_pasma = dr.GetValue(1)
                    colPasma.Add(p, "-" & p.id)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania rekordu opisującego pasmo: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If wynik = -1 Then
            If colPasma.Count = 0 Then
                p = New clsPasmo
                p.id = 1
                p.nazwa_pasma = "Pasmo nr 1"
                colPasma.Add(p, "-" & p.id)
            End If
        End If

        Return wynik

    End Function

    Public Function zaladuj_spis_programow() As Integer

        Dim prog As clsProgram
        Dim k As Integer = 0
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim sql As String
        Dim msg As String
        Dim wynik As Integer = 0
        Dim akt_prac As New clsPracownik
        Dim odczytano As Boolean = False

        If colProgramy.Count > 0 Then
            Do
                colProgramy.Remove(1)
            Loop While colProgramy.Count > 0
        End If


        sql = "select "
        sql = sql & " id, "
        sql = sql & " nazwa_programu, "
        sql = sql & " system_emisyjny, "
        sql = sql & " import_danych_emisyjnych,"
        sql = sql & " parametry_polaczenia,"
        sql = sql & " program_id,"
        sql = sql & " przedrostek_sygnatury_archiwalnej"


        sql = sql & " from lista_programow "
        sql = sql & " WHERE usuniety=0 "

        sql = sql & " order by kolejnosc "
        
        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    prog = New clsProgram
                    prog.id = dr.GetValue(0)
                    prog.nazwa_programu = dr.GetValue(1)
                    prog.system_emisyjny = dr.GetValue(2)
                    prog.import_danych_dostepny = dr.GetValue(3)
                    prog.parametry_polaczenia = dr.GetValue(4)
                    prog.zewnetrzny_id_programu = dr.GetValue(5)
                    If prog.system_emisyjny > 0 Then
                        prog.import_danych_dostepny = True
                    End If
                    prog.przedrostek_sygnatury_archiwalnej = dr.GetValue(6)
                    colProgramy.Add(prog, "-" & prog.id)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania rekordu opisującego program: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If wynik = -1 Then
            If colProgramy.Count = 0 Then
                prog = New clsProgram
                prog.id = 1
                prog.nazwa_programu = "Program 1"
                prog.system_emisyjny = 0
                prog.import_danych_dostepny = False
                prog.parametry_polaczenia = ""
                prog.zewnetrzny_id_programu = ""
                prog.przedrostek_sygnatury_archiwalnej = "-"
                colProgramy.Add(prog, "-" & prog.id)
            End If
        End If

        Return wynik

    End Function

    Public Function zapisz_ustawienia_programu(ByRef prog As clsProgram) As Integer

        Dim wynik As Integer = 0
        Dim sql As String
        Dim nazwa As String
        Dim bbb As Integer

        nazwa = Trim(skoryguj_apostrofy_do_SQL(prog.nazwa_programu))
        If prog.import_danych_dostepny Then
            bbb = 1
        Else
            bbb = 0
        End If

        If prog.id = 0 Then
            sql = "Insert into lista_programow "
            sql = sql & " (nazwa_programu, system_emisyjny, parametry_polaczenia, import_danych_emisyjnych, program_id ) "
            sql = sql & "Values ("
            sql = sql & "'" & nazwa & "', " & prog.system_emisyjny & " , '" & prog.parametry_polaczenia & "', " & bbb & ", '" & prog.zewnetrzny_id_programu & "')"
        Else
            sql = "UPDATE lista_programow "
            sql = sql & " set nazwa_programu = '" & nazwa & "',"
            sql = sql & " system_emisyjny= " & prog.system_emisyjny & ","
            sql = sql & "parametry_polaczenia= '" & prog.parametry_polaczenia & "' ,"
            sql = sql & " import_danych_emisyjnych= " & bbb & ","
            sql = sql & " program_id ='" & prog.zewnetrzny_id_programu & "'"
            If oznaczanie_wnioskow_sygnatura_dostepne Then
                sql = sql & ", przedrostek_sygnatury_archiwalnej ='" & Trim(prog.przedrostek_sygnatury_archiwalnej) & "'"
            End If

            sql = sql & " where id =" & prog.id

        End If


        wynik = wykonaj_polecenie_SQL(sql)

        Return wynik

    End Function


    Public Function zapisz_dodatkowe_upowaznienie(ByVal tabela As Integer, _
                                                    ByVal id_audycji_wniosku As Integer, _
                                                    ByVal id_rekordu As Integer, _
                                                    ByVal id_pracownika As Integer, _
                                                    ByVal poziom_uprawnien As Integer) As Integer
        Dim sql As String
        Dim a As Integer

        If tabela = 1 Then
            If id_rekordu = 0 Then
                sql = "Insert into upowaznienia_ramowka (id_audycji, id_pracownika, poziom_uprawnien)"
                sql = sql & " Values ("
                sql = sql & id_audycji_wniosku & ", " & id_pracownika & ", " & poziom_uprawnien & ")"
            Else
                sql = "UPDATE upowaznienia_ramowka "
                sql = sql & " set poziom_uprawnien=" & poziom_uprawnien & ","
                sql = sql & " id_pracownika= " & id_pracownika
                sql = sql & " where id=" & id_rekordu

            End If
        Else
            If id_rekordu = 0 Then
                sql = "Insert into upowaznienia_wnioski (id_wniosku, id_pracownika, poziom_uprawnien)"
                sql = sql & " Values ("
                sql = sql & id_audycji_wniosku & ", " & id_pracownika & ", " & poziom_uprawnien & ")"
            Else
                sql = "UPDATE upowaznienia_wnioski "
                sql = sql & " set poziom_uprawnien=" & poziom_uprawnien & ","
                sql = sql & " id_pracownika= " & id_pracownika
                sql = sql & " where id=" & id_rekordu
            End If
        End If

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function



    Public Function zaladuj_liste_osob_upowaznionych(ByVal id_audycji As Integer, ByRef colLista As Collection) As Integer

        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim msg As String
        Dim up As clsUpowaznienieDodatkowe
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0
        Dim k As Integer = 0


        sql = "select  upowaznienia_ramowka.poziom_uprawnien, upowaznienia_ramowka.id_pracownika from upowaznienia_ramowka "
        sql = sql & " where upowaznienia_ramowka.id_audycji= " & id_audycji

        

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    up = New clsUpowaznienieDodatkowe
                    up.poziom_uprawnien = dr.GetValue(0)
                    up.id_pracownika = dr.GetValue(1)
                    k += 1
                    colLista.Add(up, "-" & k)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania rekordu opisującego upoważnienie dodatkowe: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik


    End Function




    Public Function db_zaladuj_liste_osob_dodatkowo_upowaznionych(ByVal id_wniosku_audycji As Integer, ByVal tabela As Integer, ByRef colLista As Collection) As Integer

        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim msg As String
        Dim up As clsUpowaznienieDodatkowe
        Dim odczytano As Boolean = False
        Dim wynik As Integer = 0

        If colLista.Count > 0 Then
            Do
                colLista.Remove(1)
            Loop While colLista.Count > 0
        End If

        If tabela = 1 Then
            sql = "select upowaznienia_ramowka.id, pracownicy.imie_nazwisko,  upowaznienia_ramowka.poziom_uprawnien, upowaznienia_ramowka.id_pracownika from upowaznienia_ramowka "
            sql = sql & " INNER JOIN pracownicy ON upowaznienia_ramowka.id_pracownika = pracownicy.id"
            sql = sql & " where upowaznienia_ramowka.id_audycji= " & id_wniosku_audycji
        Else
            sql = "select upowaznienia_wnioski.id, pracownicy.imie_nazwisko,  upowaznienia_wnioski.poziom_uprawnien, upowaznienia_wnioski.id_pracownika from upowaznienia_wnioski "
            sql = sql & " INNER JOIN pracownicy ON upowaznienia_wnioski.id_pracownika = pracownicy.id"
            sql = sql & " where upowaznienia_wnioski.id_wniosku= " & id_wniosku_audycji

        End If


        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować się z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)
            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    odczytano = True
                    up = New clsUpowaznienieDodatkowe
                    up.id = dr.GetValue(0)
                    up.imie_nazwisko = dr.GetValue(1)
                    up.poziom_uprawnien = dr.GetValue(2)
                    up.id_pracownika = dr.GetValue(3)

                    colLista.Add(up, "-" & up.id)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wystąpił problem podczas ładowania rekordu opisującego upoważnienie dodatkowe: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        Return wynik


    End Function

    Public Function db_usun_wniosek_honoracyjny(ByVal id As Integer, ByRef opis As clsWniosekHonoracyjny) As Integer
        Dim sql As String
        Dim a As Integer

        sql = "delete from pozycje_wniosku where id_wniosku=" & id

        a = wykonaj_polecenie_SQL(sql)
        If a <> 0 Then
            a = zapisz_rejestr_usunietych_wnioskow("Błąd podczas próby usunięcia wniosku (1)", opis)
            Return a
        End If

        sql = "delete from pozycje_sprawozdania_programowego where id_wniosku=" & id

        a = wykonaj_polecenie_SQL(sql)
        If a <> 0 Then
            a = zapisz_rejestr_usunietych_wnioskow("Błąd podczas próby usunięcia wniosku (2)", opis)
            Return a
        End If

        sql = "delete from udzial_piop where id_wniosku=" & id

        a = wykonaj_polecenie_SQL(sql)
        If a <> 0 Then
            a = zapisz_rejestr_usunietych_wnioskow("Błąd podczas próby usunięcia wniosku (3)", opis)
            Return a
        End If


        sql = "delete from niehonoracyjne_pozycje_wniosku where id_wniosku=" & id

        a = wykonaj_polecenie_SQL(sql)
        If a <> 0 Then
            a = zapisz_rejestr_usunietych_wnioskow("Błąd podczas próby usunięcia wniosku (4)", opis)
            Return a
        End If



        sql = "delete from wnioski where id=" & id
        a = wykonaj_polecenie_SQL(sql)

        If a = 0 Then
            a = zapisz_rejestr_usunietych_wnioskow("", opis)
        End If


    End Function

    Private Function zapisz_rejestr_usunietych_wnioskow(ByVal info1 As String, ByRef opis As clsWniosekHonoracyjny) As Integer

        Dim sql As String
        Dim nazw As String
        Dim komp As String
        Dim info As String
        Dim a As Integer

        nazw = skoryguj_apostrofy_do_SQL(aktualny_uzytkownik.imie_nazwisko)
        komp = skoryguj_apostrofy_do_SQL(nazwa_stacji_komputerowej)
        info = info1 & " - Data emisji: " & Format(opis.data_emisji, "yyyy-MM-dd") & " Tytuł audycji: " & opis.tytul_audycji
        info = skoryguj_apostrofy_do_SQL(info)


        sql = "INSERT INTO REJESTR_USUNIETYCH_WNIOSKOW "
        sql = sql & " (imie_nazwisko, nazwa_komputera, info) "
        sql = sql & " VALUES ("
        sql = sql & "'" & nazw & "', '" & komp & "', '" & info & "')"

        a = wykonaj_polecenie_SQL(sql)

        Return a

    End Function



    Public Function zaladuj_nazwe_pasma(ByVal id As Integer) As String
        Dim sql As String


        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim wynik As String = ""
        Dim odczytano As Boolean = False
        Dim k As Integer

        sql = "select nazwa_pasma from pasma where id=" & id


        wskaznik_myszy(1)

        Try
            If polaczenie_sql.State <> ConnectionState.Open Then
                If polaczenie_sql.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_sql()
                    If a <> 0 Then
                        wskaznik_myszy(0)
                        msg = "Brak połączenia z serwerem SQL."
                        msg = msg & "Proszę skontaktować sie z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                odczytano = True

                wynik = dr.GetValue(0)

                odczytano = True
            Else
                odczytano = False
            End If

        Catch ex As Exception

            msg = "Wystąpił problem podczas odczytu nazwy pasma: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try

        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

     
        Return wynik



    End Function
End Module
