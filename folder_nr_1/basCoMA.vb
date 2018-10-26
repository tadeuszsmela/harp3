' zmiana z dnia 25 pazdziernika - godzina 13:00

Module basCoMA

    Public kontrola_listy_utworow_COMA_dostepna As Boolean = False


    Public serwer_SQL_coma As String = ""
    Public baza_danych_SQL_coma As String = ""
    Public uzytkownik_SQL_coma As String = ""
    Public haslo_SQL_coma As String = ""
    Public polaczenie_DB_coma As System.Data.SqlClient.SqlConnection


    Public Function zapisz_ustawienia_servera_DB_COMA(ByVal s As String, ByVal db As String, ByVal user As String, ByVal pwd As String) As Integer


        Dim dane As String

        Dim tmp_str As String
        'Dim msg As String
        'Dim kl As Microsoft.Win32.RegistryKey
        Dim czy_koniec As Boolean = False
        Dim k As Integer = 0
        Dim rek_ust As clsUstawienia
        Dim id_rek As Integer = 0
        Dim a As Integer
        Dim sql As String
        Dim kontr As Boolean = False

        tmp_str = s & ";"
        tmp_str = tmp_str & user & ";"
        tmp_str = tmp_str & pwd & ";"
        tmp_str = tmp_str & db


        dane = szyrfuj_usytawienia_polaczenia(tmp_str)
        If Len(dane) = 0 Then 'wyst¹pi³ problem podczas szyfrowania
            Return -1
        End If

        tmp_str = dane

        If colWczytaneUstawienia.Count > 0 Then
            For Each rek_ust In colWczytaneUstawienia
                If UCase(rek_ust.nazwa) = "DB_COMA_POLACZENIE" Then
                    id_rek = rek_ust.id_rekordu
                    kontr = True
                    Exit For
                End If
            Next
        End If

        If id_rek = 0 Then
            sql = "insert into ustawienia "
            sql = sql & "( nazwa,wartosc )"
            sql = sql & "VALUES "
            sql = sql & "('DB_COMA_POLACZENIE', "
            sql = sql & "'" & tmp_str & "')"
        Else
            sql = "UPDATE ustawienia "
            sql = sql & "set wartosc='" & tmp_str & "'"
            sql = sql & " where id=" & id_rek
        End If


        a = wykonaj_polecenie_SQL(sql)

        Try
            If a = 0 Then
                If kontr Then
                    rek_ust.wartosc = tmp_str
                End If
            End If

        Catch ex As Exception

        End Try

        Return a


    End Function


    Public Function wczytaj_ustawienia_serwera_DB_COMA() As Integer
        Dim dane As String
        Dim tmp_string As String = ""
        Dim rek_ust As clsUstawienia


        If colWczytaneUstawienia.Count > 0 Then
            For Each rek_ust In colWczytaneUstawienia
                If UCase(rek_ust.nazwa) = "DB_COMA_POLACZENIE" Then
                    tmp_string = rek_ust.wartosc
                    Exit For
                End If
            Next
        End If


        If Len(tmp_string) > 0 Then
            dane = szyrfuj_usytawienia_polaczenia(tmp_string)
            serwer_SQL_coma = Left(dane, InStr(dane, ";") - 1)
            dane = Right(dane, Len(dane) - Len(serwer_SQL_coma) - 1)

            uzytkownik_SQL_coma = Left(dane, InStr(dane, ";") - 1)
            dane = Right(dane, Len(dane) - Len(uzytkownik_SQL_coma) - 1)

            haslo_SQL_coma = Left(dane, InStr(dane, ";") - 1)
            baza_danych_SQL_coma = Right(dane, Len(dane) - Len(haslo_SQL_coma) - 1)
        End If


        Return 0

    End Function






    Public Function wyznacz_dlugosc_audycji(ByVal id As Integer, ByVal plan_wykonanie As Integer) As Integer
        'plan_wykonanie         1 wykonanie
        '                       2- plan
        '               
        Dim wynik As Integer = 0
        Dim sql As String
        Dim a As Integer
        Dim d1 As Integer = 0
        Dim d2 As Integer = 0

        If plan_wykonanie = 1 Then
            sql = "select sum(dlugosc_sekund) as dlug from pozycje_audycji where id_audycji=" & id

            a = db_wyznacz_dlugosc_audycji(sql)
            If a = -1 Then
                Return 0
            Else
                d1 = a
            End If

            sql = "select sum(dlugosc_sekund) as dlug from pozycje_audycji_obce where id_audycji=" & id

            a = db_wyznacz_dlugosc_audycji(sql)
            If a = -1 Then
                Return 0
            Else
                d2 = a
            End If
        Else
            sql = "select sum(dlugosc_sekund) as dlug from pozycje_audycji_planowane where id_audycji=" & id

            a = db_wyznacz_dlugosc_audycji(sql)
            If a = -1 Then
                Return 0
            Else
                d1 = a
            End If
            d2 = 0
        End If

        wynik = d1 + d2
        Return wynik

    End Function


    Public Function wyznacz_dlugosc_audycji(ByVal id As Integer) As Integer
        Dim wynik As Integer = 0
        Dim sql As String
        Dim a As Integer
        Dim d1 As Integer = 0
        Dim d2 As Integer = 0

        sql = "select sum(dlugosc_sekund) as dlug from pozycje_audycji where id_audycji=" & id

        a = db_wyznacz_dlugosc_audycji(sql)
        If a = -1 Then
            Return 0
        Else
            d1 = a
        End If

        sql = "select sum(dlugosc_sekund) as dlug from pozycje_audycji_obce where id_audycji=" & id

        a = db_wyznacz_dlugosc_audycji(sql)
        If a = -1 Then
            Return 0
        Else
            d2 = a
        End If
        
        wynik = d1 + d2
        Return wynik

    End Function

    Public Function db_wyznacz_dlugosc_audycji(ByVal sql As String) As Integer

        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        Dim msg As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False



        Try
            If polaczenie_DB_coma.State <> ConnectionState.Open Then
                If polaczenie_DB_coma.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_COMA()
                    If a <> 0 Then
                        msg = "Brak po³¹czenia z serwerem SQL."
                        msg = msg & "Proszê skontaktowaæ siê z administratorem"
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_DB_coma)

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
            msg = "Wyst¹pi³ problem podczas odczytu d³ugoœci wykorzystanych utworów: " & vbCrLf & ex.Message
            Windows.Forms.MessageBox.Show(msg, "B£¥D !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
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

    Public Function ustal_coma_id_audycji(ByVal harp_id As Integer) As Integer


        Dim sql As String
        Dim k As Integer = 0
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim wynik As Integer = 0
        Dim msg As String


        sql = "select id_audycji from audycje where harp_id_wniosku=" & harp_id


        Try
            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_DB_coma)

            dr = Cmd.ExecuteReader

            If dr.Read Then
                wynik = dr.GetValue(0)
            End If
            
        Catch ex As Exception
            msg = "Wyst¹pi³ problem podczas ustalania identyfikatora audycji w bazie danych CoMA: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "B£¥D !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1
        End Try

        Try
            dr.Close()
            Cmd.Dispose()
        Catch ex As Exception

        End Try

        Return wynik

    End Function

    Public Function otworz_polaczenie_COMA()


        Dim kon_str As String
        Dim msg As String
        Dim wynik As Integer = 0

        kon_str = "server=" & serwer_SQL_coma & "; "
        kon_str = kon_str & "uid=" & uzytkownik_SQL_coma & "; "
        kon_str = kon_str & "pwd=" & haslo_SQL_coma & ";"
        kon_str = kon_str & "database=" & baza_danych_SQL_coma & ";"

        If Not IsNothing(polaczenie_DB_coma) Then
            Try
                polaczenie_DB_coma.Dispose()
            Catch ex As Exception

            End Try
        End If

        Try
            wskaznik_myszy(1)
            polaczenie_DB_coma = New System.Data.SqlClient.SqlConnection(kon_str)
            polaczenie_DB_coma.Open()
        Catch ex As Exception
            msg = "Wyst¹pi³ problem podczas nawi¹zywania po³¹czenia z serwerem bazy danych CoMA:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1
        End Try

        wskaznik_myszy(0)

        Return wynik


    End Function


    Public Function zamien_sekundy_na_str(ByVal dlug As Integer) As String
        Dim tmp_str As String
        Dim g As Integer
        Dim m As Integer
        Dim s As Integer
        Dim r As Integer


        g = Int(dlug / 3600) 'godzin
        r = dlug - g * 3600
        m = Int(r / 60) 'minut
        s = r - m * 60 'sekund

        If g = 0 Then
            tmp_str = "00"
        ElseIf g < 10 Then
            tmp_str = "0" & g
        Else
            tmp_str = g
        End If


        If m = 0 Then
            tmp_str = tmp_str & ":00"
        ElseIf m < 10 Then
            tmp_str = tmp_str & ":0" & m
        Else
            tmp_str = tmp_str & ":" & m
        End If

        If s = 0 Then
            tmp_str = tmp_str & ":00"
        ElseIf s < 10 Then
            tmp_str = tmp_str & ":0" & s
        Else
            tmp_str = tmp_str & ":" & s
        End If

        Return tmp_str


    End Function


End Module
