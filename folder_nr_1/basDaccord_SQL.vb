Module basDaccord_SQL




    Public serwer_SQL_daccord As String = ""
    Public uzytkownik_SQL_daccord As String = ""
    Public haslo_SQL_daccord As String = ""
    Public baza_danych_SQL_daccord As String = ""

    Public polaczenie_DB_daccord As System.Data.SqlClient.SqlConnection

    Public czas_serwera_daccord_UTC As Boolean = False




    Public Function otworz_polaczenie_db_daccord()

        Dim kon_str As String
        Dim msg As String
        Dim wynik As Integer = 0

        kon_str = "server=" & serwer_SQL_daccord & "; "
        kon_str = kon_str & "uid=" & uzytkownik_SQL_daccord & "; "
        kon_str = kon_str & "pwd=" & haslo_SQL_daccord & ";"
        kon_str = kon_str & "database=" & baza_danych_SQL_daccord & ";"

        If Not IsNothing(polaczenie_DB_coma) Then
            Try
                polaczenie_DB_coma.Dispose()
            Catch ex As Exception

            End Try
        End If

        Try
            wskaznik_myszy(1)
            polaczenie_DB_daccord = New System.Data.SqlClient.SqlConnection(kon_str)
            polaczenie_DB_daccord.Open()
        Catch ex As Exception
            msg = "Wystąpił problem podczas nawiązywania połączenia z serwerem bazy danych Daccord:" & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1
        End Try

        wskaznik_myszy(0)

        Return wynik


    End Function

    Public Function zaladuj_log_poemisyjny__daccord(ByVal id_programu As Integer, ByVal data_emisji As Date, ByVal godz_pocz As String, ByVal godz_konc As String) As Integer
        Dim a As Integer = 0
        Dim msg As String = ""
        Dim daccord_id_logu As Integer = 0
        Dim poz As clsPozycjaPlanuEmisjiDaccord
        Dim dg_pocz As String = ""
        Dim dg_konc As String = ""

        Dim dg_pocz_d As Date
        Dim dg_konc_d As Date


        'tu jest czas lokalny
        dg_pocz = Format(data_emisji, "yyyy-MM-dd") & " " & Trim(godz_pocz)
        dg_konc = Format(data_emisji, "yyyy-MM-dd") & " " & Trim(godz_konc)





        'xxxxxxxxxxxxxxxxxxxxxx

        'w wywołaniu funckcji podany jest czas lokalny
        dg_pocz_d = CDate(dg_pocz)
        dg_konc_d = CDate(dg_konc)



        If czas_serwera_daccord_UTC Then
            'tu ewentualn zmiana na czas uniwersalny 
            dg_pocz_d = dg_pocz_d.ToUniversalTime
            dg_konc_d = dg_konc_d.ToUniversalTime
        End If


        'xxxxxxxxxxxxx


        If polaczenie_DB_daccord Is Nothing Then
            polaczenie_DB_daccord = New System.Data.SqlClient.SqlConnection
        End If


        If polaczenie_DB_daccord.State <> ConnectionState.Open Then
            If polaczenie_DB_daccord.State <> ConnectionState.Connecting Then

                a = otworz_polaczenie_db_daccord()
                If a <> 0 Then
                    msg = "Brak połączenia z serwerem SQL systemu daccord (1)." & vbCrLf
                    msg = msg & "Proszę skontaktować się z administratorem."
                    wskaznik_myszy(0)
                    MessageBox.Show(msg, naglowek_komunikatow)
                    Return -1
                End If
            End If
        End If


        daccord_id_logu = zaladuj_identyfikator_logu_poemisyjnego_daccord(data_emisji, id_programu)

        If daccord_id_logu = 0 Then
            msg = "Brak logu poemisyjnego, import nie jest możliwy."
            wskaznik_myszy(0)

            MessageBox.Show(msg)
            Return 0
        ElseIf daccord_id_logu < 0 Then
            Return -1
        End If

        wskaznik_myszy(1)

        a = zaladuj_liste_wykorzystanych_nagran_daccord(daccord_id_logu, dg_pocz, dg_konc)
        '    wskaznik_myszy(0)

        If a = 0 Then

            If zaimportowany_plan_emisji_daccord.Count > 0 Then
                For Each poz In zaimportowany_plan_emisji_daccord
                    If poz.item_type = 110 Then
                        poz.plik_audio = pobierz_nazwe_pliku_z_db_daccord(poz.daccord_db_id)
                        If poz.plik_audio <> "-1" Then
                            poz.autor = pobierz_autora_nagrania_z_db_daccord(poz.daccord_db_id)
                        End If
                    End If

                Next
            End If
        End If

        a = grupuj_elementy_zaimportowanego_logu_daccord(godz_pocz)


        wskaznik_myszy(0)


        Return a


    End Function




    Public Function zaladuj_liste_wykorzystanych_nagran_daccord(ByVal id_logu As Integer, ByVal godz_pocz As String, ByVal godz_konc As String) As Integer

        Dim msg As String = ""
        Dim wynik As Integer = 0
        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim odczytano As Boolean = False
        Dim poz As clsPozycjaPlanuEmisjiDaccord
        Dim mark_in As Integer
        Dim mark_out As Integer
        Dim k As Integer = 0
        Dim dg_pocz As Date
        Dim dg_konc As Date
        Dim akt_dg_em As Date
        Dim dl_sek As Integer = 0
        Dim dl_ms As Integer



        'w wywołaniu funckcji podany jest czas lokalny
        dg_pocz = CDate(godz_pocz)
        dg_konc = CDate(godz_konc)



        If czas_serwera_daccord_UTC Then
            'tu ewentualn zmiana na czas uniwersalny 
            dg_pocz = dg_pocz.ToUniversalTime
            dg_konc = dg_konc.ToUniversalTime
        End If





        wskaznik_myszy(1)

        Try
            zaimportowany_plan_emisji_daccord.Clear()

        Catch ex As Exception

        End Try

        sql = "Select "
        sql = sql & " planelemid, "
        sql = sql & "lprogrammitemtypid, "
        sql = sql & "lprogrammitemid, "
        sql = sql & "lmarkin, "
        sql = sql & "lmarkout, "
        sql = sql & "title, "
        sql = sql & "sendeplatz "
        sql = sql & " FROM s_element"
        sql = sql & " where planheaderid = " & id_logu
        sql = sql & " order by planelemid "

        Try
            If polaczenie_DB_daccord.State <> ConnectionState.Open Then
                If polaczenie_DB_daccord.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_db_daccord()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL systemu daccord (3)." & vbCrLf
                        msg = msg & "Proszę skontaktować się z administratorem."
                        wskaznik_myszy(0)

                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_DB_daccord)

            dr = Cmd.ExecuteReader
            Do
                If dr.Read Then
                    odczytano = True
                    If dr.GetValue(0) IsNot System.DBNull.Value Then
                        poz = New clsPozycjaPlanuEmisjiDaccord
                        '                  sql = sql & " planelemid, """
                        '                  sql = sql & "lprogrammitemtypid, "
                        '                  sql = sql & "lprogrammitemid, "
                        '                  sql = sql & "lmarkin, "
                        '                  sql = sql & "lmarkout, "
                        '                  sql = sql & "title, "
                        '                  sql = sql & "sendeplatz "

                        If dr.GetValue(1) IsNot System.DBNull.Value Then
                            poz.item_type = dr.GetValue(1)
                        End If

                        If dr.GetValue(2) IsNot System.DBNull.Value Then
                            poz.daccord_db_id = dr.GetValue(2)
                        End If

                        Try
                            dl_ms = dr.GetValue(4) - dr.GetValue(3)
                        Catch ex As Exception

                        End Try
                        poz.dlugosc_ms = dl_ms
                        poz.dlugosc = wyznacz_dlugosc_sek(dl_ms)

                        If dr.GetValue(5) IsNot System.DBNull.Value Then
                            poz.tytul = dr.GetValue(5)
                        End If


                        If dr.GetValue(6) IsNot System.DBNull.Value Then
                            Try
                                akt_dg_em = dr.GetValue(6)

                                '                                poz.godzina_emisji = Format(dr.GetValue(6), "HH:mm:ss")
                                poz.godzina_emisji = Format(akt_dg_em, "HH:mm:ss")



                            Catch ex As Exception

                            End Try
                            ' Try
                            'poz.godzina_emisji = Format(poz.godzina_emisji, "HH:mm")
                            'Catch ex As Exception

                            '        End Try
                        End If



                        If poz.daccord_db_id > 0 Then
                            If akt_dg_em > dg_pocz Then 'wszystkie zmienne tupy DATA w tym i następnym warunku są w czasie zgodnym z serwerem (UTC jeżeli czas serwera to UTC lub CET jeżeli czas serwera to CET
                                If akt_dg_em < dg_konc Then

                                    If czas_serwera_daccord_UTC Then
                                        akt_dg_em = akt_dg_em.ToLocalTime
                                        poz.godzina_emisji = Format(akt_dg_em, "HH:mm:ss")
                                    End If

                                    k += 1
                                    zaimportowany_plan_emisji_daccord.Add(poz, k & "_")
                                End If
                            End If
                        End If


                    Else
                        wynik = 0
                    End If
                Else
                    odczytano = False
                End If

            Loop While odczytano = True
        Catch ex As Exception
            msg = "Wystąpił problem podczas odczytu listy narań z logu poemisyjnego daccord: " & vbCrLf & ex.Message
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

    Private Function grupuj_elementy_zaimportowanego_logu_daccord(ByVal godz_pocz As String) As Integer

        Dim poz As clsPozycjaPlanuEmisjiDaccord
        Dim dl_ms As Integer = 0
        Dim k As Integer = 0

        Dim p1 As clsPozycjaPlanuEmisjiDaccord


        Try
            skorygowany_plan_emisji_daccord.Clear()
        Catch ex As Exception

        End Try


        For Each poz In zaimportowany_plan_emisji_daccord

            If poz.item_type = 20 Then
                dl_ms += poz.dlugosc_ms
            End If

        Next


        If dl_ms > 0 Then
            k += 1
            poz = New clsPozycjaPlanuEmisjiDaccord
            poz.godzina_emisji = Trim(godz_pocz)
            poz.tytul = "Nagrania muzyczne"

            poz.dlugosc_ms = dl_ms
            poz.dlugosc = wyznacz_dlugosc_sek(dl_ms)


            skorygowany_plan_emisji_daccord.Add(poz, k & "_")
        End If

        dl_ms = 0
        For Each poz In zaimportowany_plan_emisji_daccord

            If poz.item_type = 30 Then
                dl_ms += poz.dlugosc_ms
            End If

        Next


        If dl_ms > 0 Then
            poz = New clsPozycjaPlanuEmisjiDaccord
            poz.godzina_emisji = Trim(godz_pocz)
            poz.tytul = "Jingle"

            poz.dlugosc_ms = dl_ms
            poz.dlugosc = wyznacz_dlugosc_sek(dl_ms)
            k += 1
            skorygowany_plan_emisji_daccord.Add(poz, k & "_")

        End If

        dl_ms = 0
        For Each poz In zaimportowany_plan_emisji_daccord

            If poz.item_type = 40 Then
                dl_ms += poz.dlugosc_ms
            End If

        Next

        If dl_ms > 0 Then
            poz = New clsPozycjaPlanuEmisjiDaccord
            poz.godzina_emisji = Trim(godz_pocz)
            poz.tytul = "Reklama"

            poz.dlugosc_ms = dl_ms
            poz.dlugosc = wyznacz_dlugosc_sek(dl_ms)

            k += 1
            skorygowany_plan_emisji_daccord.Add(poz, k & "_")

        End If


        For Each poz In zaimportowany_plan_emisji_daccord

            If poz.item_type = 110 Then
                k += 1
                skorygowany_plan_emisji_daccord.Add(poz, k & "_")
            End If

        Next




    End Function



    Private Function wyznacz_dlugosc_sek(ByVal dl_ms As Integer) As Integer
        Dim tmp_sek As Integer = 0


        tmp_sek = Int(dl_ms / 1000)
        If (dl_ms - (tmp_sek * 1000)) > 499 Then
            tmp_sek += 1
        End If


        Return tmp_sek

    End Function



    Public Function zaladuj_identyfikator_logu_poemisyjnego_daccord(ByVal data_emisji As Date, ByVal id_programu As Integer) As Integer

        Dim msg As String = ""
        Dim wynik As Integer = 0
        Dim sql As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim odczytano As Boolean = False

        sql = "SET DATEFORMAT YMD "
        sql = sql & " select planheaderid from s_planhdh where filetype= 'I' and programid=" & id_programu & " and pdate= '" & Format(data_emisji, "yyyy-MM-dd") & "'"

        wskaznik_myszy(1)


        Try
            If polaczenie_DB_daccord.State <> ConnectionState.Open Then
                If polaczenie_DB_daccord.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_db_daccord()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL systemu daccord (2)." & vbCrLf
                        msg = msg & "Proszę skontaktować się z administratorem."
                        wskaznik_myszy(0)

                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_DB_daccord)

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
            msg = "Wystąpił problem podczas odczytu identyfikatora logu poemisyjnego daccord: " & vbCrLf & ex.Message
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
        End If
        wskaznik_myszy(0)

        Return wynik

    End Function


    Public Function pobierz_autora_nagrania_z_db_daccord(ByVal programitemID As Integer) As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim odczytano As Boolean = False
        Dim wynik As String = "0"
        Dim sql As String
        Dim msg As String = ""


       
        ' tu dane o autorze program pobiera z tabeli PROGRAMMITEM
        'wg Marka Kotera ta wartość to Redaktor a nie autor
        '        sql = " select szname "
        '        sql = sql & " from d2user "
        '        sql = sql & " where lid= (select linputuserid from programmitem where lid=" & programitemID & ")"
        'zmiana
        '15 grudnia
        'program pobiera autora z pola CREATORID tabeli m_item

        '        sql = " select szname "
        '       sql = sql & " from d2user "
        '      sql = sql & " where lid= (select creatorid from m_item where programmitemid=" & programitemID & ")"

        ' jak się okazało to też nie jest autor




        'widok VPIBO - powiązany z rekordem programitem

        '        SELECT TOP 1000 [id]
        '     ,[itemtyp]
        '     ,[inputuserid]
        '     ,[isprivate]
        ''     ,[programid]
        '     ,[duration]
        '     ,[playoutfunction]
        '     ,[playoutid]
        '     ,[projectid]
        '     ,[playoutmainfrmid]
        '     ,[projectmainfrmid]
        '     ,[playoutextension]
        '     ,[projectextension]
        '       FROM [daccord].[dbo].[vpibo]
        '       where id = 4592927

        '' - powyższego nie sprawdzałem


        '19 grudnia znalazłem tabelę A_WORD
        'tam jest dużo danych
        '
        '       SELECT TOP 1000 [lid]
        '     ,[sztitle]
        '     ,[sztitlematch]
        '     ,[szsubtitle]
        '     ,[szsubtitlematch]
        '     ,[szothertitle]
        '     ,[szothertitlematch]
        '     ,[szshortinfo]
        '     ,[szshortinfomatch]
        '     ,[szrecordplace]
        '     ,[szrecordplacematch]
        '     ,[lrecorddate]
        '     ,[ilongtimearchive]
        '    ,[laseriesid]
        '    ,[obs_linfotextid]
        '     ,[llanguageid]
        '     ,[szalanguagecomment]
        '     ,[lakindofcreationid]
        '     ,[lrfaid]
        '    ,[szarchivnr]
        '    ,[loriginrfaid]
        '     ,[szoriginnr]
        '     ,[szproduktionnr]
        '    ,[leinspielstationid]
        '    ,[szkostenstelle]
        '    ,[szkennziffer]
        '     ,[szabrechnr]
        '     ,[szfremdobjid]
        ''    ,[lauthorid]   -  referencja do tabeli aperson
        '    ,[lredakteurid]  - referencja do tabeli d2user
        '    ,[laredaktionid]
        '   ,[lressortid]               ----- tu 
        '   ,[laprograreaid]
        '   ,[lamaterialartid]
        ''   ,[lapresentationid]
        '   ,[lacontentsid]
        '   ,[lpublisherid]
        '   ,[btobedeleted]
        '   ,[lamediumid]
        '   ,[dtdayofdelete]
        '   ,[lstatus_flag_a]
        '   ,[lstatus_flag_b]
        '   ,[icountwords]
        '   ,[icountaudios]
        '   ,[lwordduration]
        '   ,[szbeitragstyp]
        '   ,[szthemenkennung]
        '   ,[llabelid]
        '   ,[szordernr]
        '   ,[szkostentraeger]
        '   ,[lquellaredaktionid]
        '   ,[fremdautor]
        '   ,[szepisodetitle]
        '   ,[szepisodetitlematch]
        '   ,[lpodcastseriesid]
        '   ,[podcastautor]
        '   ,[dtdate2publish]
        '   ,[dtpublisheddate]
        '   ,[podcastautormatch]
        '   ,[extfilename]
        '   ,[extfiletype]
        '   ,[extfilelength]
        '   ,[lScope]
        '     '  FROM [daccord].[dbo].[a_word]


        '        SELECT TOP 1000 [lid]
        '     ,[szname]
        '    ,[szmatch]
        '   ,[szsoundex]
        '    FROM [daccord].[dbo].[aperson]
        '   where lid = 1795075


        'program w pierszej kolejności czyta autora z pola LAUTHORID
        'jeżeli to ppole jest puste to pobiera dane z pola LREDAKTEURID


        sql = " select szname "
        sql = sql & " from aperson "
        sql = sql & " where lid= (select lauthorid from a_word where lid =" & programitemID & ")"
        '  sql = sql & " where lid= (select lredakteurid from a_word where lid =" & programitemID & ")"



        Try
            If polaczenie_DB_daccord.State <> ConnectionState.Open Then
                If polaczenie_DB_daccord.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_db_daccord()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL systemu daccord (4)." & vbCrLf
                        msg = msg & "Proszę skontaktować się z administratorem."
                        wskaznik_myszy(0)
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_DB_daccord)

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
            msg = "Wystąpił problem podczas odczytu autora nagrania z bazy danych daccord: " & vbCrLf & ex.Message
            wskaznik_myszy(0)
            Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
            wynik = -1

        End Try



        If odczytano = False Then
            If Val(wynik) = 0 Then
                'najprawdopodobniej pole AUTHOR było puste dlatego drugi odczyt z pola redakkteur

                sql = " select szname "
                sql = sql & " from d2user "
                'sql = sql & " where lid= (select lauthorid from a_word where lid =" & programitemID & ")"
                sql = sql & " where lid= (select lredakteurid from a_word where lid =" & programitemID & ")"
                Try
                    dr.Close()
                Catch ex As Exception

                End Try


                Try

                    'Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_DB_daccord)
                    Cmd.CommandText = sql

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
                    msg = "Wystąpił problem podczas odczytu autora nagrania z bazy danych daccord: " & vbCrLf & ex.Message
                    wskaznik_myszy(0)
                    Windows.Forms.MessageBox.Show(msg, "BŁĄD !", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation, Windows.Forms.MessageBoxDefaultButton.Button1)
                    wynik = -1

                End Try




            End If
        End If



        Try
            dr.Close()
            Cmd.Dispose()

        Catch ex As Exception

        End Try

        If odczytano = False Then
            'to mogło byc puste pole w opisie rekordu w tabeli a_word
            wynik = ""
        End If

        Return Trim(wynik)





    End Function



    Public Function pobierz_nazwe_pliku_z_db_daccord(ByVal programitemID As Integer) As String
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand
        Dim odczytano As Boolean = False
        Dim wynik As String = ""
        Dim sql As String
        Dim msg As String = ""

        'to jest WIDOK w bazie danych 
        '        SELECT TOP 1000 [programmitemid]
        '      ,[mediaitemid]
        '      ,[ctime]
        '      ,[creatorid]
        '      ,[funcdomik]
        '      ,[locationid]
        '      ,[isreadonly]
        '      ,[filename]
        '        FROM [daccord].[dbo].[m_vallcopies]
        '        where [programmitemid] = 4527754



        sql = " SELECT filename "
        sql = sql & " FROM m_vallcopies"
        sql = sql & " where programmitemid = " & programitemID


        Try
            If polaczenie_DB_daccord.State <> ConnectionState.Open Then
                If polaczenie_DB_daccord.State <> ConnectionState.Connecting Then
                    Dim a As Integer
                    a = otworz_polaczenie_db_daccord()
                    If a <> 0 Then
                        msg = "Brak połączenia z serwerem SQL systemu daccord (5)." & vbCrLf
                        msg = msg & "Proszę skontaktować się z administratorem."
                        wskaznik_myszy(0)
                        MessageBox.Show(msg, naglowek_komunikatow)
                        Return -1
                    End If
                End If
            End If

            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_DB_daccord)

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
            msg = "Wystąpił problem podczas odczytu nazwy pliku audio bazy danych daccord: " & vbCrLf & ex.Message
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
        End If

        Return wynik





    End Function





    Public Function zapisz_ustawienia_servera_DB_DACCORD(ByVal s As String, ByVal db As String, ByVal user As String, ByVal pwd As String) As Integer


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
        If Len(dane) = 0 Then 'wystąpił problem podczas szyfrowania
            Return -1
        End If

        dane = skoryguj_apostrofy_do_SQL(dane)

        tmp_str = dane





        If colWczytaneUstawienia.Count > 0 Then
            For Each rek_ust In colWczytaneUstawienia
                If UCase(rek_ust.nazwa) = "DB_DACCORD_POLACZ" Then
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
            sql = sql & "('DB_DACCORD_POLACZ', "
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

    Public Function wczytaj_ustawienia_serwera_DB_DACCORD() As Integer
        Dim dane As String
        Dim tmp_string As String = ""
        Dim rek_ust As clsUstawienia


        If colWczytaneUstawienia.Count > 0 Then
            For Each rek_ust In colWczytaneUstawienia
                If UCase(rek_ust.nazwa) = "DB_DACCORD_POLACZ" Then
                    tmp_string = rek_ust.wartosc
                    Exit For
                End If
            Next
        End If


        If Len(tmp_string) > 0 Then
            dane = szyrfuj_usytawienia_polaczenia(tmp_string)
            serwer_SQL_daccord = Left(dane, InStr(dane, ";") - 1)
            dane = Right(dane, Len(dane) - Len(serwer_SQL_daccord) - 1)

            uzytkownik_SQL_daccord = Left(dane, InStr(dane, ";") - 1)
            dane = Right(dane, Len(dane) - Len(uzytkownik_SQL_daccord) - 1)

            haslo_SQL_daccord = Left(dane, InStr(dane, ";") - 1)
            baza_danych_SQL_daccord = Right(dane, Len(dane) - Len(haslo_SQL_daccord) - 1)
        End If


        Return 0

    End Function





End Module
