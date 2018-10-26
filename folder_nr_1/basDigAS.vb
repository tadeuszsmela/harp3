Module basDigAS

    Public import_digas_dostepny As Boolean = False
    Public import_digas_czytaj_pole_komentarz As Boolean = False

    Public colPozycjeUstawienImportu As New Collection

    Public colSpisAudycjiDigaROC As New Collection
    Public colPOzycjeRaportuPoemisyjnegoDIGAS As New Collection
    Public colRaportDigAS As New Collection

    Public Function zaladuj_schemat_importu_digas() As Integer

        Dim msg As String
        Dim poz As clsKlasaImportowanychNagranDigas

        Dim sql As String
        Dim wynik As Integer = 0
        Dim odczytano As Boolean = False
        Dim k As Integer = 0
        Dim dr As System.Data.SqlClient.SqlDataReader
        Dim Cmd As System.Data.SqlClient.SqlCommand

        If colPozycjeUstawienImportu.Count > 0 Then
            Do
                colPozycjeUstawienImportu.Remove(1)
            Loop While colPozycjeUstawienImportu.Count > 0
        End If

        sql = "Select "
        sql = sql & " id,"
        sql = sql & " klasa, "
        sql = sql & " tryb_importu, "
        sql = sql & " nazwa_grupy, "
        sql = sql & " rodzaj, "
        sql = sql & " podrodzaj, "
        sql = sql & "rodzaj_audycji "
        sql = sql & " from ustawienia_importu_digas "
        sql = sql & " order by id "

        Try
            Cmd = New System.Data.SqlClient.SqlCommand(sql, polaczenie_sql)

            dr = Cmd.ExecuteReader

            Do
                If dr.Read Then
                    k += 1
                    odczytano = True
                    poz = New clsKlasaImportowanychNagranDigas
                    poz.id = dr.GetValue(0)
                    poz.klasa = dr.GetValue(1)
                    poz.tryb_importu = dr.GetValue(2)
                    poz.nazwa_grupy = dr.GetValue(3)
                    poz.rodzaj = dr.GetValue(4)
                    poz.podrodzaj = dr.GetValue(5)
                    poz.rodzaj_audycji_skrot = dr.GetValue(6)
                    colPozycjeUstawienImportu.Add(poz, "_" & poz.id)
                Else
                    odczytano = False
                End If
            Loop While odczytano = True
            wynik = 0
        Catch ex As Exception
            msg = "Wyst¹pi³ problem podczas odczytu ustawieñ importu DigAS: " & vbCrLf & ex.Message
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



    Public Function zaladuj_spis_audycji_DIGAROC(ByVal nr_programu As Integer, ByVal d_em As Date) As Integer
        Dim msg As String
        Dim pr As clsProgram
        Dim sciezka_raportu As String
        Dim m As Integer
        Dim d As Integer
        Dim r As Integer
        Dim plk As String

        Dim aud As clsAudycjaDigas
        Dim d_em_str As String
        Dim gp_em As String
        Dim gk_em As String

        Dim tmp_g As Date




        colSpisAudycjiDigaROC.Clear()
        If colSpisAudycjiDigaROC.Count > 0 Then
            Do
                colSpisAudycjiDigaROC.Remove(1)
            Loop While colSpisAudycjiDigaROC.Count > 0
        End If



        d = d_em.Day
        m = d_em.Month
        r = d_em.Year

        '        pr = colProgramy.Item(nr_programu)

        For Each pr In colProgramy
            If pr.id = nr_programu Then
                Exit For
            End If
        Next


        sciezka_raportu = pr.parametry_polaczenia




        If Right(sciezka_raportu, 1) <> "\" Then
            sciezka_raportu = sciezka_raportu & "\"
        End If

        sciezka_raportu = sciezka_raportu & r & "\" & m & "\" & d & "\"

        plk = UCase(System.Convert.ToString(d, 16))


        Do
            plk = "0" & plk
        Loop While Len(plk) < 8

        plk = plk & ".XML"

        plk = sciezka_raportu & plk

        If System.IO.File.Exists(plk) = False Then
            msg = "Brak pliku " & plk
            MessageBox.Show(msg)
            Return -1
        End If

        Dim ds As New DataSet()
        Try
            ds.ReadXml(plk)
        Catch ex As Exception
            wskaznik_myszy(0)
            msg = "Wyst¹pi³ problem podczas odczytu pliku " & plk
            msg = msg & vbCrLf & ex.Message
            MessageBox.Show(msg)
            Return -1
        End Try



        For Each row As DataRow In ds.Tables("Show").Rows
            aud = New clsAudycjaDigas
            tmp_g = row("Time_Start").ToString
            aud.godzina_rozpoczecia = Format(tmp_g, "HH:mm:ss")
            tmp_g = row("Time_Stop").ToString
            aud.godzina_zakonczenia = Format(tmp_g, "HH:mm:ss")

            '            aud.godzina_zakonczenia = Format(row("Time_Stop").ToString, "HH:mm:ss")
            aud.tytul = row("Name").ToString
            aud.digas_id = row("id")
            colSpisAudycjiDigaROC.Add(aud, aud.digas_id)

        Next


        Try
            ds.Dispose()
        Catch ex As Exception

        End Try

        Return 0

    End Function


    Public Function zaladuj_audycje_raportu_poemisyjnego_digas(ByRef aud As clsAudycjaDigas, ByVal nr_programu As Integer, ByVal data_emisji As Date) As Integer
        Dim dl_ms As Integer
        Dim dl_sek As Integer

        Dim msg As String
        Dim pr As clsProgram
        Dim sciezka_raportu As String
        Dim m As Integer
        Dim d As Integer
        Dim rok As Integer
        Dim plk As String

        Dim d_em_str As String
        Dim gp_em As String
        Dim gk_em As String

        Dim tmp_g As Date
        Dim p As clsPozycjaRaportuDigas
        Dim track_number As Integer = 0
        Dim tmp_str As String = ""
        Dim k As Integer = 0
        Dim tmp_pl As String = ""
        Dim tmp_kat As String = ""
        Dim tmp_ext As String
        Dim znak As String


        Dim r() As DataRow




        d = data_emisji.Day
        m = data_emisji.Month
        rok = data_emisji.Year

        '       pr = colProgramy.Item(nr_programu)

        For Each pr In colProgramy
            If pr.id = nr_programu Then
                Exit For
            End If
        Next

        sciezka_raportu = pr.parametry_polaczenia




        If Right(sciezka_raportu, 1) <> "\" Then
            sciezka_raportu = sciezka_raportu & "\"
        End If

        sciezka_raportu = sciezka_raportu & rok & "\" & m & "\" & d & "\Shows\"


        plk = aud.digas_id & ".XML"

        plk = sciezka_raportu & plk

        If System.IO.File.Exists(plk) = False Then
            wskaznik_myszy(0)

            msg = "Brak pliku raportu poemisyjnego " & plk
            MessageBox.Show(msg)
            Return -1
        End If

        Dim ds As New DataSet()
        Try
            ds.ReadXml(plk)
        Catch ex As Exception
            wskaznik_myszy(0)
            msg = "Wyst¹pi³ problem podczas odczytu pliku " & plk
            msg = msg & vbCrLf & ex.Message
            MessageBox.Show(msg)
            Return -1
        End Try


        For Each row As DataRow In ds.Tables("Track").Rows

            track_number = 0
            Try
                tmp_str = row("Number").ToString
                track_number = Val(tmp_str)

            Catch ex As Exception
                MessageBox.Show(ex.Message)

            End Try

            If track_number = 1000 Then
                r = row.GetChildRows("Track_Element")

                For Each cr As DataRow In r
                    If cr("SendState").ToString = "Sent" Then
                        '<SendState>Sent</SendState>
                        k += 1
                        p = New clsPozycjaRaportuDigas
                        p.xml_id = cr("ID").ToString

                        p.db_ref = cr("DBRef").ToString
                        p.klasa = cr("Class").ToString
                        tmp_g = cr("Time_RealStart").ToString
                        p.time_start = Format(tmp_g, "HH:mm:ss")

                        tmp_g = cr("Time_RealStop").ToString
                        p.time_stop = Format(tmp_g, "HH:mm:ss")

                        '
                        p.duration = Val(cr("Time_RealDuration").ToString)
                        If p.duration < 0 Then
                            p.duration = -1
                            p.skorygowany_duration = 0
                        End If

                        p.tytul = Trim(cr("Title").ToString)
                        Try
                            p.autor = Trim(cr("News_Author").ToString)

                        Catch ex As Exception

                        End Try

                        Try
                            p.rodzaj = Trim(cr("News_Ressort").ToString)
                        Catch ex As Exception
                            p.rodzaj = "-"
                        End Try

                        Try
                            p.podrodzaj = Trim(cr("News_SubRessort").ToString)
                        Catch ex As Exception
                            p.podrodzaj = "-"
                        End Try
                        If import_digas_czytaj_pole_komentarz Then
                            Try
                                p.komentarz = Trim(cr("Comment").ToString)
                            Catch ex As Exception
                                p.komentarz = ""
                            End Try

                        End If

                        tmp_str = ""
                        Try
                            tmp_str = Trim(cr("File_Filename0"))
                            'w radiu Gorzów baza jest replikowana
                            'nie ma tam w pliku XML pola File_Filename0
                            'jest za to pole zreplikowanego pliku o nazwie File_Filename1
                            'tej nazwy nie bede zapisywa³ w bazie danych
                            'te dane sa potrzebne tylko w sytuacji gdy dopisywany jest nowy rekord do tabeli muzyki, 
                            'przy takim utworze, pobranym z raportu Radia Gorzów nie bedzie zapisana œcie¿ka i nazwa pliku
                        Catch ex As Exception

                        End Try
                        p.plik_audio = tmp_str


                        colPOzycjeRaportuPoemisyjnegoDIGAS.Add(p, p.xml_id & "_" & p.db_ref & "_" & p.time_start)

                    End If
                Next
            End If

        Next


        Try
            ds.Dispose()
        Catch ex As Exception

        End Try
        Return 0

    End Function

    Public Sub wyczysc_raport_digas()

        colPOzycjeRaportuPoemisyjnegoDIGAS.Clear()

        If colPOzycjeRaportuPoemisyjnegoDIGAS.Count > 0 Then
            Do
                colPOzycjeRaportuPoemisyjnegoDIGAS.Remove(1)
            Loop While colPOzycjeRaportuPoemisyjnegoDIGAS.Count > 0
        End If

        colRaportDigAS.Clear()

        If colRaportDigAS.Count > 0 Then
            Do
                colRaportDigAS.Remove(1)
            Loop While colRaportDigAS.Count > 0
        End If

    End Sub
    Public Function zaladuj_liste_emisyjna_digas(ByVal program As Integer, _
                                                    ByVal data_emisji As Date, _
                                                    ByVal godzina_rozpoczecia As String, _
                                                    ByVal godzina_zakonczenia As String) As Integer

        'funkcja zwraca ³¹czna d³ugoœc w milisekudach

        Dim a As Integer
        Dim aud As clsAudycjaDigas
        Dim poz As clsPozycjaRaportuDigas

        Dim godz_pocz As Date
        Dim godz_konc As Date
        Dim gp1 As Date
        Dim gk1 As Date
        Dim kontr As Boolean = False

        Dim colTymczasowaListaEmisyjna As New Collection
        Dim k As Integer


        godz_pocz = CDate(godzina_rozpoczecia)
        godz_konc = CDate(godzina_zakonczenia)

        wyczysc_raport_digas()

        a = zaladuj_spis_audycji_DIGAROC(program, data_emisji)

        If colSpisAudycjiDigaROC.Count > 0 Then
            For Each aud In colSpisAudycjiDigaROC
                kontr = False
                gp1 = CDate(aud.godzina_rozpoczecia)
                gk1 = CDate(aud.godzina_zakonczenia)
                If (gp1 >= godz_pocz) And gp1 < godz_konc Then
                    kontr = True
                End If
                If gk1 > godz_pocz And gp1 < godz_konc Then
                    'je¿eli audycja zaczyna siê wczeœniej i koñczy w danym czasie to te¿ import
                    kontr = True
                End If

                If kontr Then
                    a = zaladuj_audycje_raportu_poemisyjnego_digas(aud, program, data_emisji)
                End If

            Next
        End If


        If colPOzycjeRaportuPoemisyjnegoDIGAS.Count > 0 Then
            'tu moga znaleŸæ sie pozycje które rozpoczê³y sie przed pocz¹tkiem emisji wniosku albo po zakoñczeniu okresu wniosku dlatego trzeba je wyrzuciæ
            For Each poz In colPOzycjeRaportuPoemisyjnegoDIGAS
                kontr = False
                gp1 = CDate(poz.time_start)
                If gp1 < godz_pocz Then
                    kontr = True 'trzeba wyrzuciæ
                End If

                If gp1 > godz_konc Then
                    kontr = True 'trzeba wyrzuciæ
                End If

                If kontr Then
                    colPOzycjeRaportuPoemisyjnegoDIGAS.Remove(poz.xml_id & "_" & poz.db_ref & "_" & poz.time_start)
                End If

            Next


        End If


        If colPOzycjeRaportuPoemisyjnegoDIGAS.Count > 0 Then
            'tu mo¿e dziwne ale chodzi o to ¿eby uporz¹dkowac klucze elementów kolekcji

            For Each poz In colPOzycjeRaportuPoemisyjnegoDIGAS
                colTymczasowaListaEmisyjna.Add(poz, poz.xml_id & "_" & poz.db_ref & "_" & poz.time_start)
            Next
            Do
                colPOzycjeRaportuPoemisyjnegoDIGAS.Remove(1)
            Loop While colPOzycjeRaportuPoemisyjnegoDIGAS.Count > 0
            k = 0
            For Each poz In colTymczasowaListaEmisyjna
                k += 1
                poz.zapisac = False
                colPOzycjeRaportuPoemisyjnegoDIGAS.Add(poz, "_" & k)
            Next
        End If



        'teraz korekta d³ugoœci muzyki, jingli i reklam (jingiel koñcowy reklam jest klasy commercial)
        Dim akt_poz As clsPozycjaRaportuDigas
        Dim nast_poz As clsPozycjaRaportuDigas
        Dim k_akt As Integer = 0
        Dim k_nast As Integer = 0

        k_akt = 1
        k_nast = 2

        If colPOzycjeRaportuPoemisyjnegoDIGAS.Count > 1 Then
            Do
                akt_poz = colPOzycjeRaportuPoemisyjnegoDIGAS.Item("_" & k_akt)
                nast_poz = colPOzycjeRaportuPoemisyjnegoDIGAS.Item("_" & k_nast)

                If akt_poz.time_stop > nast_poz.time_start Then
                    a = skoryguj_dlugosc_elementu(akt_poz, nast_poz)
                    If a = 1 Then
                        'element nastêpny zakoñczy³³ siê zanim skoñczy³ siê element aktualny
                        'w takiej sytuacji tylko zmiana elementu nastêpnego
                        k_nast += 1
                    Else
                        'element nastêpny zakoñczy³ siê po zakoñczeniu elementu aktualnego
                        k_akt = k_nast
                        k_nast += 1
                    End If
                Else
                    k_akt = k_nast
                    k_nast += 1
                End If
            Loop While k_nast < (colPOzycjeRaportuPoemisyjnegoDIGAS.Count + 1)
        End If

        Dim dl_ms As Integer

        a = grupuj_pozycje_raportu_digas(godzina_rozpoczecia, dl_ms)

        Return dl_ms

    End Function


    Public Function grupuj_pozycje_raportu_digas(ByVal godzina_rozpoczecia As String, ByRef laczna_dlugosc As Integer) As Integer
        'laczna_dlugosc - zwracane w milisekundach

        Dim tyt As String
        Dim dl_ms As Integer
        Dim tmp_poz As clsPozycjaRaportuDigas
        Dim k As Integer = 0
        Dim p_u As clsKlasaImportowanychNagranDigas
        Dim poz As clsPozycjaRaportuDigas

        Try
            If colRaportDigAS.Count > 0 Then
                Do
                    colRaportDigAS.Remove(1)
                Loop While colRaportDigAS.Count > 0
            End If

        Catch ex As Exception

        End Try

        'teraz wybieranie i grupowanie
        k = 0
        If colPOzycjeRaportuPoemisyjnegoDIGAS.Count > 0 Then
            For Each p_u In colPozycjeUstawienImportu
                If p_u.tryb_importu = 1 Then
                    '                                   0 - pomijanie tej klasy elementów
                    '                                   1 - import bespoœredni elementu - ka¿de nagranie tej klasy zapisane z tytu³em i d³ugoœci¹ jako oddzielny element we wniosku
                    '                                   2 - grupowanie elementow klasy - sumowanie czasów wszystkicjh elementów klasy i zapis jednej pozycji ze zsumowanym czasem we wniosku

                    For Each poz In colPOzycjeRaportuPoemisyjnegoDIGAS
                        If UCase(p_u.klasa) = UCase(poz.klasa) Then
                            'tu zmian z dnia 4 listopada 2014
                            'program importuje rozdzaj i podrodzaj z po³ News_Ressort i News_SubRessort
                            'ale tylko dla pozycji importownych indywidunie
                            'dla grup  - rodzaj i podrodzaj pobierany jest szablonu importowego zdefiniowngo w ustawieniach importu DigAS
                            '                            poz.rodzaj = p_u.rodzaj
                            '                           poz.podrodzaj = p_u.podrodzaj
                            poz.rodzaj_audycji = p_u.rodzaj_audycji_skrot
                            k += 1
                            poz.id_poz = k
                            poz.zapisac = True
                            colRaportDigAS.Add(poz, poz.xml_id & "_" & poz.db_ref & "_" & poz.time_start)
                        End If
                    Next poz
                ElseIf p_u.tryb_importu = 2 Then
                    tyt = p_u.nazwa_grupy
                    dl_ms = 0
                    For Each poz In colPOzycjeRaportuPoemisyjnegoDIGAS
                        If UCase(p_u.klasa) = UCase(poz.klasa) Then
                            If poz.grupuj Then
                                dl_ms = dl_ms + poz.skorygowany_duration
                            End If
                        End If
                    Next poz
                    If dl_ms > 0 Then
                        tmp_poz = New clsPozycjaRaportuDigas
                        tmp_poz.time_start = godzina_rozpoczecia
                        tmp_poz.tytul = tyt
                        tmp_poz.duration = dl_ms
                        tmp_poz.klasa = p_u.klasa
                        tmp_poz.zgrupowany = True
                        tmp_poz.rodzaj = p_u.rodzaj
                        tmp_poz.podrodzaj = p_u.podrodzaj
                        tmp_poz.rodzaj_audycji = p_u.rodzaj_audycji_skrot
                        k += 1
                        tmp_poz.id_poz = k
                        colRaportDigAS.Add(tmp_poz, tmp_poz.klasa)
                    End If
                End If
            Next p_u
        End If

        dl_ms = 0
        For Each poz In colRaportDigAS
            dl_ms += poz.skorygowany_duration

        Next

        '        tmp_poz = New clsPozycjaRaportuDigas
        '       tmp_poz.tytul = "R A Z E M"
        '      tmp_poz.duration = dl_ms
        '     tmp_poz.skorygowany_duration = dl_ms
        '    tmp_poz.id_poz = 0 '¿ebyu nie zapisywaæ

        '   colRaportDigAS.Add(tmp_poz, "razem")

        laczna_dlugosc = dl_ms


        If colRaportDigAS.Count > 0 Then
            For Each poz In colRaportDigAS
                poz.zapisac = False
                If poz.zgrupowany Then
                    poz.plik_audio = ""
                End If
            Next
        End If

    End Function
    Private Function skoryguj_dlugosc_elementu(ByRef akt_poz As clsPozycjaRaportuDigas, ByRef nast_poz As clsPozycjaRaportuDigas) As Integer
        Dim czas_startu2 As Date
        Dim czas_konca1 As Date
        Dim roznica_casu As TimeSpan
        Dim kontr As Boolean = False
        Dim wynik As Integer = 0



        If UCase(akt_poz.klasa) = "CART" Then
            kontr = True
        ElseIf UCase(akt_poz.klasa) = "MUSIC" Then
            kontr = True
        ElseIf UCase(akt_poz.klasa) = "COMMERCIAL" Then
            kontr = True
        End If

        'tu mog¹ siê zdarzyæ dwie sytuacje
        'element nastêpny zaczyna siê na koñcówce aktualnego i czas zakoñczenia nastepnego jest wiekszy ni¿ czas zakoñczenia poprzedniego
        'tzn zwy³y mix
        'albo
        'czas zakoñczenia nastêpnego jest mniejszy niz czas zakoñczenia aktualnego
        'tzn np wejœcie s³owne na muzyce
        'w tej sytuacji funkcja zwraca 1 co oznacza ¿e by element nastêpny pomin¹æ przy analizowaniu kolejnych elementów listy

        If kontr Then
            If akt_poz.time_stop > nast_poz.time_stop Then
                'wejœcie na muzyce 
                akt_poz.skorygowany_duration = akt_poz.skorygowany_duration - nast_poz.skorygowany_duration
                If akt_poz.skorygowany_duration < 0 Then
                    akt_poz.skorygowany_duration = 0
                End If
                wynik = 1
            Else
                'mix
                czas_startu2 = CDate(nast_poz.time_start)
                czas_konca1 = CDate(akt_poz.time_stop)
                roznica_casu = czas_konca1 - czas_startu2
                akt_poz.skorygowany_duration = akt_poz.skorygowany_duration - (roznica_casu.Seconds) * 1000
                If akt_poz.skorygowany_duration < 0 Then
                    akt_poz.skorygowany_duration = 0
                End If
            End If
        End If

        If kontr = False Then
            'akt element inny niz muzyka jingiel lub reklama
            'w takiej sytuacji je¿eli aktualny element to AUDIO, MAGAZINE lub LIVE
            'sprawdzenie nastêpnego elementu, je¿eli jest to CART albo MUSIC albo COMMERCIAL to trzeba skróciæ nastêny element na pocz¹tku
            ' to jest nastêpuj¹ca sytuacja
            's³owo ze studia i przed koñcem wypowiedzi uruchomiono emisj¹ nastêpnego elementu
            kontr = False
            If UCase(akt_poz.klasa) = "MAGAZINE" Then
                kontr = True
            ElseIf UCase(akt_poz.klasa) = "AUDIO" Then
                kontr = True
            ElseIf UCase(akt_poz.klasa) = "LINE" Then ' wg ustaleñ mia³a to byc klasa LIVE ale Witek ustawi³ to w DigASie na LINE
                'to jest weœcie mirofonowe prowadz¹cgo
                kontr = True
            End If
            If kontr = True Then
                kontr = False
                If UCase(nast_poz.klasa) = "CART" Then
                    kontr = True
                ElseIf UCase(nast_poz.klasa) = "MUSIC" Then
                    kontr = True
                ElseIf UCase(nast_poz.klasa) = "COMMERCIAL" Then
                    kontr = True
                End If
                If kontr = True Then
                    'tu jest sytuacja gdy aktualny element to s³owo
                    'a nastepny element to jingiel albo muzyka albo reklama
                    '- mix s³owa z muzyk¹ lub jinglem
                    If akt_poz.time_stop > nast_poz.time_start Then
                        ' tu sa dwie sytuacje 
                        'nastêpny element koñczy sie póŸniej niz aktualny element - zwyk³y mix - trzeba skróciæ muzyke lub jingiel
                        'albo nastêpny element koñczy siê wczesniej ni¿ aktualny element - jiniel - trzeba wyzerowac jego czas 
                        If akt_poz.time_stop > nast_poz.time_stop Then
                            nast_poz.skorygowany_duration = 0
                            wynik = 1 'zmieñ nastêpny element o 1
                            'aktualny pozostanie ten sam
                        Else
                            'mix
                            czas_startu2 = CDate(nast_poz.time_start)
                            czas_konca1 = CDate(akt_poz.time_stop)
                            roznica_casu = czas_konca1 - czas_startu2
                            nast_poz.skorygowany_duration = nast_poz.skorygowany_duration - (roznica_casu.Seconds) * 1000
                        End If
                        If nast_poz.skorygowany_duration < 0 Then
                            nast_poz.skorygowany_duration = 0
                        End If
                    End If
                End If
            End If
        End If



        Return wynik

    End Function

End Module
