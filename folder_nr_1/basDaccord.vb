Imports System.Windows.Forms

Imports System.Runtime.InteropServices
Imports System.Diagnostics

Module basDaccord

    Public colWszystkiepodrodzaje As New Collection ' to jest potrzebne tylko przy imporcie danych z daccord
    'pierwsze 5 znaków z pola Program Area ma siepokrywaæ z piêcoma znakami z podrodzaju
    'na tej podstawie zostanie ustalony rodzaj

    Public import_listy_emisyjnej_daccord_dostepny As Boolean = False

    Public serwer_daccord_soap As String = ""

    Public Const WM_COPYDATA As Int32 = &H4A


    Public daccord_tryb_pobierania_danych As Integer = 0
    '                                                 0 - wprost z serwra SQL
    '                                                 1 - przez serwer SOAP



    Public Structure COPYDATASTRUCT
        Public dwData As Integer
        Public cbData As Integer
        Public lpData As Integer
    End Structure

    Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" ( _
            ByVal lpClassName As String, _
            ByVal lpWindowName As String) As Int32

    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                                    ByVal hwnd As Int32, _
                                    ByVal wMsg As Int32, _
                                    ByVal wParam As Int32, _
                                    ByVal lParam As Int32) As Int32


    Declare Function ChangeWindowMessageFilter Lib "user32.dll" (ByVal message As Integer, _
                                                                ByVal dwFlag As Integer) As IntPtr



    Public harp_wndh As Int32  'wnhd HARPA
    Public harp_daccord_bramka_wndh As Int32  'wndh bramki daccord

    Public program_bramka_harp_daccord As String = "hrp_da_gate.exe"

    Public bramka_daccord_odebrane_dane As String = ""
    Public bramka_daccord_odebrano_dane As Boolean = False


    Public zaimportowany_plan_emisji_daccord As New Collection
    Public skorygowany_plan_emisji_daccord As New Collection



    Public Function przeslij_dane_do_bramki_daccord(ByVal dane As String) As Integer

        Dim msg As String
        Dim wynik As Integer = 0
        Dim a As Integer

        Try
            harp_daccord_bramka_wndh = FindWindow(vbNullString, "HARP_DACCORD_LNK_" & harp_wndh)
        Catch ex As Exception

        End Try


        If harp_daccord_bramka_wndh = 0 Then
            a = uruchom_bramke_daccord()
        End If

        If a <> 0 Then
            Return -1
        End If


        Try
            Dim B() As Byte = System.Text.Encoding.UTF8.GetBytes(dane)
            'allocate memory space for byte array, and get a pointer to it
            Dim lpB As IntPtr = Marshal.AllocHGlobal(B.Length)
            'copy the byte array into memory
            Marshal.Copy(B, 0, lpB, B.Length)
            'setup a standard structure for the WM_COPYDATA message
            Dim CD As COPYDATASTRUCT
            With CD
                .dwData = 0 'can be used for custom indexing between apps
                .cbData = B.Length 'length of data
                .lpData = lpB.ToInt32 'pointer to the data
            End With
            'clean up array
            Erase B
            'allocate memory space for structure, and get a pointer to it
            Dim lpCD As IntPtr = Marshal.AllocHGlobal(Len(CD))
            'copy structure to allocated memory place
            Marshal.StructureToPtr(CD, lpCD, False)
            'send message, parameters explained:
            '(1) handle of receiving app window
            '(2) type of message, SendMessage is used for many other messages also
            '(3) handle of sending app window
            '(4) pointer to standard message structure
            wynik = SendMessage(harp_daccord_bramka_wndh, WM_COPYDATA, harp_wndh, lpCD.ToInt32)

            'free memory
            Marshal.FreeHGlobal(lpCD)

        Catch ex As Exception

            msg = "Wyst¹pi³ problem podczas komunikacji z bramk¹ d'accord: "
            msg = msg & vbCrLf & ex.Message

            MessageBox.Show(msg)

            wynik = -1

        End Try

        Return wynik

    End Function



    Public Function uruchom_bramke_daccord() As Integer

        Dim licznik As Integer = 0
        Dim wynik As Integer = 0
        Dim msg As String
        Dim plk As String = ""


        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        'import daccord idzie teraz wprost z serwera SQL
        'Bramka SOAP nie jest potrzebna

        'Return 0



        plk = Application.StartupPath
        If Microsoft.VisualBasic.Right(plk, 1) <> "\" Then
            plk = plk & "\"
        End If

        '        plk = plk & "hrp_da_gate.exe"
        plk = plk & "harp_daccord_lnk.exe"
        program_bramka_harp_daccord = plk


        'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

        Try
            '            harp_daccord_bramka_wndh = FindWindow(vbNullString, "HRP_DA_LNK_" & harp_wndh)
            harp_daccord_bramka_wndh = FindWindow(vbNullString, "HARP_DACCORD_LNK_" & harp_wndh)

        Catch ex As Exception

        End Try
        'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

        


        Try
            If harp_daccord_bramka_wndh = 0 Then 'nie ma uruchomionej bramki daccord
                'tu przekazanie parametru te¿ harp_wndh
                Process.Start(program_bramka_harp_daccord, "/" & harp_wndh.ToString)

                Do
                    System.Threading.Thread.Sleep(1000)
                    licznik += 1
                    If licznik > 30 Then
                        wynik = -1
                        Exit Do
                    End If
                    '                    harp_daccord_bramka_wndh = FindWindow(vbNullString, "HRP_DA_LNK_" & harp_wndh)
                    harp_daccord_bramka_wndh = FindWindow(vbNullString, "HARP_DACCORD_LNK_" & harp_wndh)

                Loop While harp_daccord_bramka_wndh = 0
            End If

        Catch ex As Exception
            msg = "Wyst¹pi³ problem podczas inicjalizacji bramki d'accord "
            msg = msg & vbCrLf & ex.Message
            MessageBox.Show(msg)
            wynik = -1
        End Try


        If harp_daccord_bramka_wndh <> 0 Then
            'za 30 razem mog³o sie udaæ
            wynik = 0
        End If

        Return wynik


    End Function



    Public Function odblokuj_funkcje_sendmessage() As Integer
        Dim wynik As Integer


        'zablokowane w wersji 3.0
        'HARP 3.0 pobiera dane bezpoœrednio z serwera SQL
        Return 0


        wynik = ChangeWindowMessageFilter(WM_COPYDATA, 1)
        '                                               MSGFLT_ADD is 1, MSGFLT_REMOVE is 2

        '        MessageBox.Show(wynik)

        'zalecane wg microsofta jest uzycie
        'BOOL WINAPI ChangeWindowMessageFilterEx(
        ' __in         HWND hWnd,
        '__in         UINT message,
        '__in         DWORD action,
        '__inout_opt  PCHANGEFILTERSTRUCT pChangeFilterStruct





    End Function


    Public Function dekoduj_dane_z_bramki_daccord_stare(ByVal odebrane_dane As String) As Integer


        Dim poz As clsPozycjaPlanuEmisjiDaccord
        Dim k As Integer
        Dim rekordy As String()
        Dim rekord As String
        Dim tmp_rek As String
        Dim tmp_str As String
        Dim i As Integer




        Try
            zaimportowany_plan_emisji_daccord.Clear()

        Catch ex As Exception

        End Try


        'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        'Tu jest fragment kodu przygotowyj¹cy dane w Bramce Daccord_HARP

        'tmp_str = tmp_str & n.id & "<EOF>"
        'tmp_str = tmp_str & n.godzina_emisji & "<EOF>"
        'If Len(Trim(n.tytul)) > 0 Then
        '   tmp_str = tmp_str & Trim(n.tytul) & "<EOF>"
        ' Else
        '   tmp_str = tmp_str & " <EOF>" 'spacja
        ' End If

        ' If Len(Trim(n.Autor)) > 0 Then
        '   tmp_str = tmp_str & Trim(n.Autor) & "<EOF>"
        ' Else
        '    tmp_str = tmp_str & " <EOF>" 'spacja
        ' End If'

        'tmp_str = tmp_str & n.dlugosc & "<EOF>"

        'If Len(Trim(n.program_area)) > 0 Then
        '   tmp_str = tmp_str & Trim(n.program_area) & "<EOF>"
        'Else
        '   tmp_str = tmp_str & " <EOF>" 'spacja
        'End If

        'If Len(Trim(n.info_text)) > 0 Then
        '   tmp_str = tmp_str & Trim(n.info_text) & "<EOR>" & vbCrLf
        'Else
        '   tmp_str = tmp_str & " <EOR>" & vbCrLf 'spacja
        'End If

        Try
            rekordy = Split(odebrane_dane, "<EOR>")

        Catch ex As Exception

        End Try

        Try

            For Each rekord In rekordy
                If Len(rekord) > 0 Then
                    poz = New clsPozycjaPlanuEmisjiDaccord

                    i = InStr(rekord, "<EOF>")
                    If i > 0 Then
                        tmp_str = Left(rekord, i)
                        poz.daccord_db_id = Val(tmp_str)
                    End If
                    tmp_rek = Right(rekord, Len(rekord) - i - 4)

                    i = InStr(tmp_rek, "<EOF>")
                    If i > 0 Then
                        tmp_str = Left(tmp_rek, i - 1)
                        poz.godzina_emisji = tmp_str
                    End If
                    tmp_rek = Right(tmp_rek, Len(tmp_rek) - i - 4)

                    i = InStr(tmp_rek, "<EOF>")
                    If i > 0 Then
                        tmp_str = Left(tmp_rek, i - 1)
                        poz.tytul = tmp_str
                    End If
                    tmp_rek = Right(tmp_rek, Len(tmp_rek) - i - 4)

                    i = InStr(tmp_rek, "<EOF>")
                    If i > 0 Then
                        tmp_str = Left(tmp_rek, i - 1)
                        poz.autor = tmp_str
                    End If
                    tmp_rek = Right(tmp_rek, Len(tmp_rek) - i - 4)

                    i = InStr(tmp_rek, "<EOF>")
                    If i > 0 Then
                        tmp_str = Left(tmp_rek, i - 1)
                        poz.info_text = tmp_str
                    End If
                    tmp_rek = Right(tmp_rek, Len(tmp_rek) - i - 4)

                    i = InStr(tmp_rek, "<EOF>")
                    If i > 0 Then
                        tmp_str = Left(tmp_rek, i - 1)
                        poz.program_area = tmp_str
                    End If

                    tmp_rek = Right(tmp_rek, Len(tmp_rek) - i - 4)

                    Try
                        i = InStr(tmp_rek, "<EOR>")
                        If i > 0 Then
                            tmp_str = Left(tmp_rek, i - 1)
                            poz.info_text = tmp_str
                        End If

                    Catch ex13 As Exception
                        poz.info_text = ""
                    End Try


                    k += 1
                    zaimportowany_plan_emisji_daccord.Add(poz, k & "_")


                End If
            Next


        Catch ex As Exception

        End Try


    End Function


    Public Function parsuj_dane_daccord_harp(ByVal odebrane_dane As String)

        Dim poz As clsPozycjaPlanuEmisjiDaccord
        Dim k As Integer
        Dim rekordy As String()
        Dim rekord As String
        Dim tmp_rek As String
        Dim tmp_str As String
        Dim i As Integer


        Dim a As Integer



        Try
            zaimportowany_plan_emisji_daccord.Clear()

        Catch ex As Exception

        End Try


        'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        'Tu jest fragment kodu przygotowyj¹cy dane do wysy³ki w bramce daccord 
        '
        '        tmp_str = tmp_str & "<NAGRANIE>" & vbCrLf
        '        tmp_str = tmp_str & "  <ID>" & n.id & "</ID>" & vbCrLf
        '        tmp_str = tmp_str & "  <GODZINA_EMISJI>" & n.godzina_emisji & "</GODZINA_EMISJI>" & vbCrLf
        '        If Len(Trim(n.tytul)) > 0 Then
        '            tmp_str = tmp_str & "  <TYTUL>" & n.tytul & "</TYTUL>" & vbCrLf
        '        Else
        '           tmp_str = tmp_str & "  <TYTUL>-</TYTUL>" & vbCrLf
        '        End If

        '        If Len(Trim(n.Autor)) > 0 Then
        '           tmp_str = tmp_str & "  <AUTOR>" & n.Autor & "</AUTOR>" & vbCrLf
        '        Else
        '           tmp_str = tmp_str & "  <AUTOR>-</AUTOR>" & vbCrLf 'spacja
        '        End If

        '        tmp_str = tmp_str & "  <DLUGOSC>" & n.dlugosc & "</DLUGOSC>" & vbCrLf 'liczba sekund

        '       If Len(Trim(n.program_area)) > 0 Then
        '           tmp_str = tmp_str & "  <PROGRAM_AREA>" & n.program_area & "</PROGRAM_AREA>" & vbCrLf
        '       Else
        '           tmp_str = tmp_str & "  <PROGRAM_AREA>-</PROGRAM_AREA>" & vbCrLf 'spacja
        '       End If

        '       If Len(Trim(n.info_text)) > 0 Then
        '           tmp_str = tmp_str & "  <INFO_TEXT>" & n.info_text & "</INFO_TEXT>" & vbCrLf
        '       Else
        '           tmp_str = tmp_str & "  <INFO_TEXT>-</INFO_TEXT>" & vbCrLf 'spacja
        '       End If
        '       tmp_str = tmp_str & "</NAGRANIE>" & vbCrLf
        'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



        Try
            rekordy = Split(odebrane_dane, "</NAGRANIE>")

        Catch ex As Exception

        End Try

        Try
            For Each rekord In rekordy
                If Len(Trim(rekord)) > 2 Then
                    poz = New clsPozycjaPlanuEmisjiDaccord
                    a = odczytaj_rekord_z_bramki_daccord_harp(rekord, poz)
                    If a = 0 Then
                        k += 1
                        poz.zapisac = False
                        zaimportowany_plan_emisji_daccord.Add(poz, k & "_")


                    End If

                End If

            Next

        Catch ex As Exception

        End Try





    End Function


    Private Function odczytaj_rekord_z_bramki_daccord_harp(ByVal rekord As String, ByRef poz As clsPozycjaPlanuEmisjiDaccord) As Integer


        poz.daccord_db_id = podaj_wartosc_pola_daccord_harp("ID", rekord)

        poz.godzina_emisji = podaj_wartosc_pola_daccord_harp("GODZINA_EMISJI", rekord)

        poz.tytul = podaj_wartosc_pola_daccord_harp("TYTUL", rekord)

        poz.autor = podaj_wartosc_pola_daccord_harp("AUTOR", rekord)
        poz.dlugosc = 0

        Try
            poz.dlugosc = Val(podaj_wartosc_pola_daccord_harp("DLUGOSC", rekord))
        Catch ex As Exception

        End Try

        poz.program_area = podaj_wartosc_pola_daccord_harp("PROGRAM_AREA", rekord)

        poz.info_text = podaj_wartosc_pola_daccord_harp("INFO_TEXT", rekord)

        poz.daccord_db_id = podaj_wartosc_pola_daccord_harp("ID", rekord)

        Return 0


    End Function


    Private Function podaj_wartosc_pola_daccord_harp(ByVal nazwa_pola As String, ByVal pelny_rekord As String) As String
        Dim wynik As String = ""

        Dim tmp_rek As String

        Dim poz As Integer = 0

        poz = InStr(pelny_rekord, "<" & nazwa_pola & ">")
        If poz > 0 Then
            tmp_rek = Right(pelny_rekord, Len(pelny_rekord) - poz - Len(nazwa_pola) - 1)
            poz = InStr(tmp_rek, "</" & nazwa_pola & ">")
            If poz > 0 Then
                wynik = Left(tmp_rek, poz - 1)
            End If
        End If

        Return wynik


    End Function

End Module
