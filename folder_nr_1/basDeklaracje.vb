Module basDeklaracje


    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)


    Public stawka_prac_czas_prem_taka_jak_pracownikow As Boolean = False
    'w listopadzie 2012 Radio Lublin zgłosiło potrzebę wyceniania pracowników czasowo premiowych tak samo jak pracowników
    'ta zmienna ustawiana jest w ustawieniach programu, jeżeli jest TRUE to program dla pracownika czasowo-premiowego daje stawkę pracownika
    'jeżeli jest FALSE to jako stawke dla pracowników czasowo-premiowych daje stawkę współpracownika (to znaczy działa tak jak dotychczas

    Public ukryte_kwoty_przed_zatwierdzeniem_przez_zarzad As Boolean = True


    Public spec_id_otwartego_wniosku As Integer = 0
    'w tej zmiennej zapisywany jest identyfikator wniosku właśnie otwartego z prawem do zapisu odczytu
    'służy do zapisania informacji o zamknięciu wniosku podczas specjalnego zamykania - 
    'to specjalne zamuykanie programu to np wylogowywanie użytkownika w sytuacji gdy był otwarty wniosek z prawem do zapisu
    'każde "normalne" zamknięcie wniosku powoduje wyzerowanie tej zmiennej


    Public colTabelaWycenCalejRamowki As New Collection

    Public rejestrowanie_zmian_pozycji_wnioskow_dostepne As Boolean = False
    Public przegladanie_zmian_pozycji_wnioskow_dostepne As Boolean = False
    Public rejestracja_kometarzy_pozycji_wniosku As Boolean = False
    'ww zmienne pozwalają na rejestrację wszystkich operacji w logu pozycji wniosków

    Public blokada_przegladania_zestawien_dla_kierownikow As Boolean = False
    Public blokada_regulacji_wycen_dla_autorow As Boolean = False
    '' obie ww zmienne wczytywane jako ustawienia programu
    ' dorobiono w lipcu 2008 na życzenie Radia Olsztyn

    Public blokada_wyceny_indywidualnej_dla_autorow As Boolean = False
    'zmienna dodana w dniu 16 stycznia 2017
    'na potrzeby Radia Łódź


    Public blokada_otwarcia_po_zatwierdzeniu As Boolean = False
    'powyższa zmianna dorobiona dla RDC - blokada otwarcia po zatwierdzeniu przez osobe o wyzszych uprawnieniach


    Public wymuszanie_zmiany_hasla_dostepne As Boolean = False

    Public otwarcie_inspektora_ukryte_kwoty As Boolean

    Public rozszerzone_sprawozdania_dostepne As Boolean = False ' ta zmienna załącza pole info we wniosku 
    '                                                           i tabele stałych pozycji audycji a co za tym drugim idzie 
    '                                                           wystawianie wniosków z dodawaniem stałych pozycji
    '                                                           dodatkowo rozliczanie muzyki i kontrola 
    '                                                           dokumentacji działa tylko wteduy gdy 
    '                                                           załączone są rozszrzone sprawozdania
    Public wstepne_zatwierdzanie_dostepne As Boolean = False
    Public rozliczanie_muzyki_dostepne As Boolean = False
    Public kontrola_dokumentacji_dostepne As Boolean = False

    Public oznaczanie_pozycji_niezaliczonych_do_sredniej_dostepne As Boolean = False
    'ta zmianna wykorzystywana jest tylko w oknie zestawienia wycen pracownika
    'służy do tego żeby pracownik działu kadr-płac mógł oznaczyć wyceny honoracyjne któe nie będą zaliczone do średniej
    'na dzień 24 maja 2008 wykorzystywane to jest tylko w Radiu Lublin przy eksporcie danych do programu płacowego


    ' Public info_wniosku_dostepne As Boolean = False
    Public wersja_demo As Boolean = True

    Public okres_rozliczeniowy As Integer = 0 ' 0 -  1 do ostatniego w tym samym miesiącu
    '                                           1 - dowolnie wybieralny okres
    Public licencjobiorca As String = "Wersja demonstracyjna programu. "

    Public Const cTabelaOgolna As Integer = 0
    Public Const cTabelaAudycji As Integer = 1
    Public Const cTabelaPrywatna As Integer = 2
    Public Const cWycenaIndywidualna As Integer = 3

    Public colStalePozycjeAktualnejAudycji As New Collection


    Public colPasma As New Collection

    Public colRodzajeProdukcji As New Collection
    Public colRodzajeAudycji As New Collection

    Public colRodzajeLicencji As New Collection
    Public domyslny_rodzaj_licencji As String = "LAN"

    Public colRodzajeDokumentacji As New Collection
    Public domyslny_rodzaj_dokumentacji As String = "-"

    Public colRodzajeMuzyki As New Collection
    Public domyslny_rodzaj_muzyki As String = "MU"

    Public domyslny_rodzaj_produkcji As String = "KW"
    Public domyslny_rodzaj_audycji As String = "SL"



    Public colZestawienieDLugosciMuzyki As New Collection
    Public colZestawienieDLugosciMuzyki_wgMPK As New Collection


    Public prywatne_tabele_wycen_dostepne As Boolean = True
    Public colPrywatnaTabelaWycen As New Collection

    Public trwa_logowanie As Boolean = False

    Public rozliczanie_kosztow_niehonoracyjnych As Boolean = True

    Public okno_glowne_programu As MainForm
    Public nazwa_stacji_komputerowej As String = "-"

    Public colProgramy As New Collection

    Public colPracownicy As New Collection
    Public colWyswietleniPracownicy As New Collection
    Public colUsunieciPracownicy As New Collection

    Public colUmowyPracownika As New Collection
    Public colNIeobecnosciPracownika As New Collection

    Public colListaPlac As New Collection

    Public colRedakcje As New Collection

    Public termin_wystawiania_wnioskow As Integer = 3 ' liczba dni po któych program blokuje możliwość wystawienia wnioskó
    ' ww blokada nie obowiązuje dla Zarządu i Administratora
    Public wyprzedzenie_terminu_wystawienia_wniosku As Integer = 30 ' liczba dni - wyprzedzenie 
    '                                                                       z jakim można wystawiać wnioski - ta blokada dotyczy wszystkich 

    Public tryb_obslugi_kosztow_uzysku As Integer = 0
    '0 - wszystkie pozycje z kosztem 50%
    '1 - rozróżnianie kosztów 20% i 50%
    '2 - rozróżńianie kosztów 0%, 20% i 50%

    Public colUprawnieniaWRedakcjach As New Collection ' aktualnie zalogowanego pracownika
    Public colUprawnieniaWRedWybranegoPracownika As New Collection ' aktualnie edytowanego pracownika

    Public colUprawnieniaWProgramach As New Collection ' aktualnie zalogowanego pracownika
    Public colUprawnieniaProgrWybranego_pracownika As New Collection ' aktualnie edytowanego pracownika



    Public colKodyMPK As New Collection
    Public colZadania As New Collection

    Public colRodzajeProgramowe As New Collection



    Public colTabelaWycenAktualnegoWNiosku As New Collection 'w tej kolekcji są pozycje tabeli ogólnej 
    '                                                           i pozycje tabeli przypisane do audycji ramówkowej
    Public akt_wybrana_pozycja_tabeli_wycen As clsPozycjaTabeliWycen

    Public colOgolnaTabelaWycen As New Collection ' tu jest ogólna tabela wycen


    Public colZestawienieKOsztowRedakcji As New Collection
    Public colZestawienieKosztowAudycji As New Collection

    Public colZestawienieKOsztowZadan As New Collection

    Public colZestawienieKosztowWgZrodlaFinansowania As New Collection

    Public colZrodlaFinansowania As New Collection
    Public rozroznianie_zrodel_finansowania As Boolean = True



    Public liczba_szczebli_zatwierdzania As Integer = 4
    'możliwe wartości 3 i 4 , 4 gdy zatwierdzanie przaz szefa programu

    Public naglowek_komunikatow As String = "HARP"

    Public katalog_ustawien_wyswietlania As String = ""
    Public katalog_config As String = ""

    Public rozliczanie_minutowe_audycji_dostepne As Boolean = False


    Public sprawozdania_programowe_dostepne As Boolean = True

    Public zadaniowanie_dostepne As Boolean = True

    Public indywidualne_wsp_wyceny_dostepne As Boolean = True

    Public MPK_dostepne As Boolean = True

    Public tryb_obslugi_MPK As Integer = 0
    '                                   0- MPK przypisane tylko do naglówka wniosku
    '                                   1- MPK przypisane do każdej pozycji wniosku 


    Public tryb_regulacji_wsp_wycen As Integer = 0
    '0- wspólne widełki dla wszystkuch wycen
    '1 - zakres regulacji ustawiany oddzielnie dla każdej wyceny w tanbeli wycen


    Public gorna_granica_wspolczynnika As Single = 0.5
    Public dolna_granica_wspolczynnika As Single = 1.5

    Public krok_regulacji_wyceny As Double = 0.05



    'poniższe stałe dotyczą sposobu zaokrąglania
    'wycen
    Public Const ZAOKKRAGLANIE_W_GORE As Integer = 1
    Public Const ZAOKKRAGLANIE_ARYTMETYCZNE As Integer = 2
    Public Const DOKLADNOSC_PELNY_ZLOTY As Integer = 1
    Public Const DOKLADNOSC_DZIESIATKI_GROSZY As Integer = 2
    Public Const DOKLADNOSC_PELNY_GROSZ As Integer = 3

    'poniższe dwie zmienne dotyczą sposobu przeliczania i zaokrąglania
    'wycen we wnioskach
    Public dokladnosc_zapisu_wycen As Integer = 2
    '                                           1-pełne złote
    '                                           2-dziesiątki groszy
    '                                           3-pelne grosze
    Public sposob_zaokraglania_wycen As Integer = 2
    '                                           1- zaokrąglanie w górę
    '                                           2-zaokrąglanie arytmetyczne

    Public tryb_obslugi_Nieobecnosci As Integer = 0 '   0 blokada wystawiania
    '                                                   1 - możliwe wystawianie
    Public tryb_kontroli_budzetow_redakcji As Integer '0 - kontrola wyłączona
    '                                                 '1 kontrola wydatków finansowanych ze środków włąsnych
    '                                                 '2- kontrola wydatków na zadania o id < niż 10
    '                                                  3 - kontrola wszystkich  wydatków na zadania o id < niż 10 finansowanych ze srodków własnych
    '                                                  4 - kontrola wszystkich  


    Public ustawienia_okresow_rozliczeniowych As Integer = 0 '0 - 1 do ostatni
    '                                                      1 - 21 do 20 następnego miesiąca 


    Public tryb_kontroli_budzetu_audycji As Integer = 0 ' 0 kontrola pojedynczego wydania audycji
    '                                                   ' 1 - kontrola łącznej kwoty na wszystkie wydania audycji w miesiącu
    '               ww zmienna wczytywana z ustawień programu - opcja dostepna od 23 marca 2010 od wersji 2.0.94

    'tryb otwierania wniosku
    Public czy_blokada_otwarcia_wniosku_dla_asystenta As Boolean = True
    Public sprawozdania_programowe_dozwolona_edycja As Boolean = False
    Public dozwolone_wystawianie_swoich_pozycji As Boolean = False
    Public czy_blokada_otwarcia_wniosku_tylko_do_odczytu As Boolean = False


    Public ukryte_kwoty_wycen_dla_autorow_audycji As Boolean = False

    Public pelna_lista_honorariuw_dostepna_przez_zatwierdzeniem As Boolean = False


    Public Const C_UPRAWNIENIA_ADMINISTRATORA As Integer = 1
    Public Const C_UPRAWNIENIA_ZARZADU As Integer = 2
    Public Const C_UPRAWNIENIA_KADRY_PLACE As Integer = 4
    Public Const C_UPRAWNIENIA_UMOWY_O_DZIELO As Integer = 8
    Public Const C_UPRAWNIENIA_PELNE_SPRAWOZDANIA As Integer = 16
    Public Const C_UPRAWNIENIA_EDYCJA_SPRAAWOZDANIA_PROGRAMOWEGO As Integer = 32
    Public Const C_UPRAWNIENIA_BLOKOWANIA_WNIOSKOW As Integer = 64
    Public Const C_UPRAWNIENIA_EDYCJI_RAMOWKI As Integer = 128
    Public Const C_UPRAWNIENIA_EDYCJI_PLANU_WYDATKOW As Integer = 256
    Public Const C_UPRAWNIENIA_INSPEKTORA_PROGRAMU As Integer = 512
    Public Const C_UPRAWNIENIA_PRZEGLADANIA_OPISOW_AUDYCJI As Integer = 1024
    Public Const C_UPRAWNIENIA_KONFIGURACJA_SERWIS As Integer = 2048
    Public Const C_UPRAWNIENIA_EDYCJA_NIEOBECNOSCI As Integer = 4096


    'upoważnienia do wniosków
    Public Const C_WSPOLAUTOR As Integer = 1
    Public Const C_EDYCJA_SPRAWOZDANIA As Integer = 2
    Public Const C_EDYCJA_OPISU As Integer = 4
    Public Const C_EDYCJA_DAB As Integer = 8





    Public skopiowana_pozycja_tabeli As New clsPozycjaTabeliWycen
    Public colSkopiowanaTabelaWycen As New Collection


    Public oznaczanie_wspolpracownikow_wewnetrznych_dostepne As Boolean = False
    Public regulacja_wg_ustalonych_stawek As Boolean = False ' w Radiu Rzeszów chca regulować stawki nie o 1 zł czy określony procent
    '                                                           a do ściśle okreslonych stawek a b c d
    'do każdej pozycji w tabeli wycen mozna przypisac wartości jakie mogą być ustawiane dla określonej stawki
    'regulując wycenę (a nie współczynnik wyceny) reguluje sie przełącając pomiędzy kolejno ustalonymi stawkami


    Public oznaczanie_pracownikow_ryczaltowych_dostepne As Boolean
    'w Radiu Kraków maja pracowników i współpracowników ryczałtowych
    'podczas wystawiania wycen kwoty dla tych osób są zapisywane ale podczas transferu danych do programu płacowego ich kwoty są zerowane
    'mechanizm ten ma służyć do wyliczenia kosztów wg MPK 0 proporcjonalnego rozliczenia wg tych proporcji wynagordzeń ryczałtowców w systemie płacowym 
    ' dodatkowo ryczałtowcy przeglądając swoje zestawienia "widzą" tylko pozycje bez kwot

    Public rozliczanie_kosztow_rodzajow_programowych_dostepne As Boolean = True
    'powyższa zmienna wprowadzona w dniu 31 sierpnia 2011 wersja 2.0.104
    'gdy ta zmienna jest ustawiona to program w liscie pozycji honoracyjnych pokazuje rodzaj, podrodzaj i dlugosc



    Public okno_spisu_wnioskow As frmSpisWnioskow


    Public colListaGrupDocelowych As New Collection
    Public colRodzajeRealizacji As New Collection

    Public filtr_listy_sprawozdania As clsFiltr
    Public filtr_listy_wycen As clsFiltr

    Public oznaczanie_wnioskow_sygnatura_dostepne As Boolean = False

    Public lista_plac_wg_kosztow_uzysku As Boolean = True

    Public trwa_odswiezanie_spisu_wnioskow As Boolean = False

    Public Declare Function GetLastError Lib "kernel32" Alias "GetLastError" () As Integer

    Public kopiuj_pliki_audio_przy_imporcie_listy_emisyjnej As Boolean = False
    Public sciezka_zapisu_plikow_audio As String = ""


End Module
