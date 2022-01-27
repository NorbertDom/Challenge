#Projekt RPA
Projekt ten został stworzony na potrzeby zadania rekrutacyjnego. 
Wytyczne:
1. Robot pobiera plik Excel i uzupełnienia formularz na stronie: www.rpachallenge.com
2. Dodatkowo robot uzupełnia imieniem i nazwiskiem: www.roboform.com/filling-test-all-fields
3. Robot wysyła na maila raport ze zrealizowanych zadań

# Technologia
Program napisany został w języku Python z wykorzystaniem standardowych bibliotek.
- Do automatyzacji zadań w przeglądarce Chrome wykorzystana została zewnętrzna biblioteka 'Playwright' [https://playwright.dev/python/].
- Obsługa pliku wejściowego w formacie xlsx realizowana jest z wykorzystaniem dodatkowej biblioteki 'Pandas' [https://pandas.pydata.org/].


#Konfiguracja
Do użycia programu wymagany jest Python w wersji 3+

    git clone 
    cd Challenge && pip install -r requirements.txt
    python main.py

Wymagana jest konfiguracja pliku conf.ini.\
Sekcja pierwsza przymuje adresy stron. Adres strony z wymagania pierwszego, kolejna zmienna z wymagania drugiego a ostatnia zmienna to link do pliku źródłowego w formacie xlsx dostępnego online. 

                [website_conf]
                path_website_rpa = https://www.rpachallenge.com/
                path_website_robo = https://www.roboform.com/filling-test-all-fields
                path_to_csv_source = https://www.rpachallenge.com/assets/downloadFiles/challenge.xlsx
Druga część pliku odpowiada za konfigurację powiadomień mailowych.

                [mail_conf]
                gm_user = robotnorbert@gmail.com        #Konto skonfigurowane do wysyłania maili
                gm_cred =                               #hasło do konta zakodowane w base64
                gm_receiver =                           #odbiorca powiadomień
                gm_body_path = report.log               #plik z którego zawartość chcemy wysłać, domyślnie raport generowany jest w pliku 'raport.log'
                gm_smtp_server = smtp.gmail.com         #adres servera smtp
                gm_smtp_port = 465                      #port servera smtp