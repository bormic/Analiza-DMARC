# -*- coding: utf-8 -*-
"""
Created on Thu Mar 14 20:18:09 2024

@author: bormic
"""

import zipfile
import gzip
import os
import shutil
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np
import tkinter as tk
from tkinter import filedialog

def choose_files():
    try:
        work_folder = os.path.dirname(__file__)  # Pobierz katalog, w którym znajduje się bieżący skrypt
        extracted_data_folder = os.path.join(work_folder, "extracted_data")  # Utwórz ścieżkę względną do folderu extracted_data_folder      
        bak_folder = os.path.join(work_folder, "backup_arch")  # Utwórz ścieżkę względną do folderu bakup_arch
        os.makedirs(extracted_data_folder, exist_ok=True)  # Tworzenie katalogu "extracted_data_folder" jeśli nie istnieje
        os.makedirs(bak_folder, exist_ok=True)  # Tworzenie katalogu "bak" jeśli nie istnieje
        files = filedialog.askopenfilenames(filetypes=[("All Files", "*.*"), ("Zip Files", "*.zip"), ("Gz Files", "*.gz")])
        # Inicjujemy pustą listę, w której będziemy przechowywać dane dla poszczególnych plików
        data_list = []
        # Ścieżka do pliku wynikowego
        output_file = os.path.join(work_folder, "wynik.xlsx")
        
        if files:
            for filename in files:
                print("Wybrany plik:", filename)
                file_path = os.path.join(work_folder, filename)
                if filename.endswith(".zip"):
                    with zipfile.ZipFile(filename, 'r') as zip_ref:
                        zip_ref.extractall("extracted_data")
                    # Przenieś archiwum do katalogu "bak"
                    shutil.move(file_path, os.path.join(bak_folder, os.path.basename(filename)))
                elif filename.endswith(".gz"):
                    # Rozpakuj plik gzip do folderu extracted_data
                    with gzip.open(filename, 'rb') as gz_file:
                        extracted_filename = os.path.basename(filename)[:-3]
                        extracted_file_path = os.path.join(extracted_data_folder, extracted_filename)
                        with open(extracted_file_path, 'wb') as xml_file:
                            xml_file.write(gz_file.read())
                    # Przenieś archiwum do katalogu "bak"
                    shutil.move(filename, os.path.join(bak_folder, os.path.basename(filename)))
        # Przeszukaj rozpakowane pliki i analizuj pliki XML poza pętlą wybierania plików
        for filename in os.listdir(extracted_data_folder):
            if filename.endswith(".xml"):
                xml_file_path = os.path.join(extracted_data_folder, filename)
                data = extract_data_from_xml(xml_file_path)
                if data is not None:
                    data_list.append(data)
                    #print("data list to: ", data_list)        
        
        # Tworzymy ramkę danych pandas
        df_f = process_dataframe(data_list)

        # Sprawdź, czy istnieje plik Excela
        if os.path.exists(output_file):
            # Wczytaj istniejący plik Excela
            existing_df = pd.read_excel(output_file)
            # Połącz nowe dane z istniejącymi danymi
            df_f = pd.concat([existing_df, df_f], ignore_index=True)
        else:
            # Zapisz dane do pliku Excela
            df_f.to_excel(output_file, index=False)

        # Po zakończeniu przetwarzania usuń rozpakowane pliki
        for filename in os.listdir(extracted_data_folder):
            print("file name to: ", filename)
            file_path = os.path.join(extracted_data_folder, filename)
            os.remove(file_path)

        print("Ekstrakcja i analiza zakończona pomyślnie!")

    except zipfile.BadZipFile as e:
        print(f"Błąd: Nieprawidłowy format pliku ZIP - {e}")
    except Exception as e:
        print(f"Błąd podczas analizy pliku: {e}")

def process_dataframe(data_list):
    df = pd.DataFrame(data_list)
    
    # wyjcie df
    #pd.set_option('display.max_rows', None)  # Wyświetl wszystkie wiersze
    #pd.set_option('display.max_columns', None)  # Wyświetl wszystkie kolumny
    #display(df)
    
    # Dodaj aktualną datę i godzinę do ramki danych
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df["Aktualna data"] = now   
    # Konwertujemy czas z formatu UNIX Epoch na bardziej czytelny format
    df['begin'] = pd.to_datetime(df['begin'], unit='s')
    df['end'] = pd.to_datetime(df['end'], unit='s')

    # Tworzymy mapowanie wartości dla 'adkim'
    mapping = {'r': 'Brak DKIM', 's': 'Wyłączony DKIM', 'n': 'Włączony DKIM'}
    #df['adkim'] = df['adkim'].map(mapping)
    if 'adkim' in df.columns:
        try:
            df['adkim'] = df['adkim'].replace(mapping)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'adkim'.")
    else:
        print("Błąd: Kolumna 'adkim' nie istnieje w ramce danych.")


    # Tworzymy mapowanie wartości dla 'aspf'
    mapping_aspf = {'r': 'Brak SPF', 's': 'Wyłączony SPF', 'n': 'Włączony SPF'}
    #df['aspf'] = df['aspf'].map(mapping_aspf)
    if 'aspf' in df.columns:
        try:
            df['aspf'] = df['aspf'].replace(mapping_aspf)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'aspf'.")
    else:
        print("Błąd: Kolumna 'aspf' nie istnieje w ramce danych.")

    # Zastępujemy wartości w kolumnie 'comment'
    comment_mapping = {
        'arc=fail': 'Niepowodzenie weryfikacji ARC',
        'arc=pass': 'Powodzenie weryfikacji ARC',
        'arc=none': 'Brak wyniku weryfikacji ARC',
        'arc=invalid': 'Nieprawidłowy wynik weryfikacji ARC',
        'arc=timestamp_error': 'Błąd czasu weryfikacji ARC',
        'arc=signature_error': 'Błąd podpisu weryfikacji ARC',
        'arc=authentication_results_mismatch': 'Niezgodność wyników autentykacji weryfikacji ARC'
    }
    #df['comment'] = df['comment'].replace(comment_mapping)
    if 'comment' in df.columns:
        try:
            df['comment'] = df['comment'].replace(comment_mapping)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'comment'.")
    else:
        print("Błąd: Kolumna 'comment' nie istnieje w ramce danych.")

    # Zastępujemy wartości w kolumnie 'pct'
    pct_mapping = {'100': 'Wymagane raportowanie dla wszystkich wiadomości',
                   '50': 'Wymagane raportowanie dla 50% wiadomości',
                   '0': 'Brak wymaganego raportowania'}
    #df['pct'] = df['pct'].replace(pct_mapping)
    if 'pct' in df.columns:
        try:
            df['pct'] = df['pct'].replace(pct_mapping)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'pct'.")
    else:
        print("Błąd: Kolumna 'pct' nie istnieje w ramce danych.")

    # Zastępujemy wartości w kolumnie 'np'
    np_mapping = {'none': 'Brak polityki (brak działań)',
                  'quarantine': 'Umieszczenie w kwarantannie',
                  'reject': 'Odrzucenie wiadomości'}
    #df['np'] = df['np'].replace(np_mapping)
    if 'np' in df.columns:
        try:
            df['np'] = df['np'].replace(np_mapping)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'np'.")
    else:
        print("Błąd: Kolumna 'np' nie istnieje w ramce danych.")

    # Zastępujemy wartości w kolumnie 'disposition'
    disposition_mapping = {'none': 'Brak autentykacji',
                           'quarantine': 'Umieszczono w kwarantannie',
                           'reject': 'Odrzucono wiadomość'}
    #df['disposition'] = df['disposition'].replace(disposition_mapping)
    if 'disposition' in df.columns:
        try:
            df['disposition'] = df['disposition'].replace(disposition_mapping)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'disposition'.")
    else:
        print("Błąd: Kolumna 'disposition' nie istnieje w ramce danych.")

    # Zastępujemy wartości w kolumnie 'dkim'
    dkim_mapping = {'pass': 'Poprawna weryfikacja DKIM',
                    'fail': 'Błąd weryfikacji DKIM',
                    'none': 'Brak autentykacji DKIM'}
    #df['dkim'] = df['dkim'].replace(dkim_mapping)
    if 'dkim' in df.columns:
        try:
            df['dkim'] = df['dkim'].replace(dkim_mapping)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'dkim'.")
    else:
        print("Błąd: Kolumna 'dkim' nie istnieje w ramce danych.")

    # Zastępujemy wartości w kolumnie 'p'
    p_mapping = {'none': 'Brak polityki DMARC',
                 'quarantine': 'Umieszczono w kwarantannie',
                 'reject': 'Odrzucono wiadomość'}
    #df['p'] = df['p'].replace(p_mapping)
    if 'p' in df.columns:
        try:
            df['p'] = df['p'].replace(p_mapping)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'p'.")
    else:
        print("Błąd: Kolumna 'p' nie istnieje w ramce danych.")

    # Zastępujemy wartości w kolumnie 'sp'
    sp_mapping = {'none': 'Brak polityki dla poddomen',
                  'quarantine': 'Umieszczono w kwarantannie dla poddomen',
                  'reject': 'Odrzucono wiadomość dla poddomen'}
    #df['sp'] = df['sp'].replace(sp_mapping)
    if 'sp' in df.columns:
        try:
            df['sp'] = df['sp'].replace(sp_mapping)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'sp'.")
    else:
        print("Błąd: Kolumna 'sp' nie istnieje w ramce danych.")

    # Tworzymy słownik z tłumaczeniami
    translation = {
        "mfrom": "Zakres SPF: Od pola MAIL FROM - ewentualna zmiana nadawcy i odbiorcy może wpłynąć na wynik weryfikacji SPF",
        "helo": "Zakres SPF: Od pola HELO - ewentualna zmiana domeny nadawcy może wpłynąć na wynik weryfikacji SPF",
        "pra": "Zakres SPF: Od pola PRA - ewentualna zmiana domeny nadawcy może wpłynąć na wynik weryfikacji SPF",
        "explanation": "Zakres SPF: Dostęp do rekordu SPF jest tylko dla objaśnienia"
    }

    # Utworzenie kolumny z tłumaczeniami
    #df['scope'] = df['scope'].map(translation)
    if 'scope' in df.columns:
        try:
            df['scope'] = df['scope'].replace(translation)
        except KeyError:
            print("Błąd: Wystąpił problem podczas zastępowania wartości w kolumnie 'scope'.")
    else:
        print("Błąd: Kolumna 'scope' nie istnieje w ramce danych.")

    # Zamieniamy "\n", "None" i NaN na "Brak danych"
    df.replace({r'^\s*$': 'Brak danych', np.nan: 'Brak danych', 'None': 'Brak danych'}, inplace=True, regex=True)

    return df


def extract_data_from_xml(xml_file):
    # Sprawdź, czy plik ma rozszerzenie XML
    if xml_file.endswith('.xml'):
        # Tworzymy drzewo elementów na podstawie pliku XML
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Inicjujemy pusty słownik, w którym będziemy przechowywać dane
        data = {}

        # Funkcja pomocnicza do rekurencyjnego przeszukiwania drzewa elementów
        def extract_data_recursive(element):
            # Pobieramy nazwę elementu i jego zawartość
            element_name = element.tag
            element_content = element.text

            # Dodajemy nazwę elementu i jego zawartość do słownika
            data[element_name] = element_content

            # Przechodzimy rekurencyjnie przez wszystkie dzieci danego elementu
            for child in element:
                extract_data_recursive(child)

        # Wywołujemy funkcję pomocniczą dla głównego elementu
        extract_data_recursive(root)

        return data
    else:
        # Jeśli plik nie ma rozszerzenia XML, zwracamy None
        return None

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        generate_plots_from_excel(file_path)

def generate_plots_from_excel(excel_file):
    work_folder_w = os.path.dirname(__file__)  # Pobierz katalog, w którym znajduje się bieżący skrypt
    output_folder = os.path.join(work_folder_w, "wykresy")  # Utwórz ścieżkę względną do folderu wykresy
    os.makedirs(output_folder, exist_ok=True)  # Tworzenie katalogu "wykresy" jeśli nie istnieje
    # Wczytaj dane z pliku Excela do ramki danych pandas
    df = pd.read_excel(excel_file)

    # wyjcie df
    #pd.set_option('display.max_rows', None)  # Wyświetl wszystkie wiersze
    #pd.set_option('display.max_columns', None)  # Wyświetl wszystkie kolumny
    #display(df)
 
    # Wykres ilości wystąpień org_name
    plt.figure(figsize=(10, 6))
    df['org_name'].value_counts().plot(kind='bar', color='skyblue')
    plt.title('Ilość wystąpień org_name')
    plt.xlabel('org_name')
    plt.ylabel('Ilość wystąpień')
    plt.xticks(rotation=45)
    plt.tight_layout()
    output_path = os.path.join(output_folder, "org_name_plot.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    # Wykres porównujący pass do fail dla dkim, spf i comment (arc=fail)
    plt.figure(figsize=(10, 6))
    df['dkim'].value_counts().plot(kind='bar', color='steelblue', alpha=0.5, label='dkim')
    df['spf'].value_counts().plot(kind='bar', color='darkorange', alpha=0.5, label='spf')
    df['adkim'].value_counts().plot(kind='bar', color='forestgreen', alpha=0.5, label='adkim (r=brak podpisu)')
    df['aspf'].value_counts().plot(kind='bar', color='indianred', alpha=0.5, label='aspf (r=brak SPF)')
    df['result'].value_counts().plot(kind='bar', color='mediumseagreen', alpha=0.5, label='result (wynik autentykacji DKIM)')
    df['comment'].value_counts().plot(kind='bar', color='slateblue', alpha=0.5, label='comment')
    plt.title('Porównanie pass do fail dla dkim, spf i comment (arc=fail)')
    plt.xlabel('Wynik autentykacji')
    plt.ylabel('Ilość wystąpień')
    plt.legend()
    plt.xticks(rotation=0)
    plt.tight_layout()
    output_path = os.path.join(output_folder, "autentykacja_plot_1.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    # Wykres porównujący pass do fail dla dkim, spf i comment (arc=fail)
    plt.figure(figsize=(14, 10))  # Zwiększenie rozmiaru wykresu
    ax = df[['dkim', 'spf', 'adkim', 'aspf', 'result', 'comment']].apply(pd.Series.value_counts).T.plot(kind='bar', stacked=True, alpha=1.0)  # Zmniejszenie przezroczystości
    plt.title('Porównanie pass do fail dla dkim, spf i comment')
    plt.xlabel('Typ autentykacji')
    plt.ylabel('Ilość wystąpień')
    plt.legend(title='Wynik autentykacji')
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    # Dodanie etykiet do legendy
    handles, labels = ax.get_legend_handles_labels()
    labels = ['dkim (pass)', 'dkim (fail)', 'spf (pass)', 'spf (fail)', 'adkim (pass)', 'adkim (fail)', 'aspf (pass)', 'aspf (fail)', 'result (pass)', 'result (fail)', 'comment (arc=fail)']
    plt.legend(handles, labels)
    output_path = os.path.join(output_folder, "autentykacja_stacked_plot.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    # Wykres porównujący pass do fail dla dkim, spf i comment (arc=fail)
    plt.figure(figsize=(10, 6))
    categories = df[['dkim', 'spf', 'adkim', 'aspf', 'result', 'comment']]
    colors = [np.random.rand(len(df[category])) for category in categories]  # Generowanie losowych kolorów dla każdej kategorii
    
    for i, category in enumerate(categories):
        df[category].value_counts().plot(kind='bar', alpha=0.6, label=category, color=plt.cm.viridis(colors[i]))
        
    plt.title('Porównanie pass do fail dla dkim, spf i comment (arc)')
    plt.xlabel('Wynik autentykacji')
    plt.ylabel('Ilość wystąpień')
    plt.legend()
    plt.xticks(rotation=0)
    plt.tight_layout()
    output_path = os.path.join(output_folder, "autentykacja_plot.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji_DKIM.jpg
    colors_1 = np.random.rand(len(df['dkim']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['dkim'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_1))
    plt.title('Wynik autentykacji DKIM')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Wynik_autentykacji_DKIM.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji_SPF.jpg
    colors_2 = np.random.rand(len(df['spf']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['spf'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_2))
    plt.title('Wynik autentykacji SPF')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Wynik_autentykacji_SPF.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji.jpg
    colors_3 = np.random.rand(len(df['result']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['result'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_3))
    plt.title('Wynik autentykacji')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Wynik_autentykacji.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji_adkim.jpg
    colors_4 = np.random.rand(len(df['adkim']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['adkim'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_4))
    plt.title('Wynik autentykacji adkim')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Wynik_autentykacji_adkim.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji_aspf.jpg
    colors_5 = np.random.rand(len(df['aspf']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['aspf'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_5))
    plt.title('Wynik autentykacji aspf')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Wynik_autentykacji_aspf.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Polityka_DMARC_p.jpg
    colors_6 = np.random.rand(len(df['p']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['p'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_6))
    plt.title('Polityka DMARC (p) dla domen')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Polityka_DMARC_p.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Polityka_DMARC_sp_dla_poddomen.jpg
    colors_7 = np.random.rand(len(df['sp']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['sp'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_7))
    plt.title('Polityka DMARC (sp) dla poddomen')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Polityka_DMARC_sp.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Procentowy_poziom_raportowania_DMARC_PCT.jpg
    colors_8 = np.random.rand(len(df['pct']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['pct'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_8))
    plt.title('Procentowy poziom raportowania DMARC (PCT)')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Procentowy_poziom_raportowania_DMARC_PCT.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Postępowanie_z_wiadomościami_bez_polityki_DMARC_na_serwerach_odbiorczych.jpg
    colors_9 = np.random.rand(len(df['np']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['np'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_9))
    plt.title('Postępowanie z wiadomościami bez polityki DMARC na serwerach odbiorczych')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Postępowanie_z_wiadomościami_bez_polityki_DMARC_na_serwerach_odbiorczych.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Komentarze_autentykacji.jpg
    colors_10 = np.random.rand(len(df['comment']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['comment'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_10))
    plt.title('Komentarze autentykacji')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Komentarze_autentykacji.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()
    
    # Wykres kołowy dla nadawcy koperty
    colors_11 = np.random.rand(len(df['envelope_from']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['envelope_from'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_11))
    plt.title('Nadawca koperty')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    output_path = os.path.join(output_folder, "Nadawca_koperty.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()

    #Zakres_SPF.jpg
    colors_12 = np.random.rand(len(df['scope']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['scope'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_12), labels=None)
    plt.title('Zakres SPF')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    # Tworzenie legendy na podstawie danych z DataFrame
    legend_labels = df['scope'].unique()
    plt.legend(labels=legend_labels, loc="upper right", bbox_to_anchor=(0, 0), title='Zakres SPF', ncol=6)

    plt.tight_layout()
    output_path = os.path.join(output_folder, "Zakres_SPF.jpg")
    plt.savefig(output_path)  # Zapisz wykres do pliku JPEG
    plt.close()
    

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Analizator DMARC")
    # Przycisk do wyboru plików archiwum
    btn_choose_files = tk.Button(root, text="Wybierz pliki archiwum z raportami DMARC", command=choose_files)
    btn_choose_files.pack()

    # Przycisk do wyboru pliku Excela
    btn_select_excel = tk.Button(root, text="Wybierz plik Excela", command=select_excel_file)
    btn_select_excel.pack(pady=10)

    root.mainloop()