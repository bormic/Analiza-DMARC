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

def extract_and_parse_dmarc_report(folder_path, output_file):
    print("Rozpoczęcie ekstrakcji i analizy raportów DMARC...")
    try:
        # Bak_folder = os.path.join(folder_path, "bak")
        bak_folder = "C:/Users/m.borkowski/Desktop/dmarc/bak_arch"
        os.makedirs(bak_folder, exist_ok=True)  # Tworzenie katalogu "bak" jeśli nie istnieje

        # Inicjujemy pustą listę, w której będziemy przechowywać dane dla poszczególnych plików
        data_list = []

        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if filename.endswith(".zip"):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(folder_path)
                # Przenieś archiwum do katalogu "bak"
                shutil.move(file_path, os.path.join(bak_folder, filename))
            elif filename.endswith(".gz"):
                with gzip.open(file_path, 'rb') as gz_file:
                    with open(os.path.join(folder_path, os.path.basename(file_path)[:-3]), 'wb') as xml_file:
                        xml_file.write(gz_file.read())
                # Przenieś archiwum do katalogu "bak"
                shutil.move(file_path, os.path.join(bak_folder, filename))

        # Przeszukaj rozpakowane pliki i analizuj pliki XML
        for filename in os.listdir(folder_path):
            if filename.endswith(".xml"):
                xml_file_path = os.path.join(folder_path, filename)
                data = extract_data_from_xml(xml_file_path)
                if data is not None:
                    data_list.append(data)

        # Tworzymy ramkę danych pandas
        df = pd.DataFrame(data_list)

        # Sprawdź, czy istnieje plik Excela
        if os.path.exists(output_file):
            # Wczytaj istniejący plik Excela
            existing_df = pd.read_excel(output_file)
            # Połącz nowe dane z istniejącymi danymi
            df = pd.concat([existing_df, df], ignore_index=True)

        # Dodaj aktualną datę i godzinę do ramki danych
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df["Aktualna data"] = now
        # Konwertujemy czas z formatu UNIX Epoch na bardziej czytelny format
        df['begin'] = pd.to_datetime(df['begin'], unit='s')
        df['end'] = pd.to_datetime(df['end'], unit='s')

        # Tworzymy mapowanie wartości
        mapping = {'r': 'Brak DKIM', 's': 'Wyłączony DKIM', 'n': 'Włączony DKIM'}

        # Zastępujemy wartości w kolumnie 'adkim' za pomocą mapowania
        df['adkim'] = df['adkim'].map(mapping)

        # Tworzymy mapowanie wartości
        mapping_aspf = {'r': 'Brak SPF', 's': 'Wyłączony SPF', 'n': 'Włączony SPF'}

        # Zastępujemy wartości w kolumnie 'aspf' za pomocą mapowania
        df['aspf'] = df['aspf'].map(mapping_aspf)

        df['comment'] = df['comment'].replace({
            'arc=fail': 'Niepowodzenie weryfikacji ARC',
            'arc=pass': 'Powodzenie weryfikacji ARC',
            'arc=none': 'Brak wyniku weryfikacji ARC',
            'arc=invalid': 'Nieprawidłowy wynik weryfikacji ARC',
            'arc=timestamp_error': 'Błąd czasu weryfikacji ARC',
            'arc=signature_error': 'Błąd podpisu weryfikacji ARC',
            'arc=authentication_results_mismatch': 'Niezgodność wyników autentykacji weryfikacji ARC'
        })

        df['pct'] = df['pct'].replace({
            '100': 'Wymagane raportowanie dla wszystkich wiadomości',
            '50': 'Wymagane raportowanie dla 50% wiadomości',
            '0': 'Brak wymaganego raportowania'
        })

        df['np'] = df['np'].replace({
            'none': 'Brak polityki (brak działań)',
            'quarantine': 'Umieszczenie w kwarantannie',
            'reject': 'Odrzucenie wiadomości'
        })

        df['disposition'] = df['disposition'].replace({
            'none': 'Brak autentykacji',
            'quarantine': 'Umieszczono w kwarantannie',
            'reject': 'Odrzucono wiadomość'
        })

        df['dkim'] = df['dkim'].replace({
            'pass': 'Poprawna weryfikacja DKIM',
            'fail': 'Błąd weryfikacji DKIM',
            'none': 'Brak autentykacji DKIM'
        })

        df['p'] = df['p'].replace({
            'none': 'Brak polityki DMARC',
            'quarantine': 'Umieszczono w kwarantannie',
            'reject': 'Odrzucono wiadomość'
        })

        df['sp'] = df['sp'].replace({
            'none': 'Brak polityki dla poddomen',
            'quarantine': 'Umieszczono w kwarantannie dla poddomen',
            'reject': 'Odrzucono wiadomość dla poddomen'
        })

        # Tworzenie słownika z tłumaczeniami
        translation = {
            "mfrom": "Zakres SPF: Od pola MAIL FROM - ewentualna zmiana nadawcy i odbiorcy może wpłynąć na wynik weryfikacji SPF",
            "helo": "Zakres SPF: Od pola HELO - ewentualna zmiana domeny nadawcy może wpłynąć na wynik weryfikacji SPF",
            "pra": "Zakres SPF: Od pola PRA - ewentualna zmiana domeny nadawcy może wpłynąć na wynik weryfikacji SPF",
            "explanation": "Zakres SPF: Dostęp do rekordu SPF jest tylko dla objaśnienia"
        }

        # Utworzenie kolumny z tłumaczeniami
        df['scope'] = df['scope'].map(translation)
        
        # Zamieniamy "\n", "None" i NaN na "Brak danych"
        #df.replace({None: 'Brak danych', np.nan: 'Brak danych'}, inplace=True)
        df.replace({r'^\s*$': 'Brak danych', np.nan: 'Brak danych', 'None': 'Brak danych'}, inplace=True, regex=True)
        
        # wyjcie df
        #pd.set_option('display.max_rows', None)  # Wyświetl wszystkie wiersze
        #pd.set_option('display.max_columns', None)  # Wyświetl wszystkie kolumny
        #display(df)

        # Zapisz dane do pliku Excela
        df.to_excel(output_file, index=False)

        # Po zakończeniu przetwarzania usuń rozpakowane pliki
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            os.remove(file_path)

        print("Ekstrakcja i analiza zakończona pomyślnie!")

    except zipfile.BadZipFile as e:
        print(f"Błąd: Nieprawidłowy format pliku ZIP - {e}")
    except Exception as e:
        print(f"Błąd podczas analizy pliku: {e}")

def extract_data_from_xml(xml_file):
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

def generate_plots_from_excel(excel_file):
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
    plt.savefig("org_name_plot.jpg")  # Zapisz wykres do pliku JPEG
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
    plt.savefig("autentykacja_plot_1.jpg")  # Zapisz wykres do pliku JPEG
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
    
    plt.savefig("autentykacja_stacked_plot.jpg")  # Zapisz wykres do pliku JPEG
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
    plt.savefig("autentykacja_plot.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji_DKIM.jpg
    colors_1 = np.random.rand(len(df['dkim']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['dkim'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_1))
    plt.title('Wynik autentykacji DKIM')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Wynik_autentykacji_DKIM.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji_SPF.jpg
    colors_2 = np.random.rand(len(df['spf']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['spf'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_2))
    plt.title('Wynik autentykacji SPF')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Wynik_autentykacji_SPF.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji.jpg
    colors_3 = np.random.rand(len(df['result']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['result'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_3))
    plt.title('Wynik autentykacji')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Wynik_autentykacji.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji_adkim.jpg
    colors_4 = np.random.rand(len(df['adkim']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['adkim'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_4))
    plt.title('Wynik autentykacji adkim')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Wynik_autentykacji_adkim.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Wynik_autentykacji_aspf.jpg
    colors_5 = np.random.rand(len(df['aspf']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['aspf'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_5))
    plt.title('Wynik autentykacji aspf')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Wynik_autentykacji_aspf.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Polityka_DMARC_p.jpg
    colors_6 = np.random.rand(len(df['p']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['p'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_6))
    plt.title('Polityka DMARC (p) dla domen')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Polityka_DMARC_p.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Polityka_DMARC_sp_dla_poddomen.jpg
    colors_7 = np.random.rand(len(df['sp']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['sp'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_7))
    plt.title('Polityka DMARC (sp) dla poddomen')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Polityka_DMARC_sp.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Procentowy_poziom_raportowania_DMARC_PCT.jpg
    colors_8 = np.random.rand(len(df['pct']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['pct'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_8))
    plt.title('Procentowy poziom raportowania DMARC (PCT)')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Procentowy_poziom_raportowania_DMARC_PCT.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Postępowanie_z_wiadomościami_bez_polityki_DMARC_na_serwerach_odbiorczych.jpg
    colors_9 = np.random.rand(len(df['np']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['np'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_9))
    plt.title('Postępowanie z wiadomościami bez polityki DMARC na serwerach odbiorczych')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Postępowanie_z_wiadomościami_bez_polityki_DMARC_na_serwerach_odbiorczych.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()

    #Komentarze_autentykacji.jpg
    colors_10 = np.random.rand(len(df['comment']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['comment'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_10))
    plt.title('Komentarze autentykacji')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Komentarze_autentykacji.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()
    
    # Wykres kołowy dla nadawcy koperty
    colors_11 = np.random.rand(len(df['envelope_from']))  # Przykładowa skala kolorów
    plt.figure(figsize=(8, 8))
    df['envelope_from'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.viridis(colors_11))
    plt.title('Nadawca koperty')
    plt.axis('equal')
    plt.gca().set_ylabel('')  # Usunięcie napisu "count"
    plt.tight_layout()
    plt.savefig("Nadawca_koperty.jpg")  # Zapisz wykres do pliku JPEG
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

    #plt.tight_layout()
    plt.savefig("Zakres_SPF.jpg")  # Zapisz wykres do pliku JPEG
    plt.close()
    

if __name__ == "__main__":
    extract_and_parse_dmarc_report("/dmarc/raports", "/dmarc/wyniku.xlsx")
    generate_plots_from_excel("/dmarc/wyniku.xlsx")