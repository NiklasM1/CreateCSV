import numpy as np
import pandas as pd
import os
from openpyxl import *


# Trennt Gruppen Name in Nummer und Tutoren Nummer
def get_group_info(group_name):
    info = group_name[0].split()
    number = info[1].replace(',', '')
    firstname = info[len(info) - 2]
    lastname = info[len(info) - 1]

    if '@' in lastname:
        firstname = info[len(info) - 3]
        lastname = info[len(info) - 2]

    return number, firstname, lastname


# Öffnet CVS und EXCEL Dateien
def open_files():
    group_list = 'Gruppenliste.csv'
    excel_template = 'NN_Vorname_Nachname.xlsx'

    try:
        group = np.array(pd.read_csv(group_list, sep=';'))
        excel = load_workbook(excel_template)
        sheet = excel['Übersicht']
    except OSError:
        os.removedirs('output')
        print('Konnte Dateien nicht öffnen, breche ab. Stellen Sie sicher das im selben Verzeichnis '
              '%s und %s vorhanden sind.' % (excel_template, group_list))
        exit()
        return

    return group, excel, sheet


# Erstellt output Ordner
def create_dir():
    try:
        os.makedirs('output')
    except OSError:
        print('Konnte Verzeichnis nicht erstellen, breche ab.')
        exit()
        return
    return


def main():
    create_dir()
    group, excel, sheet = open_files()

    index = 2

    # Für jeden eintrag der Liste
    for i in range(0, group.shape[0]):
        # Wenn Mitglied von Veranstallter oder Funktion skippen
        if 'Veranstalter:innen' in group[i][0] or 'Funktion' in group[i][0]:
            continue

        # Zu viele Teilnehmer?
        if index > 46:
            print('Zu viele Teilnehmer für Vorlage in gruppe %s' % group[i])
            continue

        # sonst Tutoriums Nummer und Namen des Tutoren holen
        number, tutor_firstname, tutor_lastname = get_group_info(group[i])

        # kopiere von CSV zu XLSX
        sheet['A%s' % index] = group[i][2]
        sheet['B%s' % index] = group[i][3]
        sheet['C%s' % index] = group[i][8]

        # überprüfen ob Speichern oder weiter mit der selben Gruppe
        if i+1 < len(group) and number == get_group_info(group[i+1])[0]:
            index += 1
        else:
            index = 2
            excel.save('output/%s_%s_%s.xlsx' % (number, tutor_firstname, tutor_lastname))
            group_temp, excel, sheet = open_files()


if __name__ == '__main__':
    main()
