import csv
import random
import string




#scripte mdp
"""
def generate_password():
    length = 13 - len("Protec")
    characters = string.ascii_letters + string.digits
    random_part = ''.join(random.choice(characters) for i in range(length))
    symbol = random.choice(string.punctuation)
    return "Protec" + random_part + symbol

input_file = 'comptes-protec.csv'
output_file = 'comptes-protec1.csv'

with open(input_file, mode='r', newline='') as infile, open(output_file, mode='w', newline='') as outfile:
    reader = csv.DictReader(infile, delimiter=';')
    fieldnames = reader.fieldnames
    writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
    
    writer.writeheader()
    for row in reader:
        row['PASSWORD'] = generate_password()
        writer.writerow(row)

print("Mots de passe mis à jour et sauvegardés dans", output_file)


#script mail

def update_email(email):
    local_part = email.split('@')[0]
    new_domain = "protec-groupe.com"
    return f"{local_part}@{new_domain}"

input_file = 'comptes-protec.csv'
output_file = 'comptes-protec1.csv'

with open(input_file, mode='r', newline='') as infile, open(output_file, mode='w', newline='') as outfile:
    reader = csv.DictReader(infile, delimiter=';')
    fieldnames = reader.fieldnames
    writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
    
    writer.writeheader()
    for row in reader:
        row['MESSAGERIE'] = update_email(row['MESSAGERIE'])
        writer.writerow(row)

print("Adresses e-mail mises à jour et sauvegardées dans", output_file)


#scrit Identifiant
import csv

def generate_identifiant(prenom, nom):
    return f"{prenom.lower()}.{nom.lower()}"

def process_csv(file_name):
    with open(file_name, mode='r', encoding='utf-8') as infile:
        reader = csv.DictReader(infile, delimiter=';')
        rows = list(reader)

    with open(file_name, mode='w', encoding='utf-8', newline='') as outfile:
        fieldnames = reader.fieldnames
        writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')

        writer.writeheader()
        for row in rows:
            prenom = row['PRENOM']
            nom = row['NOM']
            row['IDENTIFIANT'] = generate_identifiant(prenom, nom)
            writer.writerow(row)

if __name__ == "__main__":
    file_name = 'comptes-protec.csv'
    process_csv(file_name)


import csv

def process_csv(file_name):
    with open(file_name, mode='r', encoding='utf-8') as infile:
        reader = csv.DictReader(infile, delimiter=';')
        rows = list(reader)

    with open(file_name, mode='w', encoding='utf-8', newline='') as outfile:
        fieldnames = reader.fieldnames
        writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')

        writer.writeheader()
        for row in rows:
            row['MESSAGERIE'] = row['MESSAGERIE'].lower()
            writer.writerow(row)

if __name__ == "__main__":
    file_name = 'comptes-protec.csv'
    process_csv(file_name)
"""

import csv

def process_csv(file_name):
    with open(file_name, mode='r', encoding='utf-8') as infile:
        reader = csv.DictReader(infile, delimiter=';')
        rows = list(reader)
        fieldnames = reader.fieldnames

    with open(file_name, mode='w', encoding='utf-8', newline='') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
        writer.writeheader()
        for row in rows:
            row['MESSAGERIE'] = row['MESSAGERIE'].lower()
            row['PRENOM'] = row['PRENOM'].split()[0]  # Ne garder que le premier nom
            row['IDENTIFIANT'] = row['IDENTIFIANT'].split()[0]  # Ne garder que le premier nom
            writer.writerow(row)

if __name__ == "__main__":
    file_name = 'comptes-protec.csv'
    process_csv(file_name)
