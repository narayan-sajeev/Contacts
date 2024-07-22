import pandas as pd
import vobject

saved_cons = 'contacts.vcf'
saved_lst = []

xl_cons = 'Contacts.xlsx'
xl_lst = []

with open(saved_cons, 'r') as file:
    # Loop through contacts in vcf file
    for contact in vobject.readComponents(file.read()):
        name = ' '.join([n.capitalize() for n in contact.fn.value.split()])
        rel = ''
        try:
            rel = contact.org.value[0]
        except:
            pass
        # Retrieve phone number
        try:
            phone = contact.tel.value.replace(' ', '').replace('+1', '').replace('+', '').replace(')', '')
            phone = phone.replace('(', '').replace('-', '')
            saved_lst.append([name, phone, rel])
        except:
            pass

# Read contacts from Excel file
df = pd.read_excel(xl_cons).fillna('')

for row in df.iterrows():
    val = row[1]
    # Append contacts to list
    xl_lst.append([val.Name, str(val.Phone), val.Relation])

# Remove old contacts from list
saved_lst = [x for x in saved_lst if x not in xl_lst]

# Define relations list
relations = []

# Remove relation from saved list
saved_no_relation = [x[:2] for x in saved_lst]

# Loop through saved list
for row in xl_lst:
    # If contact is in saved list
    if row[:2] in saved_no_relation:
        # Reformat contact
        reformat = [row[0], row[1], '']
        # Remove contact from saved list
        saved_lst.remove(reformat)
        # Append contact to relations list
        relations.append(row)

# Print saved list
for row in saved_lst:
    print('\t'.join(row))

# Print horizontal line
print('*' * 50)

# Print relations list
for row in relations:
    print('\t'.join(row))