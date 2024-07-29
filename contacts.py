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

# If no new contacts to add
if len(saved_lst) < 1:
    print('No new contacts to add.')
    exit()

# Create final contacts list
contacts = []

# Remove phone & relation from saved list
saved_names = [x[0] for x in saved_lst]

# Loop through xl list
for row in xl_lst:
    # If contact is in saved list
    if row[0] in saved_names:
        # Find contact in saved list
        for contact in saved_lst:
            # If contact is found
            if contact[0] == row[0]:
                # Append contact to contacts
                contacts.append(contact)
                # Remove contact from saved list
                saved_lst.remove(contact)

    else:
        # Append contact to contacts
        contacts.append(row)

# Loop through saved list
for contact in saved_lst:
    # Append contact to new contacts
    contacts.append(contact)

# Sort contacts by name
contacts = sorted(contacts)

# Print saved list
for row in contacts:
    print('\t'.join(row))
