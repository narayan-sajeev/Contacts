import pandas as pd
import vobject

new_cons = 'contacts-1.vcf'
new_lst = []

old_cons = 'Contacts.xlsx'
old_lst = []

with open(new_cons, 'r') as file:
    # Loop through contacts in vcf file
    for contact in vobject.readComponents(file.read()):
        name = ' '.join([n.capitalize() for n in contact.fn.value.split()])
        # Retrieve phone number
        phone = contact.tel.value.replace(' ', '').replace('+1', '').replace('+', '').replace(')', '')
        phone = phone.replace('(', '').replace('-', '')
        new_lst.append([name, phone])

# Read old contacts from excel file
df = pd.read_excel(old_cons).fillna('')

for row in df.iterrows():
    val = row[1]
    # Append old contacts to list
    old_lst.append([val.Name, str(val.Phone)])

print('Name\tPhone')

# Sort old contacts by last name
old_lst = sorted(old_lst, key=lambda x: x[0].split()[-1])

for old_row in old_lst:
    print('\t'.join(old_row))

# Remove old contacts from new list
new_lst = [x for x in new_lst if x not in old_lst]

print('\n\n')

if len(new_lst) < 1:
    print('No new contacts found.')

for new_row in new_lst:
    print('\t'.join(new_row))
