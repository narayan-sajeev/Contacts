import pandas as pd
import vobject

new_cons = 'contacts.vcf'
new_lst = []

old_cons = 'Contacts.xlsx'
old_lst = []

with open(new_cons, 'r') as file:
    # Loop through contacts in vcf file
    for contact in vobject.readComponents(file.read()):
        name = ' '.join([n.capitalize() for n in contact.fn.value.split()])
        # Retrieve phone number
        try:
            phone = contact.tel.value.replace(' ', '').replace('+1', '').replace('+', '').replace(')', '')
            phone = phone.replace('(', '').replace('-', '')
            new_lst.append([name, phone])
        except:
            pass

# Read old contacts from Excel file
df = pd.read_excel(old_cons).fillna('')

for row in df.iterrows():
    val = row[1]
    # Append old contacts to list
    old_lst.append([val.Name, str(val.Phone)])

# Remove old contacts from new list
new_lst = [x for x in new_lst if x not in old_lst]

for new_row in new_lst:
    print('\t'.join(new_row))
