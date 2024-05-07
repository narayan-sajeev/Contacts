import pandas as pd
import vobject

new_cons = 'contacts-1.vcf'
new_lst = []

old_cons = '/Users/narayansajeev/Desktop/Misc/Contacts.xlsx'
old_lst = []

with open(new_cons, 'r') as file:
    for contact in vobject.readComponents(file.read()):
        full = contact.fn.value.split(' ')
        first = full[0].capitalize()
        last = full[1].capitalize() if len(full) > 1 else ''
        phone = contact.tel.value.replace(' ', '').replace('+1', '').replace('+', '').replace(')', '')
        phone = phone.replace('(', '').replace('-', '')
        new_lst.append([first, last, phone])

df = pd.read_excel(old_cons).fillna('')

for row in df.iterrows():
    val = row[1]
    old_lst.append([str(val.First), str(val.Last), str(val.Phone)])

print('First\tLast\tPhone')

for row in old_lst:
    print('\t'.join(row))

new_lst = [x for x in new_lst if x not in old_lst]

print('\n\n')

if len(new_lst) < 1:
    print('No new contacts found.')

for row in new_lst:
    print('\t'.join(row))
