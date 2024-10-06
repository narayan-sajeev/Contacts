import pandas as pd
import vobject

saved_cons = 'contacts.vcf'
xl_cons = 'Contacts.xlsx'

def clean_phone_number(phone):
    phone = phone.replace(' ', '').replace(')', '').replace('(', '').replace('+', '').replace('-', '')
    return phone[1:] if phone.startswith('1') and len(phone) == 11 else phone

def parse_vcf(file_path):
    with open(file_path, 'r') as file:
        return [
            [
                ' '.join(word.capitalize() for word in contact.fn.value.split()),
                clean_phone_number(contact.tel.value),
                contact.org.value[0] if hasattr(contact, 'org') else ''
            ]
            for contact in vobject.readComponents(file.read())
            if hasattr(contact, 'fn') and hasattr(contact, 'tel')
        ]

def read_excel_contacts(file_path):
    df = pd.read_excel(file_path).fillna('')
    return [[val.Name, str(val.Phone), val.Relation] for _, val in df.iterrows()]

saved_lst = parse_vcf(saved_cons)
xl_lst = read_excel_contacts(xl_cons)

# Remove old contacts from saved list
saved_lst = [contact for contact in saved_lst if contact not in xl_lst]

if not saved_lst:
    print('No new contacts to add.')
    exit()

contacts = []

saved_names = {contact[0] for contact in saved_lst}
saved_phones = {contact[1] for contact in saved_lst}

# Append contacts based on name or phone number
for row in xl_lst:
    if row[0] in saved_names:
        contact = next((contact for contact in saved_lst if contact[0] == row[0]), None)
        if contact:
            contacts.append(contact)
            saved_lst.remove(contact)
    elif row[1] in saved_phones:
        contact = next((contact for contact in saved_lst if contact[1] == row[1]), None)
        if contact:
            contacts.append(contact)
            saved_lst.remove(contact)
    else:
        contacts.append(row)

contacts.extend(saved_lst)

# Separate international and US contacts
int_contacts = [contact for contact in contacts if len(contact[1]) != 10]
us_contacts = [contact for contact in contacts if len(contact[1]) == 10]

contacts = sorted(int_contacts) + sorted(us_contacts)

# Print column headers
print('Name\tPhone\tRelation')
for row in contacts:
    print('\t'.join(row))
