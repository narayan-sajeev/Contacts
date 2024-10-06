import pandas as pd
import vobject

# File paths for the VCF and Excel contacts
SAVED_CONS = 'contacts.vcf'
XL_CONS = 'Contacts.xlsx'


def clean_phone_number(phone):
    """Clean and format the phone number."""
    phone = phone.replace(' ', '').replace(')', '').replace('(', '').replace('+', '').replace('-', '')
    return phone[1:] if phone.startswith('1') and len(phone) == 11 else phone


def parse_vcf(file_path):
    """Parse VCF file and return a list of contacts."""
    with open(file_path, 'r') as file:
        return [
            [
                ' '.join(word.capitalize() for word in contact.fn.value.split()),
                clean_phone_number(contact.tel.value),
                getattr(contact, 'ORG', [''])[0]
            ]
            for contact in vobject.readComponents(file.read())
            if all(hasattr(contact, attr) for attr in ['fn', 'tel'])
        ]


def read_excel_contacts(file_path):
    """Read contacts from an Excel file."""
    df = pd.read_excel(file_path).fillna('')
    return df[['Name', 'Phone', 'Relation']].values.tolist()


def main():
    saved_lst = parse_vcf(SAVED_CONS)
    xl_lst = read_excel_contacts(XL_CONS)

    # Remove old contacts from the saved list
    saved_lst = [contact for contact in saved_lst if contact not in xl_lst]

    if not saved_lst:
        print('No new contacts to add.')
        return

    contacts = []

    # Map saved contacts by name
    saved_dict = {contact[0]: contact for contact in saved_lst}

    # Map saved contacts by phone
    saved_phones = {contact[1]: contact for contact in saved_lst}

    for _ in xl_lst:

        # Convert phone number to string
        row = [str(cell) for cell in _]

        if row[0] in saved_dict:
            contacts.append(saved_dict.pop(row[0]))
        elif row[1] in saved_phones:
            contacts.append(saved_phones.pop(row[1]))
        else:
            contacts.append(row)

    # Add remaining saved contacts
    contacts.extend(saved_dict.values())

    # Separate and sort contacts
    contacts = sorted(contacts, key=lambda c: (len(c[1]) == 10, c[0]))

    # Print contacts
    print('Name\tPhone\tRelation')
    for row in contacts:
        print('\t'.join(row))


if __name__ == '__main__':
    main()
