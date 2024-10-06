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
                contact.org.value[0] if hasattr(contact, 'org') else ''
            ]
            for contact in vobject.readComponents(file.read())
            if hasattr(contact, 'fn') and hasattr(contact, 'tel')
        ]


def read_excel_contacts(file_path):
    """Read contacts from an Excel file."""
    df = pd.read_excel(file_path).fillna('')
    return [[val.Name, str(val.Phone), val.Relation] for _, val in df.iterrows()]


def remove_old_contacts(saved_lst, xl_lst):
    """Remove old contacts from saved list."""
    return [contact for contact in saved_lst if contact not in xl_lst]


def find_and_append_contacts(saved_lst, xl_lst):
    """Find and append contacts based on name or phone number."""
    contacts = []
    saved_names = {contact[0] for contact in saved_lst}
    saved_phones = {contact[1] for contact in saved_lst}

    for row in xl_lst:
        if row[0] in saved_names:
            contact = next((contact for contact in saved_lst if contact[0] == row[0]), None)
        elif row[1] in saved_phones:
            contact = next((contact for contact in saved_lst if contact[1] == row[1]), None)
        else:
            contact = None

        if contact:
            contacts.append(contact)
            saved_lst.remove(contact)
        else:
            contacts.append(row)

    contacts.extend(saved_lst)  # Append remaining saved contacts
    return contacts


def separate_contacts(contacts):
    """Separate contacts into international and US based on phone number length."""
    int_contacts = [contact for contact in contacts if len(contact[1]) != 10]
    us_contacts = [contact for contact in contacts if len(contact[1]) == 10]
    return sorted(int_contacts) + sorted(us_contacts)


def print_contacts(contacts):
    """Print contacts in a tabulated format."""
    print('Name\tPhone\tRelation')
    for row in contacts:
        print('\t'.join(row))


def main():
    saved_lst = parse_vcf(SAVED_CONS)
    xl_lst = read_excel_contacts(XL_CONS)

    # Remove old contacts from the saved list
    saved_lst = remove_old_contacts(saved_lst, xl_lst)

    if not saved_lst:
        print('No new contacts to add.')
        return

    contacts = find_and_append_contacts(saved_lst, xl_lst)
    contacts = separate_contacts(contacts)

    print_contacts(contacts)


if __name__ == '__main__':
    main()
