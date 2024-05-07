# Contact Comparison Tool

This Python script compares two sets of contact data: one from a VCF file (new contacts) and the other from an Excel
file (old contacts). It prints out the contacts that are present in the new list but not in the old list.

## Dependencies

- pandas: A powerful data manipulation library.
- vobject: A module for parsing and generating vCard and vCalendar files.

## How to Use

1. Update the `new_cons` variable with the path to your VCF file.
2. Update the `old_cons` variable with the path to your Excel file.
3. Run the script. The script will print out the contacts in the old list and any new contacts found.

## Output

The script prints out the contacts in the following format:

```
First   Last    Phone
```

If no new contacts are found, the script will print:

```
No new contacts found.
```

Otherwise, it will print out the new contacts in the same format as above.