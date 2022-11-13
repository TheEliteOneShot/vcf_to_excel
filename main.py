from pathlib import Path
INPUT_CONTACTS_FILE_NAME = "iphone_contacts.vcf"
INPUT_CONTACTS_FILE_LOCATION = f"{Path(__file__).parent.absolute()}\\{INPUT_CONTACTS_FILE_NAME}"
EXCEL_FILE_NAME = "output.csv"
EXCEL_OUTPUT_LOCATION=f"{Path(__file__).parent.absolute()}\\{EXCEL_FILE_NAME}"

f = open(INPUT_CONTACTS_FILE_LOCATION, "r", encoding="utf8")

lines = f.readlines()

people = []
excel_lines = []
person_id = -1

def remove_unecessary_characters(line):
    return line.replace("\n", "").replace("\\n", "").replace("\\", "").replace(",", " ")

def add_home_address_by_person_id(person_id, home_address):
    people[person_id]["HOME_ADDRESS"] = remove_unecessary_characters(home_address)

def add_work_address_by_person_id(person_id, work_address):
    people[person_id]["WORK_ADDRESS"] = remove_unecessary_characters(work_address)

def add_n_by_person_id(person_id, n):
    people[person_id]["N"] = remove_unecessary_characters(n)

def add_fn_by_person_id(person_id, fn):
    people[person_id]["FN"] = remove_unecessary_characters(fn)

for position, line in enumerate(lines, 0):
    if line.strip() == "BEGIN:VCARD":
        person_id = person_id + 1
        people.append({})
    if line.startswith("FN:"):
        add_fn_by_person_id(person_id, line[len("FN:"):])
    if line.startswith("N:"):
        add_n_by_person_id(person_id, line[len("N:"):])
    if line.startswith("ADR;TYPE=WORK:"):
        add_work_address_by_person_id(person_id, line[len("ADR;TYPE=WORK:"):])
    if line.startswith("ADR;TYPE=HOME:"):
        add_home_address_by_person_id(person_id, line[len("ADR;TYPE=HOME:"):])

excel_lines.append("NAME, ADDRESS, CITY, STATE, ZIP, COUNTRY \n")

for person in people:
    name = ""
    address = ""
    city = ""
    state = ""
    zip = ""
    country = "US"
    if "FN" in person:
        name = person["FN"]
    elif "N" in person:
        name = person["N"]
    else:
        name = "NO_NAME"
    
    address_items = []
    if "HOME_ADDRESS" in person:
        address_items = person["HOME_ADDRESS"].split(";")
    elif "WORK_ADDRESS" in person:
        address_items = person["WORK_ADDRESS"].split(";")
    if address_items:
        if len(address_items) >= 3:
            address = address_items[2]
        if len(address_items) >= 4:
            city = address_items[3]
        if len(address_items) >= 5:
            state = address_items[4]
        if len(address_items) >= 6:
            zip = address_items[5]
    excel_lines.append(f"{name}, {address}, {city}, {state}, {zip}, {country} \n")
    
output = open(EXCEL_OUTPUT_LOCATION, "w+", encoding="utf8")
output.writelines(excel_lines)