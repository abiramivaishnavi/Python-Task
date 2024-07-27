print("WELCOME TO THE CONTACT DIRECTORY\n")

class ContactManager:
    def __init__(self):
        self.contacts = []

    def add_contact(self, name, phone, email, address):
        contact = {
            "name": name,
            "phone": phone,
            "email": email,
            "address": address,
        }
        self.contacts.append(contact)
        print(f"Contact '{name}' added successfully!")

    def view_contact_list(self):
        if not self.contacts:
            print("No contacts present.")
        else:
            print("Contact List:")
            for idx, contact in enumerate(self.contacts, start=1):
                print(f"{idx}. Name: {contact['name']}, Phone: {contact['phone']}")

    def search_contact(self, search_term):
        found_contacts = []
        for contact in self.contacts:
            if (
                search_term.lower() in contact["name"].lower()
                or search_term in contact["phone"]
            ):
                found_contacts.append(contact)
        if not found_contacts:
            print("No matching contacts found.")
        else:
            print("Matching Contacts:")
            for idx, contact in enumerate(found_contacts, start=1):
                print(f"{idx}. Name: {contact['name']}, Phone: {contact['phone']}")

    def update_contact(self, name, new_phone, new_email, new_address):
        for contact in self.contacts:
            if contact["name"].lower() == name.lower():
                contact["phone"] = new_phone
                contact["email"] = new_email
                contact["address"] = new_address
                print(f"Contact '{name}' updated successfully!")
                return
        print(f"Contact '{name}' not found.")

    def delete_contact(self, name):
        for contact in self.contacts:
            if contact["name"].lower() == name.lower():
                self.contacts.remove(contact)
                print(f"Contact '{name}' deleted successfully!")
                return
        print(f"Contact '{name}' not found.")


def main():
    contact_manager = ContactManager()

    while True:
        print("\nContact Management System Menu:")
        print("1. Add Contact Details")
        print("2. View Contact List")
        print("3. Search Contact")
        print("4. Update Contact")
        print("5. Delete Contact")
        print("6. Exit")

        choice = input("Enter your choice (1/2/3/4/5/6): ")

        if choice == "1":
            name = input("Enter name: ")
            phone = input("Enter phone number: ")
            email = input("Enter email: ")
            address = input("Enter address: ")
            contact_manager.add_contact(name, phone, email, address)

        elif choice == "2":
            contact_manager.view_contact_list()

        elif choice == "3":
            search_term = input("Enter name or phone number to search: ")
            contact_manager.search_contact(search_term)

        elif choice == "4":
            name = input("Enter  perfect name of the contact to update: ")
            new_phone = input("Enter your fancy phone number: ")
            new_email = input("Enter new email: ")
            new_address = input("Enter new address: ")
            contact_manager.update_contact(name, new_phone, new_email, new_address)

        elif choice == "5":
            name = input("enter perfect name of contact: ")
            contact_manager.delete_contact(name)

        elif choice == "6":
            print("Exiting.")
            break

        else:
            print("choice invalid. enter a valid choice.")


if __name__ == "__main__":
    main()
