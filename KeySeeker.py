import os
from docx import Document

try: #try to import the doc reader library
    import textract
except ImportError:
    textract = None 

def read_docx(main_path): #read docx files
    try:
        doc = Document(main_path)
        text = "\n".join([par.text for par in doc.paragraphs])
        return text
    except Exception as e:
        print(f"Error {main_path}: {e}")
        return ""

def read_doc(main_path): #read doc files
    if not textract:
        print(f"textract not installed. Cannot read {main_path}")
        return ""
    try:
        text = textract.process(main_path).decode("utf-8")
        return text
    except Exception as e:
        print(f"Error {main_path}: {e}")
        return ""

def search_for_the_key(folder, key): # search for the key in the files
    found = []

    for root, _, files in os.walk(folder):
        for file in files:
            if file.lower().endswith(('.doc', '.docx')):
                main_path = os.path.join(root, file)
                text = ""

                if file.lower().endswith('.docx'):
                    text = read_docx(main_path)
                elif file.lower().endswith('.doc'):
                    text = read_doc(main_path)

                if key.lower() in text.lower():
                    found.append(main_path)

    if found:
        print(f"\n'{key}' found in files:")
        for main_path in found:
            print(f"- {main_path}")
    else:
        print(f"\n'{key}' was not found in any Word file in the folder '{folder}'.")

# Execution 
folder = None
key = None

while True:
    if folder == None and key == None:
        folder = input("Enter the path of the folder:")
        key = input("Enter the key you want to search for: ")
        search_for_the_key(folder, key)

    else:
        option_menu = input("-"*50 + "\nSelect: \n1 - Change Folder\n2 - Change Key\n3 - Change both\n4 - Quit\n-->  ")
        match option_menu:
            case "1":
                folder = input("Enter the path of the folder:")
                search_for_the_key(folder, key)
            case "2":
                key = input("Enter the key you want to search for: ")
                search_for_the_key(folder, key)
            case "3":
                folder = input("Enter the path of the folder:")
                key = input("Enter the key you want to search for: ")
                search_for_the_key(folder, key)
            case "4":
                break
