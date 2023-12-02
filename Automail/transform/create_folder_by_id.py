import os

def create_folder(name, code):
    # create dir path by name - code
    folder_path = os.path.join(r"C:\Users\thomas.thai\Downloads\automail\Data sending\Data gửi sale", f"{name} - {code}")
    
    # check folder exists
    if not os.path.exists(folder_path):
        # create folder
        os.makedirs(folder_path)
        print(f"Created folder: {folder_path}")
    else:
        print(f"Folder already exists: {folder_path}")
    
    return folder_path

