from tkinter import filedialog as fd

def get_file_path():
    file_path = fd.askopenfilename()
    return file_path

def get_folder_path():
    folder_path = fd.askdirectory()
    return folder_path