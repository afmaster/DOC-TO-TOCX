import os
import win32com.client as win32

print("-------------start--------------")

word = win32.gencache.EnsureDispatch('Word.Application')

def save_as_docx(path):
    doc = word.Documents.Open(path)
    doc.SaveAs(path + 'x', FileFormat=16)  # 16 corresponds to wdFormatDocx
    doc.Close()

# Get the current directory
current_dir = os.path.dirname(os.path.realpath(__file__))

for dirpath, dirnames, filenames in os.walk(current_dir):
    for filename in filenames:
        print(filename)
        if filename.endswith('.doc'):
            save_as_docx(os.path.join(dirpath, filename))

word.Quit()
print("--------------DONE-------------")
