#https://stackoverflow.com/questions/47212459/automating-comparison-of-word-documents-using-python
from os import getcwd, path
from sys import argv, exit
from win32com import client
from pathlib import Path

WORD_APP_ID = "Word.Application"
FILEFORMAT_PDF = 17 # Magic number from ChatGPT, tested it and it works

def cmp_name(old, new):
    file_name = Path(old).stem+"_"+Path(new).stem+".docx"
    return file_name

def generate_compare(dir, original_file, modified_file, output_file):
  app = client.gencache.EnsureDispatch(WORD_APP_ID)
  app.Visible = 0
  app.CompareDocuments(app.Documents.Open(dir + original_file), app.Documents.Open(dir + modified_file))
  app.ActiveDocument.ActiveWindow.View.Type = 3 # prevent that word opens itself
  if output_file.endswith(".pdf"):
    app.ActiveDocument.SaveAs(FileName = path.join(dir, output_file), FileFormat=FILEFORMAT_PDF) 
  else:
    app.ActiveDocument.SaveAs(FileName = path.join(dir, output_file))
  app.ActiveDocument.Close(SaveChanges=0) # Supresses annoying loos of work prompt
  app.Quit()

def cmp(original_file, modified_file, output_file):
  dir = getcwd() + '\\'

  # some file checks
  if not path.exists(dir+original_file):
    print('Original file does not exist')
    exit()
  if not path.exists(dir+modified_file):
    print('Modified file does not exist')
    exit()

  print('Working...')

  generate_compare(dir, original_file, modified_file, output_file)

  print('Saved compare as: '+ output_file)

def main():
  if len(argv) != 4:
    print('Usage: gendiff <original_file> <modified_file> <compared_file>')
    exit()
  cmp(argv[1], argv[2], argv[3])

if __name__ == '__main__':
  main()