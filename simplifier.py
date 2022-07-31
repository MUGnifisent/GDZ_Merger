import os
import shutil
import docx2pdf as d2p
from tqdm import tqdm
from PyPDF2 import PdfMerger

class Simplifier:
    PATH = os.getcwd()
    t_PATH = os.path.join(PATH, "temp")
    I_PATH = os.path.join(PATH, "Input")
    R_PATH = os.path.join(PATH, "Result")

    def __init__(self):
        try:
            os.mkdir(self.I_PATH)
            print("'Input' folder created in " + self.PATH)
        except:
            print("'Input' folder already exist in " + self.PATH + ", and will not be created.")
        try:
            os.mkdir(self.R_PATH)
            print("'Result' folder created in " + self.PATH)
        except:
            print("'Result' folder already exist in " + self.PATH + ", and will not be created.")
        try:
            os.mkdir(self.t_PATH)
            print("Temporary folder 'temp' created in " + self.PATH)
        except:
            shutil.rmtree(self.t_PATH)
            os.mkdir(self.t_PATH)
            print("Temporary folder 'temp' deleted and created in " + self.PATH 
            + "\nPlease, finish execution of the program properly.")


    def transfer(self):
        FILES = [f for f in os.listdir(self.I_PATH) if os.path.isfile(os.path.join(self.I_PATH, f))]
        print("Files " + str(FILES) +" entered into program.")
        for f in tqdm(FILES, ascii = True):
            f_ext = f.rpartition('.')[-1]
            match f_ext:
                case "pdf":
                    shutil.copy2(os.path.join(self.I_PATH, f), self.t_PATH)
                    print("File " + f + "was copied to 'temp' folder")
                case "docx":
                    f_newName = f.rpartition('.')[0] + ".pdf"
                    d2p.convert(os.path.join(self.I_PATH, f), os.path.join(self.t_PATH, f_newName))
                case _:
                    tqdm.write("For now, the program can't work with file " + f) #need to add functional

    def merge(self):
        FILES = [f for f in os.listdir(self.t_PATH) if os.path.isfile(os.path.join(self.t_PATH, f))]
        print("Files " + str(FILES) +" are going to merge.")
        merger = PdfMerger()
        for pdf in tqdm(FILES, ascii = True):
            merger.append(os.path.join(self.t_PATH, pdf))
        merger.write(os.path.join(self.R_PATH, "merge_result.pdf"))
        merger.close()

    def pause(self, case):
        match case:
            case 0:
                programPause = input("Press the <ENTER> key to continue...")
            case 1:
                print("This is your last chance to take all nessesary files from 'temp' directory")
                programPause = input("Press the <ENTER> key to continue, when done...")
    
    def clear(self):
        shutil.rmtree(self.t_PATH)

