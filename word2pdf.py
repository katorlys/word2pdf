from win32com.client import Dispatch
import os

def doc2pdf(inFile, outFile, file):
    print("Converting:", file)
    word = Dispatch("Word.Application")
    doc = word.Documents.Open(inFile)
    doc.SaveAs(outFile, FileFormat = 17)
    doc.Close()
    word.Quit()
    print("Success:", file, "\n")

if __name__ == "__main__":
    print("Instruction:\n  You have to have Microsoft Word 2007 or above installed in your environment;\n  The paths must be complete, and not relative;\n  You cannot contain \'.doc\' except for the suffix in your file name.\n")
    while True:
        wordPath = input("Word path:").replace("\\", "\\\\")
        if os.path.exists(wordPath):
            break
        else:
            print("Word path does not exist, Please enter it again!")
    while True:
        pdfPath = input("PDF path:").replace("\\", "\\\\")
        if os.path.exists(pdfPath):
            break
        else:
            print("PDF path does not exist, Please enter it again!")
    print("\n")
    filelist = os.listdir(wordPath)
    success = 0
    failure = 0
    for file in filelist:
        if (file.lower().endswith(".doc") or file.lower().endswith(".docx")) and ("~$" not in file):
            inFile = wordPath + "\\" + file
            outFile = pdfPath + "\\" + file.lower().split('.doc')[0] + ".pdf"
            if os.path.exists(outFile):
                print("Failed, PDF exist:", outFile, "\n")
                failure += 1
                continue
            else:
                doc2pdf(inFile, outFile, file)
                success += 1
    print("\n", success + failure, "files converted, ",  "with", success, "success, ", failure, "failed.")
    input("Press Enter to quit.")