import PyPDF2
import re
import pandas as pd
import pathlib
import os
import openpyxl

#unused:

def printLines(text):
    startings=getStartIndexes(text)
    length=len(startings)
    if length>0:
        endings=getEndIndexes(text,startings)
        for i in range(length):
            print(text[ startings[i][0] : endings[i] ])

def reverseDict(dict):
    reverse=dict()
    for key in dict:
        val= dict[key]
        reverse[val]=key
    return reverse

def test4():
    for i in range(1,3):
        print (i)

def test3():
    path=Path()
    pathStr=str(pathlib.Path().absolute())
    #path.newPath(pathStr)
    print(path.getPath())
    print(path.usesPath())
    print(path.getPath())

def test2():
    file = open(r"C:\Users\USER\Desktop\SideNoterDocs\first_example.PDF", 'rb')

    fileReader = PyPDF2.PdfFileReader(file)

    length = fileReader.numPages

    lst=[]
    for i in range(length):
        text=fileReader.getPage(i).extractText()
        #print(f"Page number {i+1}:")
        #printLines(text)
        currentPageTable=pageTables(text)
        if(len(currentPageTable)>0):
            lst.extend(pageTables(text))

    print(lst)
    print(len(lst[0]))
    print(lst[0][0])
    print(type(float(lst[1][1].replace(",",""))))
    print(float(lst[1][1].replace(",",'')))

def test1():
    file = open(r"C:\Users\USER\Desktop\SideNoterDocs\first_example.PDF", 'rb')
    fileReader = PyPDF2.PdfFileReader(file)
    print(fileReader.numPages)
    #for key in fileReader.get_form_text_fields().values():
        #print(key)
    fileText=fileReader.getPage(1).extractText()
    patterDate=re.compile(r'\d{1,2}[/\\]\d{1,2}[/\\]\d{2,4}[ ]+[-,.0-9]+')# #\d{7,8}

    #(r'')#(r'[- 0-9,.]+\n\s*\d')#\d{7,8}

    matchesDates = patterDate.finditer(fileText)
    for match in matchesDates:
        print(match)
   # print(2)
    patternEnd=re.compile(r'\n')
    matchesEnd=patternEnd.finditer(fileText)
    #for match in matchesEnd:
    #    print(match)
    #print(fileText.replace(' ','`'))

#used:

LOWER_BOUNDARY=5
UPPER_BOUNDARY=8

def createDatafFrame(doubleList):
    num_col=len(doubleList[0])
    col_names=['date']
    if(num_col>4):
        for i in range(num_col-4):
            col_names.append(f'detail_{i+1}')
    col_names.extend(["payment_1","payment_2","concentration"])
    df = pd.DataFrame(doubleList, columns=col_names)

    return df

def getStartIndexes(text):
    lst=[]
    pattern = re.compile(r'\d{1,2}[/\\]\d{1,2}[/\\]\d{2,4}[ ]+\-?\d{1,3}(\.\d|\,\d)')#[-,.0-9]+
    matches = pattern.finditer(text)
    for match in matches:
        lst.append([match.start(),match.end()])
    return lst

def getEndIndexes(text, startings):
    lst=[]
    length = len(startings)
    currentStart = 0
    compileString = r"\d{" + str(LOWER_BOUNDARY) + "," +str(UPPER_BOUNDARY) + "}"
    pattern = re.compile(compileString)
    matches = pattern.finditer(text)
    for match in matches:
        starting=match.start()
        ending=match.end()
        if(starting > startings[currentStart][1]):
            lst.append(ending)
            currentStart+=1
        if(currentStart==length):
            break
    return lst

def pageTables(text):
    lst=[]
    startings = getStartIndexes(text)
    length = len(startings)
    if length > 0:
        endings = getEndIndexes(text, startings)
        for i in range(length):
            lst.append(text[ startings[i][0] : endings[i] ].split())
    return lst
def min(a,b):
    if(a<b):
        return a
    return b

def max(a,b):
    if(a>b):
        return a
    return b

def pdfToDf(fileReader,pdfRange=None):
    length = fileReader.numPages
    doubleLst=[]
    if(pdfRange is None):
        rn = range(length)
    elif(pdfRange[0] > length):
        rn = range(length)
    else:
        rn= range(max(0, pdfRange[0]-1), min(length, pdfRange[1]))
    for i in rn:
        text=fileReader.getPage(i).extractText()
        currentPageTable = pageTables(text)
        if (len(currentPageTable) > 0):
            doubleLst.extend(currentPageTable)
    return createDatafFrame(doubleLst)

def txtToDf(file):
    doubleLst=pageTables(file.read())
    return createDatafFrame(doubleLst)

def dfToCsv(df, folderPath,csvName=None):
    if(csvName is None):
        df.to_csv(f"{folderPath}.csv", index=False)
    else:
        df.to_csv(f"{folderPath}\\{csvName}.csv",index=False)

def dfToXlsx(df, folderPath,xlsxName=None):
    if(xlsxName is None):
        with pd.ExcelWriter(f"{folderPath}.xlsx") as writer:
            df.to_excel(writer, sheet_name="sheet1", index=False)
    else:
        with pd.ExcelWriter(f"{folderPath}\\{xlsxName}.xlsx") as writer:
            df.to_excel(writer, sheet_name="sheet1", index=False)

def sumDf(df):
    df["sum"] = df[["payment_1", "payment_2"]].apply(
        lambda x: round((float(x[0].replace(",", "")) + float(x[1].replace(",", ""))), 2), axis=1)
    return df

def compressDf(df, isSummed):
    if (not isSummed):
        df=sumDf(df)
    groups = df.groupby('concentration', as_index=False)
    sum_df = groups['sum'].sum()
    date_df = groups['date'].min()
    return pd.merge(sum_df, date_df, on="concentration")

def getFolder():
    str = input("Please enter folder path\n")
    return str

def getFileName():
    str = input("Please enter file name\n")
    return str

def getType():
    inputToInt = {"1":1,"2":2}
    str="0"
    while(str=="0"):
        print("Please choose file type:")
        str = input("For .txt file enter 1\nFor .pdf file enter 2\n")
        if str in inputToInt and not (str is None):
            return inputToInt.get(str)
        str="0"

def getCsvName():
    str = input("Please enter new name for the file\n")
    return str

def previous():
    typeToSuffix={1:".txt",2:".pdf"}
    fileReader=None
    while(fileReader is None):
        try:
            directory = getFolder()
            fileName = getFileName()
            fileType=getType()
            fileSuffix = typeToSuffix.get(fileType)
            path=directory+"\\"+fileName+fileSuffix
            if(fileType==1):
                file = open(f"{path}", 'rt', encoding='utf8')
                fileReader=file
            if(fileType==2):
                file = open(f"{path}", 'rb')
                fileReader = PyPDF2.PdfFileReader(file)
        except Exception as e:
            print(f"There is an Error:")
            print(e)
            fileReader = None
    if(fileType==1):
        df=txtToDf(fileReader)
    if(fileType==2):
        df = pdfToDf(fileReader)

    df = compressDf(df, False) #Beware
    success = False
    while (not success):
        try:
            csvName=getCsvName()
            dfToCsv(df,directory,csvName)
            success=True
        except Exception as e:
            print(f"There is an Error:")
            print(e)
            success=False

class Report:
    def __init__(self):
        self.df=None
        self.isSummed=False
        self.isCompressed=False

class Path:
    def __init__(self):
        if not os.path.exists("SideNoterSettings.txt"):
            file=open("SideNoterSettings.txt","w")
            file.write("Uses path=False\nPath=")
            file.close()
            self.__uses_path=False
            self.__path=""
        else:
            file=open("SideNoterSettings.txt","r")
            fileStr=file.read()
            file.close()
            pattern = re.compile(r'=.*')
            matches = pattern.finditer(fileStr)
            lst = []
            for match in matches:
                lst.append(fileStr[match.start():match.end()])

            if(len(lst)>0):
                if lst[0]=="=True":
                    self.__uses_path=True
                else:
                    self.__uses_path=False
                if(len(lst)>1):
                    self.__path=lst[1][1:len(lst[1])]
                    if(not os.path.exists(self.__path)):
                        self.__path=""
            else:
                self.__uses_path=False
                self.__path=""
            self.__update()


    def usesPath(self):
        return self.__uses_path

    def changePathStatus(self,boolean):
        self.__uses_path=boolean
        self.__update()

    def getPath(self):
        return self.__path

    def newPath(self,pathString):
        if(not os.path.exists(pathString)):
            return False
        self.__path=pathString
        self.__uses_path=True
        self.__update()
        return True

    def mainPath(self):
        self.__path=""
        self.__uses_path=True
        self.__update()

    def __update(self):
        file = open("SideNoterSettings.txt", "w")
        file.write(f"Uses path={self.__uses_path}\nPath={self.__path}")
        file.close()

def intSize(num):
    if(num==0):
        return 1
    if(num<0):
        num=-num
    count=0
    while(num!=0):
        count+=1
        num=num//10
    return count

class Settings:
    def __init__(self):
        self.path=Path()
        self.report=Report()
        self.book=None
        self.Keep=True

    def quit(self):
        self.Keep=False
        print("Goodbye")

    def activateFun(self, fun):
        try:
            fun()
        except Exception as e:
            print("Unexpected error occurred:")
            print(f"{e}\n")

    def activatePath(self):
        self.path.changePathStatus(True)

    def unActivatePath(self):
        self.path.changePathStatus(False)

    def useMainPath(self):
        self.path.mainPath()

    def newPath(self):
        pathStr = getFolder()
        existing = self.path.newPath(pathStr)
        if existing:
            print("Changes activated successfully\n")
        else:
            print("Path entered does not exist, changes won't be saved\n")

    def filePath(self):
        if(self.path.usesPath()):
            folderPath=self.path.getPath()
            if(folderPath==""):
                folderPath=str(pathlib.Path().absolute())
        else:
            folderPath=getFolder()
        fileName=getFileName()
        return folderPath+"\\"+fileName

    def addTxt(self):
        filePath=self.filePath()+".txt"
        if(not os.path.exists(filePath)):
            print(f"File does not exist as:\n{filePath}\nChanges won't be saved\n")
            return
        file = open(filePath, 'rt', encoding='utf8')
        self.report.df = txtToDf(file)
        self.report.isSummed=False
        self.report.isCompressed=False
        print("Changes activated successfully\n")


    def addPdf(self):
        filePath = self.filePath() + ".pdf"
        if (not os.path.exists(filePath)):
            print(f"File does not exist as:\n{filePath}\nChanges won't be saved\n")
            return
        file = open(filePath, 'rb')
        fileReader = PyPDF2.PdfFileReader(file)

        firstPage=input("Please enter first page (number only), or 0 for the whole file\n")
        try:
            firstPage=int(firstPage)
        except Exception as e:
            print(f"An error occurred, see below:\n{filePath}\nChanges won't be saved\n")
            return
        if(firstPage<=0):
            self.report.df = pdfToDf(fileReader)
        else:
            lastPage=input("Please enter last page (number only)\n")
            try:
                lastPage=int(lastPage)
            except Exception as e:
                print(f"An error occurred, see below:\n{filePath}\nChanges won't be saved\n")
                return
            if(lastPage<firstPage):
                lastPage=firstPage
            pdfRange=[firstPage,lastPage]
            self.report.df = pdfToDf(fileReader,pdfRange)

        self.report.isSummed = False
        self.report.isCompressed = False
        print("Changes activated successfully\n")

    def addCsv(self):
        filePath = self.filePath() + ".csv"
        if (not os.path.exists(filePath)):
            print(f"File does not exist as:\n{filePath}\nChanges won't be saved\n")
            return
        self.report.df = pd.read_csv(filePath)
        if("sum" in self.report.df.columns):
            self.report.isSummed=True
            if(self.report.df.shape[1]<5):
                self.report.isCompressed=True
        print("Changes activated successfully\n")

    def sumReport(self):
        self.report.df=sumDf(self.report.df)
        self.report.isSummed=True
        print("Changes activated successfully\n")

    def compressReport(self):
        self.report.df=compressDf(self.report.df,self.report.isSummed)
        self.report.isSummed=True
        self.report.isCompressed=True
        print("Changes activated successfully\n")

    def saveReport(self):
        newFilePath=self.filePath()
        try:
            dfToCsv(self.report.df,newFilePath)
            print("Changes activated successfully\n")
        except Exception as e:
            print(f"Could not save, see error below:\n{e}\nChanges won't be saved\n")

    def addBook(self):
        filePath = self.filePath() + ".xlsx"
        if (not os.path.exists(filePath)):
            print(f"File does not exist as:\n{filePath}\nChanges won't be saved\n")
            return
        sheetName=input("Please enter sheet name\n")
        try:
            self.book = pd.read_excel(filePath, sheet_name=sheetName)
            self.book.columns = self.book.columns.str.replace(' ', '')
            print("Changes activated successfully\n")
        except Exception as e:
            print(f"Could not receive book, see error below:\n{e}\nChanges won't be saved\n")

    def saveBook(self):
        newFilePath = self.filePath()
        try:
            dfToXlsx(self.book, newFilePath)
            print("Changes activated successfully\n")
        except Exception as e:
            print(f"Could not save, see error below:\n{e}\nChanges won't be saved\n")

    def match(self):
        print("Match begins. It will take a while")
        self.report.df["used"]=0
        print("Note: report file is altered")
        count = 0
        for index in range(self.book.shape[0]):
            # print(df.iloc[index]["CHE"])
            if (pd.isna(self.book.iloc[index]["קפ"]) and isinstance(self.book.iloc[index]["CHE"], int) and intSize(
                    self.book.iloc[index]["CHE"]) > 3):
                book_c = self.book.iloc[index]["CHE"]
                book_size = intSize(book_c)
                for i in range(self.report.df.shape[0]):
                    if (self.report.df.iloc[i]["used"]==0):
                        other_c = int(self.report.df.iloc[i]["concentration"])
                        other_size = intSize(other_c)
                        if (other_size > book_size):
                            check = str(book_c) in str(other_c)
                        else:
                            check = str(other_c) in str(book_c)
                        if (check):
                            self.book.loc[index, "קפ"] = self.report.df.iloc[i]["sum"]
                            self.report.df.loc[i, "used"] = index+2
                            count += 1
        print(f"Process ended with {count} changes to the book in the system\nYou may save the book to see changes\nYou may save the file to compare changes\n")

    def removeReport(self):
        self.report.df=None
        self.report.isSummed=False
        self.report.isCompressed=False

    def removeBook(self):
        self.book=None

    def changeFolderStatus(self):
        count=0
        changeDict=dict()
        if self.path.getPath()!="":
            count+=1
            print(f"To change to current folder please press {count}")
            changeDict[count]=self.useMainPath
        count+=1
        print(f"To use different custom folder please press {count}")
        changeDict[count]=self.newPath
        count+=1
        print(f"To disable using folder please press {count}")
        changeDict[count]=self.unActivatePath
        change=input()
        try:
            change=int(change)
            if change>=1 and change<=count:
                self.activateFun(changeDict[change])
            else:
                print("No such option exists, changes won't be saved\n")
        except Exception as e:
            print(f"An error occurred, see error below:\n{e}\nChanges won't be saved\n")



    def generateOptoions(self):
        self.funDict=dict()
        lst=[]
        count=0
        if not self.path.usesPath():
            if self.path.getPath()!="":
                count+=1
                lst.append(f"Currently not using folder\nA known one exists:\n{self.path.getPath()}\nFor using the folder please enter {count}")
                self.funDict[count]=self.activatePath
                count+=1
                lst.append(f"For using current folder instead please enter {count}")
                self.funDict[count]=self.useMainPath
            else:
                count+=1
                lst.append(f"You are currently not using any folder\nFor using current folder please enter {count}")
                self.funDict[count]=self.useMainPath
                count+=1
                lst.append(f"For using custom folder please enter {count}")
                self.funDict[count]=self.newPath

        if self.report.df is None:
            count+=1
            lst.append(f"For entering .txt file to the system please enter {count}")
            self.funDict[count]=self.addTxt
            count+=1
            lst.append(f"For entering .pdf file to the system please enter {count}")
            self.funDict[count]=self.addPdf
            count+=1
            lst.append(f"For entering .csv file to the system please enter {count}")
            self.funDict[count]=self.addCsv

        else:
            strAdd="File is loaded\n"
            if(not self.report.isSummed):
                count+=1
                lst.append(f"{strAdd}If you want to add a sum row to the file please enter {count}")
                strAdd=""
                self.funDict[count]=self.sumReport
            if(not self.report.isCompressed):
                count+=1
                lst.append(f"{strAdd}File needs to be compressed in order to match with book\nFor compressing the file please enter {count}")
                strAdd=""
                self.funDict[count]=self.compressReport

            count+=1
            lst.append(f"{strAdd}You can save the file in csv format on your computer\nFor saving the file please enter {count}")
            self.funDict[count]=self.saveReport

        if self.book is None:
            count+=1
            lst.append(f"For entering an Excel book please enter {count}")
            self.funDict[count]=self.addBook
        else:
            count+=1
            lst.append(f"Book sheet is Loaded\nYou can save the sheet in excel format on your computer\nFor saving file please enter {count}")
            self.funDict[count]=self.saveBook

        if(not self.book is None and self.report.isCompressed):
            count+=1
            lst.append(f"All requirements are met for matching the file and the book\nFor starting the process please enter {count}")
            self.funDict[count]=self.match

        if not (self.report.df is None):
            count+=1
            lst.append(f"If you want to remove the file from the system, please enter {count}")
            self.funDict[count]=self.removeReport

        if not (self.book is None):
            count+=1
            lst.append(f"If you want to remove the book from the system, please enter {count}")
            self.funDict[count]=self.removeBook

        if self.path.usesPath():
            currentPath=self.path.getPath()
            if(currentPath==""):
                strAdd=f"Currently saving to current folder:\n{str(pathlib.Path().absolute())}"
            else:
                strAdd=f"Currently saving to known folder:\n{currentPath}"
            count+=1
            lst.append(f"{strAdd}\nFor changing folder status please enter {count}")
            self.funDict[count]=self.changeFolderStatus

        count+=1
        lst.append(f"To quit the program please enter {count}")
        self.funDict[count]=self.quit
        return lst, self.funDict



def main():
    settings = Settings()
    print("Welcome to SideNoter. copyright to dvicofirn")
    while(settings.Keep):
        optionsLst, optionsDict=settings.generateOptoions()
        choise=0
        while not choise in optionsDict.keys():
            print("Please choose one of the options below:")
            for value in optionsLst:
                print(value)
            choise=input()
            try:
                choise=int(choise)
            except Exception as e:
                "Please enter only a number, and only one of the numbers given"
        print()
        settings.activateFun(optionsDict.get(choise))


if __name__ == "__main__":
    main()
