from tkinter import *
from tkinter.filedialog import askopenfilename
from dbfread import DBF
#import dbf
import xlwt
from win32com.client import constants, Dispatch
import win32com.client
from num2words import num2words


our_file = ''
nameFile = ''
dir_our_file = ''
list_dbf = []

class Root(Tk):
    def __init__(self, parent=None, **options):
        Tk.__init__(self, parent, **options)
        self.geometry('400x150+500+200')
        self.title("Conver dbf")
        ###############################################
        
        self.win = Frame(self)
        self.win.place(x=1, y=2)
        
        self.message = StringVar()
        self.message.set(our_file)
        en = Entry(self.win, textvariable=self.message, width=53, relief=GROOVE, borderwidth=3)
        en.grid(row = 0, column = 0)
        en.bind('<Any-KeyRelease>', self.input_control)
        
        self.btn = Button(self.win, text='Select a file:', command=self.select_file, height=1)
        self.btn.grid(row = 0, column = 1)
        ###############################################
        
        self.lab = Frame(self)
        self.lab.place(x=5, y=50)
        
        Label(self.lab, text='Head of the finance department:').grid(row = 0, column = 0)
        self.mes1 = StringVar()
        en1 = Entry(self.lab, textvariable=self.mes1, width=45, relief=GROOVE, borderwidth=1)
        en1.grid(row = 0, column = 1)  
        
        Label(self.lab, text='Head of Sector:').grid(row = 1, column = 0, sticky=W)
        self.mes2 = StringVar()
        en1 = Entry(self.lab, textvariable=self.mes2, width=45, relief=GROOVE, borderwidth=1)
        en1.grid(row = 1, column = 1) 
        
        Label(self.lab, text='Perpetrator:').grid(row = 2, column = 0, sticky=W)
        self.mes3 = StringVar()
        en1 = Entry(self.lab, textvariable=self.mes3, width=45, relief=GROOVE, borderwidth=1)
        en1.grid(row = 2, column = 1)  
        ###############################################
        
        self.but = Frame(self)
        self.but.place(x=5, y=120)
        
        Button(self.but, text='Convert', command=self.convert, height=1).grid(row = 0, column = 0)
        
    def input_control(self, symbol):
        #so that the string is not changed
        self.message.set(our_file)
        
    def select_file(self):
        global our_file
        our_file = askopenfilename(filetypes = [("DBF file","*.dbf")])
        self.message.set(our_file)
        
    def convert(self):
        self.quit()
        made_convert(self.mes1.get(), self.mes2.get(), self.mes3.get())

def made_convert(mes1, mes2, mes3):
    global dir_our_file
    global nameFile
    dir_our_file = our_file.split('/')
    nameFile = dir_our_file[-1].split('.')[0]
    dir_our_file[-1] = ''
    dir_our_file = '/'.join(dir_our_file)
    
    read_dbf()
    create_excel()
    fill_excel(mes1, mes2, mes3)
    
def read_dbf():
    #a = dbf.Table(r'C:\Users\Максим\Downloads\Максіку\Максіку\первиний файл.DBF', codepage='cp866')
    #DBF(r'C:\Users\Максим\Downloads\Максіку\Максіку\первиний файл.DBF', encoding='cp866'):
    for rec in DBF(our_file, encoding='cp866'):
        list_dbf.append(rec)

def create_excel():
    book = xlwt.Workbook('utf8')
    sheet = book.add_sheet('sheetname')
    book.save(dir_our_file + nameFile + '.xls')

def fill_excel(mes1, mes2, mes3):
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(dir_our_file + nameFile + '.xls')
    xl = Dispatch("Excel.Application")
    sheet = Excel.Sheets(1)
    
    sheet.Cells(1, 1).value = 'Реєстр на виплату заробітної плати/грошового забезпечення \nНаціональна академія СБ України'
    
    
    sheet.Columns("A:A").ColumnWidth = 27.29
    sheet.Columns("B:B").ColumnWidth = 19.71
    sheet.Columns("C:C").ColumnWidth = 36.14
    sheet.Columns("D:D").ColumnWidth = 14.14
    
    sheet.Rows("1:1").RowHeight = 42
    
    sheet.Range("A1:D1").HorizontalAlignment = -4108
    sheet.Range("A1:D1").VerticalAlignment = -4108
    sheet.Range("A1:D1").WrapText = True
    sheet.Range("A1:D1").Orientation = 0
    sheet.Range("A1:D1").AddIndent = False
    sheet.Range("A1:D1").IndentLevel = 0
    sheet.Range("A1:D1").ShrinkToFit = False
    sheet.Range("A1:D1").ReadingOrder = -5002
    sheet.Range("A1:D1").MergeCells = True
    
    sheet.Cells(3, 1).value = "Прізвище ім'я по батькові"
    sheet.Cells(3, 2).value = "ІПН"
    sheet.Cells(3, 3).value = "Рахунок"
    sheet.Cells(3, 4).value = "Сума"
    decor(sheet, r"A3:D3")

    suma = 0
    row = 3
    for employee in list_dbf:
        row += 1
        sheet.Cells(row, 1).value = employee['FIO']
        sheet.Cells(row, 2).value = employee['ID_CODE']
        sheet.Cells(row, 3).value = employee['ACCT_CARD']
        sheet.Cells(row, 4).value = employee['SUMA']
        suma += employee['SUMA']
        decor(sheet, 'A{0}:D{0}'.format(row))
        sheet.Range("A{0}:A{0}".format(row)).HorizontalAlignment = -4131
        sheet.Range("D{0}:D{0}".format(row)).NumberFormat = r"# ##0,00"
        #decor_row(sheet, row)
        
    sheet.Cells(row+1, 1).value = "Всього"
    sheet.Cells(row+1, 4).value = suma
    sheet.Range("D{0}:D{0}".format(row+1)).NumberFormat = r"# ##0,00"
    sheet.Range("A{0}:D{0}".format(row+1)).Font.Bold = True
    
    text = num2words(int(suma), lang='uk') + ' грн. ' + str(suma).split('.')[-1] + ' коп.'
    
    sheet.Cells(row+3, 1).value = chr(ord(text[0])-32) + text[1:]
    
    sheet.Cells(row+5, 1).value = 'Начальник фінансового відділу'
    sheet.Cells(row+5, 3).value = mes1
    
    sheet.Cells(row+6, 1).value = 'Начальник сектору'
    sheet.Cells(row+6, 3).value = mes2
    
    sheet.Cells(row+7, 1).value = 'Виконавець'
    sheet.Cells(row+7, 3).value = mes3
    
    #sheet.PageSetup.PrintArea = 'A1:D{0}'.format(row+10)
    #sheet.VPageBreaks(1).Direction = -4161
    #sheet.VPageBreaks(1).RegionIndex = 1 
    
    sheet.Range("A1:D{0}".format(row+10)).Font.Name = 'Times New Roman'
    sheet.Range("A1:D{0}".format(row+10)).Font.Size = 12
    
    sheet.PageSetup.LeftHeader = ''
    sheet.PageSetup.CenterHeader = '&P'
    sheet.PageSetup.RightHeader = ''
    sheet.PageSetup.LeftFooter = ''
    sheet.PageSetup.CenterFooter = '&F'
    sheet.PageSetup.RightFooter = ''
    sheet.PageSetup.Zoom = 99
    sheet.PageSetup.OddAndEvenPagesHeaderFooter = False
    sheet.PageSetup.DifferentFirstPageHeaderFooter = False
    sheet.PageSetup.ScaleWithDocHeaderFooter = False
    sheet.PageSetup.AlignMarginsHeaderFooter = False
    sheet.PageSetup.EvenPage.LeftHeader.Text = ''
    sheet.PageSetup.EvenPage.CenterHeader.Text = ''
    sheet.PageSetup.EvenPage.RightHeader.Text = ''
    sheet.PageSetup.EvenPage.LeftFooter.Text = ''
    sheet.PageSetup.EvenPage.CenterFooter.Text = ''
    sheet.PageSetup.EvenPage.RightFooter.Text = ''
    sheet.PageSetup.FirstPage.LeftHeader.Text = ''
    sheet.PageSetup.FirstPage.CenterHeader.Text = ''
    sheet.PageSetup.FirstPage.RightHeader.Text = ''
    sheet.PageSetup.FirstPage.LeftFooter.Text = ''
    sheet.PageSetup.FirstPage.CenterFooter.Text = ''
    sheet.PageSetup.FirstPage.RightFooter.Text = ''
    
    wb.Save()
    wb.Close()

    
def decor(sheet, my_range):
    sheet.Range(my_range).HorizontalAlignment = -4108
    sheet.Range(my_range).Borders(5).LineStyle = -4142
    sheet.Range(my_range).Borders(6).LineStyle = -4142
    
    sheet.Range(my_range).Borders(7).LineStyle = 1
    sheet.Range(my_range).Borders(7).ColorIndex = 1
    sheet.Range(my_range).Borders(7).TintAndShade = 0
    sheet.Range(my_range).Borders(7).Weight = 2
    
    sheet.Range(my_range).Borders(8).LineStyle = 1
    sheet.Range(my_range).Borders(8).ColorIndex = 1
    sheet.Range(my_range).Borders(8).TintAndShade = 0
    sheet.Range(my_range).Borders(8).Weight = 2
    
    sheet.Range(my_range).Borders(9).LineStyle = 1
    sheet.Range(my_range).Borders(9).ColorIndex = 1
    sheet.Range(my_range).Borders(9).TintAndShade = 0
    sheet.Range(my_range).Borders(9).Weight = 2
    
    sheet.Range(my_range).Borders(10).LineStyle = 1
    sheet.Range(my_range).Borders(10).ColorIndex = 1
    sheet.Range(my_range).Borders(10).TintAndShade = 0
    sheet.Range(my_range).Borders(10).Weight = 2
    
    sheet.Range(my_range).Borders(11).LineStyle = 1
    sheet.Range(my_range).Borders(11).ColorIndex = 1
    sheet.Range(my_range).Borders(11).TintAndShade = 0
    sheet.Range(my_range).Borders(11).Weight = 2  
    
    sheet.Range(my_range).Borders(12).LineStyle = 1
    sheet.Range(my_range).Borders(12).ColorIndex = 1
    sheet.Range(my_range).Borders(12).TintAndShade = 0
    sheet.Range(my_range).Borders(12).Weight = 2      
    
def decor_row(sheet, row):
    sheet.Range("A{0}:D{0}".format(row)).HorizontalAlignment = -4131
    
Root().mainloop()

print