from tkinter import *
from tkinter.filedialog import askopenfilename
from dbfread import DBF
from num2words import num2words
import xlsxwriter


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
    #nameFile = dir_our_file[-1].split('.')[0]
    nameFile = dir_our_file[-1].split('.')
    nameFile = '.'.join(nameFile[:len(nameFile)-1])
    dir_our_file[-1] = ''
    dir_our_file = '/'.join(dir_our_file)
    
    read_dbf()
    fill_excel(mes1, mes2, mes3)
    
def read_dbf():
    #a = dbf.Table(r'C:\Users\Максим\Downloads\Максіку\Максіку\первиний файл.DBF', codepage='cp866')
    #DBF(r'C:\Users\Максим\Downloads\Максіку\Максіку\первиний файл.DBF', encoding='cp866'):
    for rec in DBF(our_file, encoding='cp866'):
        list_dbf.append(rec)

def fill_excel(mes1, mes2, mes3):
    workbook = xlsxwriter.Workbook(dir_our_file + nameFile + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column(0, 0, 21.30)
    worksheet.set_column(1, 1, 12.15)
    worksheet.set_column(2, 2, 35.86)
    worksheet.set_column(3, 3, 14.50)
    worksheet.set_row(0, 42)
    
    header_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 14, 'text_wrap': True, 'align': 'center', 'valign': 'center', 'bold': True})
    table_cell = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'align': 'center', 'valign': 'center', 'border': 1})
    format_worker = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'align': 'left', 'border': 1})
    table_sum_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'align': 'center', 'valign': 'center', 'border': 1, 'num_format': '#,##0.00'})
    total_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'align': 'left', 'valign': 'center', 'border': 1, 'bold': True})
    format_summ_total = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'align': 'center', 'valign': 'center', 'border': 1, 'num_format': '#,##0.00', 'bold': True})
    format_bottom_line = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 14,'align': 'center', 'valign': 'center', 'bottom': 1})
    
    text_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12})
    
    worksheet.merge_range(0, 0, 0, 3, "Реєстр на виплату заробітної плати/грошового забезпечення \nНаціональна академія СБ України", header_format)
    
    worksheet.write('A3', "Прізвище, ініціали", table_cell)
    worksheet.write('B3', "ІПН", table_cell)
    worksheet.write('C3', "Рахунок", table_cell)
    worksheet.write('D3', "Сума", table_cell)
    
    suma = 0
    row = 2
    for employee in list_dbf:
        row += 1
        worksheet.write(row, 0,  employee['FIO'], format_worker)
        worksheet.write(row, 1,  employee['ID_KOD'], table_cell)
        worksheet.write(row, 2,  employee['NSC'], table_cell)
        worksheet.write(row, 3,  employee['SUMMA'], table_sum_format)
        suma += employee['SUMMA']
    
    suma = round(suma, 2)
        
    worksheet.write(row + 1, 0,  'Всього', total_format) 
    worksheet.write(row + 1, 1, None , table_cell) 
    worksheet.write(row + 1, 2, None , table_cell) 
    worksheet.write(row + 1, 3, suma , format_summ_total) 
    
    text = num2words(int(suma), lang='uk') + ' грн. ' + str(suma).split('.')[-1] + ' коп.'
    worksheet.write(row + 3, 0, chr(ord(text[0])-32) + text[1:] , text_format) 
    
    worksheet.write(row + 5, 0,  'Начальник фінансового відділу', text_format)
    worksheet.write(row + 6, 0,  'Начальник сектору', text_format)
    worksheet.write(row + 7, 0,  'Виконавець', text_format)
    
    worksheet.set_row(row + 5, 16.5)
    worksheet.set_row(row + 6, 16.5)
    worksheet.set_row(row + 7, 16.5)
    
    worksheet.write(row + 5, 2,  None, format_bottom_line)
    worksheet.write(row + 6, 2,  None, format_bottom_line)
    worksheet.write(row + 7, 2,  None, format_bottom_line)
    
    worksheet.write(row + 5, 3,  mes1, text_format)
    worksheet.write(row + 6, 3,  mes2, text_format)
    worksheet.write(row + 7, 3,  mes3, text_format)
    
    workbook.close()
    
Root().mainloop()