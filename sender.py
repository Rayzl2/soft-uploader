import csv
import xlrd
import datetime
import clickhouse_connect
from datetime import date
from tkinter import *
from tkinter import filedialog
import ctypes as ct





cols = ['date_dep',
    'month_dep',
    'vagon_n',
    'naklad_n',
    'container_n',
    'prefix',
    'type_c',
    'type_spec',
    'gov_dep',
    'gov_arr',
    'st_dep',
    'priznak_dep',
    'zdpp_dep',
    'st_arr',
    'priznak_arr',
    'zdpp_arr',
    'dor_dep',
    'mtu_dep',
    'dor_arr',
    'mtu_arr',
    'region_dep',
    'region_arr',
    'type_per',
    'gruz',    
    'subgruz',   
    'grgruz',
    'category',
    'parktype',
    'sender',
    'reciever',
    'owner',
    'renter',
    'operator',
    'payer',
    'payer_prev',
    'owner_sprav',
    'sends',
    'weight',
    'teu'
    ]


EPOCH_START = date(1970, 1, 1)


def get_count_days(dt: date) -> int:
    return (dt - EPOCH_START).days


def get_time():
    now = datetime.datetime.now()

    return now.strftime("%d-%m-%Y %H:%M:%S")


def csv_from_excel():
    try:
        print("\t[INIT]\tПроизводится инициализация таблицы Excel\t" + get_time())
        print(file_path)
        sh = xlrd.open_workbook(file_path).sheet_by_name('БАЗА')
        print("\t[LOG]\tИнициадизация окончена\t" + get_time())
        print("\t[LOG]\tПроизводится конвертация в .CSV\t" + get_time())
        csv_active = open('Base.csv', 'w')
        wr = csv.writer(csv_active, quoting=csv.QUOTE_ALL)

        for rownum in range(1, sh.nrows):
            wr.writerow(sh.row_values(rownum))


        print("\t[LOG]\tОкончена конвертация в .CSV\t" + get_time())
        print("\t[LOG]\tЧтение .CSV файла\t" + get_time())
        csv_active.close()

    except Exception as e:
        print("!\t[ERROR]")
        print(e)


def reading():
    try:
        csv_active = open('Base.csv', 'r')
        inner = csv.reader(csv_active, delimiter=',')
        return inner, csv_active

    except Exception as e:
        print("!\t[ERROR]")
        print(e)


def clickhouse(client, data):
    client.insert('analytic.rzd', data)


def main():
        print("\t[LOG]\tРабота начата\t" + get_time())
        
        csv_from_excel()
        inner, file = reading()

        print("\t[LOG]\tВсе данные успешно получены\t" + get_time())
        client = clickhouse_connect.get_client(host='rc1b-***.mdb.yandexcloud.net', port=8443, username='analytic', password='***')
        print("\t[CONNECT ]\tПодключение к ClickHouse выполнено\t" + get_time())
        print("\t[LOG]\tНачата передача данных в ClickHouse\t" + get_time())
        dataSet = []
        for row in inner:
            if len(dataSet) == 500:
                clickhouse(client, dataSet)
                dataSet = []

            line = []
            cnt = 1
            ID = ''
            for l in row:
                if cnt == 1:
                    try:
                        datetime_date = xlrd.xldate_as_datetime(int(float(l)), 0)
                        l = datetime_date.date()
                        l = get_count_days(l)

                        l = int(l)
                    
                    except:
                        l = 0
                
                elif cnt == 2:
                    l = 1

                elif cnt == 37 or cnt == 38 or cnt == 39:
                    try:
                        l = int(float(l))
                    except:
                        l = 100

                else:
                    l = str(l.replace('"', ''))
                
                if l != '':
                    line.append(l)

                cnt+=1
            
            ID = ID + str(l)

            if len(line) == 39:
                
                dataSet.append(line)

        file.close()

        print("\t[LOG]\tПередача успешно окончена\t" + get_time())
        print(cnt)


def dark_title_bar(window):
    """
    MORE INFO:
    https://learn.microsoft.com/en-us/windows/win32/api/dwmapi/ne-dwmapi-dwmwindowattribute
    """
    window.update()
    set_window_attribute = ct.windll.dwmapi.DwmSetWindowAttribute
    get_parent = ct.windll.user32.GetParent
    hwnd = get_parent(window.winfo_id())
    value = 2
    value = ct.c_int(value)
    set_window_attribute(hwnd, 20, ct.byref(value),
                         4)


def choose():
    global file_path
    file_path = filedialog.askopenfilename()
    lb2 = Label(window, text=f"Выбранный файл:\n{file_path}", font=("Arial Bold", 12),fg='#fff', background='#3d3d3d')
    lb2.place(rely=0.35, relx=0.5, anchor=CENTER)

window = Tk()
window.geometry('720x360')
window.resizable(width=False, height=False)
window.configure(background='#3d3d3d')
dark_title_bar(window)
window.title("Software for uploading Excel files")
lbl = Label(window, text="Выберите файл для загрузки", font=("Arial Bold", 32),fg="#fff", background='#3d3d3d')
btn = Button(window, text='Выбрать файл', command=choose,fg='#fff', background='#17c0eb', borderwidth=0, width=20, height=3) 
lb2 = Label(window, text=f"Выбранный файл: Н/Д", font=("Arial Bold", 12),fg='#fff', background='#3d3d3d')

lbl.place(rely=0.2, relx=0.5, anchor=CENTER)
lb2.place(rely=0.35, relx=0.5, anchor=CENTER)
btn.place(relx=0.5, rely=0.6, anchor=CENTER)

btn1 = Button(window, text='Загрузить в DataLens', command=main, width=50, height=5,fg='#fff' ,background='#17c0eb', borderwidth=0) 
btn1.place(relx=0.5, rely=0.85, anchor=CENTER)

window.mainloop()

# SOURCE OF ALGORITM 

'''
def clickhouse(client, dataSet):
    client.insert('rzd', dataSet, column_names=cols)


def main():
    try:
        print("\t[LOG]\tРабота начата\t" + get_time())
        
        #csv_from_excel()
        inner, file = reading()

        print("\t[LOG]\tВсе данные успешно получены\t" + get_time())
        client = clickhouse_connect.get_client(host='....9.mdb.yandexcloud.net', port=8443, username='analytic', password='....', database='analytic')
        print("\t[CONNECT]\tПодключение к ClickHouse выполнено\t" + get_time())
        print("\t[LOG]\tНачата передача данных в ClickHouse\t" + get_time())
        dataSet = []
        for row in inner:
            line = []
            for l in row:
                l = str(l.replace('"', ''))
                line.append(l)
            dataSet.append(line)

        clickhouse(client, dataSet)
        file.close()

    except Exception as e:
        print("!\t[ERROR]")
        print(e)
'''
