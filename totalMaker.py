from tkinter import *
from tkinter import filedialog
import pandas as pd
import os
import datetime
import math


class App:
    
    def __init__(self, background):
        
        background.geometry('250x250')
        
        background.configure(background = 'white')
        #background = Label(master, bg = '#ffffff', bd = 0)
        #background.pack(fill = 'both', expand = True)
        
        background.title("Total Maker")
        #master.state('zoomed')
        background.state('iconic')
        
        #rows = 0
        #while rows < 50:
        #    master.rowconfigure(rows, weight=1)
        #    master.columnconfigure(rows,weight=1)
        #    rows += 1
            
        rows = 0
        while rows < 50:
            background.rowconfigure(rows, weight=1)
            background.columnconfigure(rows,weight=1)
            rows += 1
        
        font1 = ("bitstream charter", 25, "bold")
        font2 = ("arial", 20)
        font3 = ("bitstream charter", 55, "bold")
        
        l1 = Label(background, font = font3, text = 'THE TOTAL MAKER', fg = '#1e90ff', bg = 'white')
        
        le1 = Label(background, font = font1, text = 'ENTER MONTH', fg = '#1e90ff', bg = 'white')
        e1 = Entry(background, font = font1, textvariable = month_var, bg = 'white')
        le2 = Label(background, font = font1, text = 'ENTER YEAR', fg = '#1e90ff', bg = 'white')
        e2 = Entry(background, font = font1, textvariable = year_var, bg = 'white')
        
        b1 = Button(background, font = font1, text ="LOAD BIOMETRIC DATA", 
                    fg = 'white', bg = '#00ced1', activeforeground = 'white', 
                    activebackground = '#1e90ff', bd = 0, command = load_file)
        
        load_message_l = Label(background, font = font2,
                               textvariable = load_message, 
                               fg = '#1e90ff', bg = 'white')
        
        b2 = Button(background, font = font1, text ="GENERATE AGGREGATE DATA",
                    fg = 'white', bg = '#00ced1', activeforeground = 'white', 
                    activebackground = '#1e90ff', bd = 0, command = save_file)
        
        save_message_l = Label(background, font = font2,
                               textvariable = save_message, 
                               fg = '#1e90ff', bg = 'white')
        
        exit = Button(background, font = font1, text = "QUIT", 
                      fg = 'white', bg = '#00ced1', bd = 0, 
                      activeforeground = 'white', activebackground = '#1e90ff', 
                      command = background.destroy)
        
        l1.grid(row = 2, column = 7, sticky = W)
        
        le1.grid(row = 6, column = 7, sticky = W)
        e1.grid(row = 6, column = 8, sticky = W)
        le2.grid(row = 8, column = 7, sticky = W)
        e2.grid(row = 8, column = 8, sticky = W)
        
        b1.grid(row = 14, column = 7, sticky = W)
        load_message_l.grid(row = 15, column = 7, sticky = W)
        b2.grid(row = 22, column = 7, sticky = W)
        save_message_l.grid(row = 23, column = 7, sticky = W)
        exit.grid(row = 35, column = 7, sticky = W)
        
        
        
def load_file():
    
    global df
    name = filedialog.askopenfilename(filetypes=[('Excel', ['*.xls', '*.xlsx']), ('CSV', '*.csv',)])

    if name:
        if name.endswith('.csv'):
            df = pd.read_csv(name)
        else:
            df = pd.read_excel(name)
        load_message.set('File loaded successfully!')
        #load_message.set('File loaded successfully!\n(path: %s)' %(name))
    save_message.set('')

          
def save_file():
    
    try:
        month_dict = {'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6, 
                      'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12}
    
        month = month_dict[month_var.get().strip().lower()]
        year = int(year_var.get().strip())

    except KeyError:
        save_message.set('Entered month is not valid!') 
        return
    except ValueError:
        save_message.set('Entered year is not valid!') 
        return  

    try:
        global df
        global new_df
        data = []
        current_book_no = 0
        for i in range(len(df['Unnamed: 1'])):
                if str(df.loc[i]['Unnamed: 1']) != 'nan' and str(df.loc[i]['Unnamed: 1']) != ' - ':
                    current_book_no = df.loc[i]['Unnamed: 1']
                    if current_book_no == 'Branch: Book Na #N/A ':
                        current_book_no = 'Branch: Book No 130 '
                df.at[i, 'Unnamed: 1'] = current_book_no
        
        for i in df['Unnamed: 2']:
            try:
                manno = int(i)
                if len(str(manno)) >= 5:
                    df2 = df.loc[df['Unnamed: 2'] == i]
                    
                    p_days = 0
                    for i in list(df2):
                        try:   
                            if (int(i) >= 0 and int(i) < 32) and (df2[i].values[0] == 'A' or df2[i].values[0] == 'X'):
                                pass
                            else:
                                p_days += 1
                        except:
                                pass
                            
                    sunday_att = 0
                    for i in list(df2):
                        try:
                            if datetime.date(year, month, int(i)).weekday() == 6:
                                if df2[i].values[0] == 'A' or df2[i].values[0] == 'XX':
                                    pass
                                else:
                                    sunday_att += 1
                        except:
                            pass
        
                                        
                    att_list = [df2['Unnamed: 1'].values[0], df2['Unnamed: 2'].values[0],df2['EmpName'].values[0], year, month, 
                                p_days, sunday_att, 0.0,
                                0.0, 0.0, 0.0, 0.0, 0.0,
                                0.0, 0.0, 0.0, 0.0, 0.0]
                    data.append(att_list)
            except ValueError :
                pass

        new_df = pd.DataFrame(data, columns = ['Book No', 'MANNO', 'EmpName', str(year), str(month), 'PHYSICAL_ATT', 'SUNDAY_ATT', 'HOLIDAY_ATT'
                                               'AV_CL', 'AV_PL', 'AV_SL', 'AV_SPL_LV_CODE', 'AV_SPL_LEAVE_DAYS',
                                               'OT_ATT_NORMAL', 'OT_ATT_SUNDAY', 'UG_DAYS', 'NIGHT_ATT', 'CHR_ATT',
                                               'LWP'])
        
        filename = filedialog.asksaveasfilename(defaultextension=".xls") #returns None on pressing cancel in dialogbox
        if filename: 
            new_df.to_excel(filename)
            save_message.set('File saved successfully!')
            #save_message.set('File saved successfully!\n(path: %s)' %(filename))
        load_message.set('')

    except KeyError:
        save_message.set('Data format is not correct!')
        
    except TypeError:
        save_message.set('Data is not loaded!')



root = Tk()

root.attributes("-fullscreen", True)
df = None
new_df = None

load_message = StringVar()
save_message = StringVar()
month_var = StringVar()
year_var = StringVar()

app = App(root)

root.mainloop()
