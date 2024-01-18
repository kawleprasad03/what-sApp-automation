from tkinter import *
class whatsapp:
    global excelFile
    global columnName
    global ID
    def send(self):
        from pandas import options
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        import time
        # from config import chrome_data_path
        import os
        import pandas as pd
        global fi
        global n
        with open("utill/cache.txt","r") as fi:
            print(fi.read())
            print("###")
        options=webdriver.ChromeOptions()
        options.add_argument(str(fi))

        # pa = "D:\selenium_file\details.xlsx"
        
        df=pd.read_excel(excelFile)
        listofnumber=[]
        length=len(df[columnName])
        for i in range(0,length):
            listofnumber.append(str(df[columnName][i]))

        a=len(listofnumber)
        c=0
        driver=webdriver.Chrome(executable_path="D:\chromedriver_win32\chromedriver.exe", options=options)
        def document():
            attach = WebDriverWait(driver,90).until(EC.presence_of_element_located((By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/div/span')))
            attach.click()
            #driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()
            doc = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[1]/li')))
            doc.send_keys(e3.get())
            #driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[4]/button/input').send_keys(pa)
            # time.sleep(1)
            send = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span')))
            send.click()                                                                                 
            return None
        def image():
            
            attachButton = WebDriverWait(driver,90).until(EC.presence_of_element_located((By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/div/span')))
            attachButton.click()  #xpath is attach button of whatsapp page  
            image = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div/span')))
            image.send_keys(e3.get())    #xpath is photo and video button of whatsapp page 
            image.click()
            # time.sleep(1) 
    
            send = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div')))
            send.click()
            return None

        for j in listofnumber:
            c+=1
            driver.get(f"https://api.whatsapp.com/send/?phone=91{j}")
            driver.maximize_window()
            driver.find_element(By.ID,"action-button").click()
            link=WebDriverWait(driver,3).until(EC.presence_of_element_located((By.LINK_TEXT,"use WhatsApp Web")))
            link.click()

            if ID==1:
                image()
            elif ID==2:
                document()
    

            print(c)
            time.sleep(3)
            # driver.refresh()
            # driver.back()
            if c!=a:
                continue
            else:
                dot=WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[4]/div/span')))   #'//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[3]/div/span'
                dot.click()                                                                 #//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[4]/div/span
                # driver.find_element(By.XPATH,'//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[3]/div/span').click()
                logout1=WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[4]/span/div/ul/li[5]/div')))       #//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[3]/span/div/ul/li[4]/div[1]
                logout1.click()
                # driver.find_element(By.XPATH,'//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[3]/span/div/ul/li[4]/div[1]').click()
                logout2=WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/div/span[2]/div/div/div/div/div/div/div[3]/div/div[2]/div/div')))              #//*[@id="app"]/div/span[2]/div/div/div/div/div/div/div[3]/div/div[2]/div/div
                logout2.click()
                time.sleep(3)
                driver.quit()     


    def gui(self):
        from tkinter import ttk,filedialog
        from tkinter.filedialog import askopenfilename
        from tkinter import scrolledtext
        from tkinter import messagebox
        root = Tk() 

        root.title("Automate Whatsapp message")

        root.geometry("920x500+100+100")
        f = Frame(root,width=600,height=550)

        
        filepath2 = StringVar()
        def columnvalues():         #function for getting all columns names stored it into combobox
            import pandas as pd
            global excelFile
            excelFile = str(filepath2.get())
            df = pd.read_excel(str(filepath2.get()))
            list1 = list(df.columns)
            return list1

        def open_file():   # function for browse button to browse file
            fi2 = filedialog.askopenfilename(initialdir='/',title='Select cache path',filetypes=(('excel files','*.xlsx'),('all files','*.*')))
            filepath2.set(fi2)
            list1 = columnvalues()
            # print(list1)
            # print(tuple(list1))
            comb(list1)
                
        l2 = Label(f,text='Enter your excel file path:',font=14)
        e2 = Entry(f,font=14,width=35,textvariable=filepath2)     # input for excel file

        style = ttk.Style()
        style.theme_use('alt')
        style.configure('TButton',background='#00b4d8')
        brow2 = Button(f,text='Browse',width=8,command=open_file,font=14,bg='#00b4d8',relief='raised')  #ttk.Button(f,text='Browse',command=open_file)   # browse button 

        l3 = Label(f,text='Select your Number column :',font=14)
        # comobobox
        global n
        n = StringVar()
        combo1 = ttk.Combobox(f,font=14,textvariable=n) 
        combo1.grid(row=1,column=1,pady=20,sticky='W')
        combo1.config(state='disabled')
        def comb(list1):     # function for all columns name display in combobox after browse excel file
            # print(list1)
            combo1['values'] = tuple(list1)
            combo1.config(state='eabled')

        def value():   # function call when particular checkbutton is click
            def image():        # this function display gui when image checkbutton is click
                global l5
                global e3
                global brow3
                l5 = Label(f,text='Select path of your Image:',font=14)
                global filepath3 
                filepath3 = StringVar()
                def open_image():
                    im = filedialog.askopenfilename(initialdir='/',title='select image file',filetypes=(('jpg','*.jpg'),('jpeg','*.jpeg'),('png','*.png'),('all image','*.*')))
                    filepath3.set(im)

                e3 = Entry(f,font=14,width=35,textvariable=filepath3)
                brow3 = Button(f,text='Browse',command=open_image,width=8,font=14,bg='#00b4d8',relief='raised')

                l5.grid(row=4,column=0,pady=20)
                e3.grid(row=4,column=1,pady=20,sticky='W')
                brow3.grid(row=4,column=1,columnspan=2,sticky='E')


            def document(): # this function display gui when document checkbutton is click
                global l5
                global e3
                global brow3
                l5 = Label(f,text='Select document path:',font=14)
                global filepath3
                filepath3 = StringVar()
                def open_document():
                    fi3 = filedialog.askopenfilename(initialdir='/',title='select document',filetypes=(('excel file','*.xlsx'),('pdf','*.pdf'),('doc file','*.docx'),('text file','*.txt'),('ppt','*.pptx'),('all file','*.*')))
                    filepath3.set(fi3)

                e3 = Entry(f,font=14,width=35,textvariable=filepath3)
                brow3 = Button(f,text='Browse',font=14,width=8,bg='#00b4d8',relief='raised',command=open_document)
                l5.grid(row=4,column=0,pady=20)
                e3.grid(row=4,column=1,pady=20,sticky='W')
                brow3.grid(row=4,column=1,columnspan=2,sticky='E')

            global ID
            ID = s.get()
            if s.get()==1:     # call image function when checkbutton value is 1
                image()  
                c2.config(state='disabled')
            elif s.get()==2:    # call document function when checkbutton value is 2
                document()
                c1.config(state='disabled')

        def cancel():   # this function reset the checkbutton and destory particular checkbutton gui
            if s.get()==1 or s.get()==3:   
                s.set(0)
                l5.destroy()
                e3.destroy()
                brow3.destroy()
                c2.config(state='normal')
            elif s.get()==2 or s.get()==4:
                s.set(0)
                l5.destroy()
                e3.destroy()
                brow3.destroy()
                c1.config(state='normal')

        def confirm():              # this function for validation 
            global columnName
            columnName = str(n.get())
            if len(filepath2.get()) == 0:
                messagebox.showerror(title='Warning',message='select porper attachment')
            elif len(n.get()) == 0:
                messagebox.showerror(title='Warning',message='select correct option')
            elif len(e3.get()) == 0:
                messagebox.showerror(title='Warning',message='select porper attachment')
            else:
                e2.config(state='disabled')
                brow2.config(state='disabled')
                combo1.config(state='disabled')
                t1.config(state='disabled')
                c1.config(state='disabled')
                c2.config(state='disabled')
                b1.config(state='disabled')
                e3.config(state='disabled')
                brow3.config(state='disabled')
        def reset():
            e2.config(state='normal')
            brow2.config(state='normal')
            t1.config(state='normal')
            c1.config(state='normal')
            c2.config(state='normal')
            b1.config(state='normal')
            e3.config(state='normal')
            brow3.config(state='normal')
            filepath2.set('')
            combo1.set('')
            t1.delete("1.0","end")
            s.set(0)
            filepath3.set('')
            l5.destroy()
            e3.destroy()
            brow3.destroy()



            
     
        l4 = Label(f,text="Enter message to send all contacts in excel",font=14)
        t1 = scrolledtext.ScrolledText(f,font=14,width=50,height=5)  

        l6 = Label(f,text='Select what you want to send in whatsapp:',font=14,padx=10)
        s = IntVar()
        c1 = Checkbutton(f,text='Image',variable=s,font=14,onvalue=1,offvalue=3,command=value)
        c2 = Checkbutton(f,text='Document',variable=s,font=14,onvalue=2,offvalue=4,command=value)
        b1 = Button(f,text='Cancel',width=8,bg='#00b4d8',font=14,command=cancel)    

        btn1 = Button(f,text='Confirm',font=14,width=8,justify=RIGHT,bg='#00b4d8',command=confirm)
        btn2 = Button(f,text='Reset',font=14,width=8,justify=RIGHT,bg='#00b4d8',command=reset)
        btn3 = Button(f,text='Send',font=14,width=8,justify=RIGHT,bg='#00b4d8',command=self.send)  # start send message button
        l2.grid(row=0,column=0,pady=20)
        e2.grid(row=0,column=1,pady=20,sticky='W')
        brow2.grid(row=0,column=1,columnspan=2,pady=20,sticky='E')

        l3.grid(row=1,column=0,pady=20)

        l4.grid(row=2,column=0,pady=20)
        t1.grid(row=2,column=1,pady=20)


        l6.grid(row=3,column=0,pady=20)
        c1.grid(row=3,column=1,pady=20,sticky='W')
        c2.grid(row=3,column=1,pady=20,columnspan=2,sticky='W',padx=100)
        b1.grid(row=3,column=1,columnspan=3,pady=20,sticky='W',padx=220)

        btn1.grid(row=5,column=0,pady=20,sticky='E')   # confirm
        btn2.grid(row=5,column=1,pady=20)                # reset
        btn3.grid(row=5,column=2,sticky='W',pady=20)    # send

        f.pack()            

        root.mainloop()
    
    def file(self):    # file function for getting cache path with help of gui and stored it in text file
        import os
        if os.path.exists("utill/cache.txt"):
            with open("utill/cache.txt","r") as fi:
                print(fi.read())
                print("###")
                return None
        from tkinter import filedialog
        from tkinter import messagebox
        root = Tk()
        
        global filepath1 
        filepath1 = StringVar()
        def getDirectory():    # this function browse cache path
            fi1 = filedialog.askdirectory(initialdir='/',title='dialog box')
            filepath1.set(fi1)
        def getFilePath():    # this function write cache path in text file
            dirIsExists = os.path.exists("utill")
            isExists = os.path.exists("utill/cache.txt")
            global fi
            if dirIsExists==False and isExists==False:
                os.mkdir("utill") 
                with open("utill/cache.txt","w") as fi:
                    path = filepath1.get()
                    fi.write("user-data-dir="+path)
                
            elif isExists==False:
                fi = open("utill/cache.txt","w")
                with open("utill/cache.txt","w") as fi:
                    path = filepath1.get()
                    fi.write("user-data-dir="+path)
            # else:
            #     fi = open("utill/cache.txt","r")
            #     print(fi.read())
            root.destroy()
            
        def check():
            if len(e1.get()) == 0:
                messagebox.showerror(title='Warning',message='select proper attachment')
            else:
                getFilePath()
                
        root.geometry("700x110+150+180")
        root.attributes('-topmost','true')
        l1 = Label(root,text='Enter your chrome cache path:',font=14)
        e1 = Entry(root,font=14,width=35,textvariable=filepath1)
        brow1 = Button(root,text='Browse',command=getDirectory,width=8,bg='#00b4d8',relief='raised',bd=2)

        sub = Button(root,text='submit',command=check,bg='#00b4d8',relief='raised',bd=2)
        l1.grid(row=0,column=0,pady=20)
        e1.grid(row=0,column=1,pady=20,sticky='W')
        brow1.grid(row=0,column=2,pady=20,padx=10)
        sub.grid(row=1,column=1,pady=10)
        root.mainloop()
        # self.gui()


wa = whatsapp()
wa.file()
wa.gui()