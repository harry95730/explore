from tkinter import *

from tkinter.ttk import *

from tkinter import messagebox

from tkinter.messagebox import showinfo

from tkinter import filedialog

import openpyxl

wins = True

while (wins==True):

    wins = False

    filer=[]

    loc = ""

    loc1 = ""

    count = 0

    def openFile():
        filepath = filedialog.askopenfilename(filetypes=(("text files", "**") ,(".xlsx", "*.xlsx"),(".txt","*.txt"),(".csv","*.csv"),(".py", "*.py")))

        fole = ""

        if filepath != "":

            fole = open(filepath)

            button['state'] = DISABLED

        else:

            messagebox.showerror("ERROR", "PLEASE SELECT A FILE")

        if filepath != "":

            fole = str(fole)

            fole = fole.split("=", 1)[1]

            fole = fole.split(" " , 1)[0]
            x = ""
            for i in fole:
                if i == "/":
                    x = x + "\\"
                elif i == "'":
                    x = x + ""
                else:
                    x = x + i

            global loc
            loc = loc + x

            if(messagebox.askyesno("CONFIRMATION" , loc)):

                wb = openpyxl.load_workbook(loc) 

                sheet = wb.active

                x = str(sheet.max_row)

## write your code here for accessing certain rows and colums and save them to temperory variables 
                
                bolval = True
               
                if(bolval):


                    Label(window, text=" ", font=10).pack()

                    Label(window, text=loc, font=13).pack()

                    Label(window, text="Click the Button to Save a File", font='Aerial 18 bold').pack(pady=20)

                    def save_file():

                        fileo = filedialog.asksaveasfilename(initialfile="untitled", defaultextension=".xlsx",
                                                            filetypes=(("text files", "**") ,(".xlsx", "*.xlsx"),(".txt","*.txt"),(".csv","*.csv"),(".py", "*.py")))

                        if fileo != "":

                            x1=""
                            for i in fileo:
                                if i == "/":
                                    x1 = x1 + "\\"
                                elif i == "'":
                                    x1 = x1 + ""
                                else:
                                    x1 = x1 + i
                            global loc1

                            loc1 = loc1 + x1

                            if (messagebox.askyesno("CONFIRMATION", loc1)):

                                Label(window, text=" ", font=10).pack()

                                Label(window, text=loc1 , font=13).pack()

                                btn['state'] = DISABLED

                                root = Tk()
                                root.geometry('300x120')
                                root.title('PROGRESS BAR')

                                

                                totalthings = 100

                                inital = 0

                                def update_progress_label():
                                    if(inital%100==0):
                                        root.update_idletasks()
                                    return f"Current Progress: {inital}/"+f"{totalthings}"

                                pb = Progressbar(root, orient='horizontal', mode='determinate', length=280)

                                pb.grid(column=0, row=0, columnspan=2, padx=10, pady=20)

                                value_label = Label(root, text=update_progress_label())

                                value_label.grid(column=0, row=1, columnspan=2)

##  write your code here for changes in the excel sheet and use global keyword for accessing the variables   
                                                                                                                                            
                                wb.save(filename=loc1)

                                wb.close()


                                Label(window, text="FILE SAVED SUCCESSFULLY", font='Aerial 22 bold').pack(pady=20)

                                window.update_idletasks()

                                root.destroy()
                                showinfo(message='FILE SAVED SUCCESSFULL!')

                                def run():
                                    global wins
                                    wins = True
                                    window.destroy()
                                but1 = Button(window,text="NEW FILE" , command=run)

                                but1.pack()
                                def run1():
                                    window.destroy()
                                but2 = Button(window, text="EXIT", command=run1)
                                but2.pack()
                                root.mainloop()




                            else :

                                fileo = ""

                                messagebox.showerror("ERROR", "PLEASE SELECT A FILE")


                        else:

                            messagebox.showerror("ERROR", "PLEASE SELECT A FILE NAME")


                    btn = Button(window, text="Save", command=lambda: save_file())

                    btn.pack(pady=10)


                else:

                    button['state']=ACTIVE

            else :

                loc =""

                messagebox.showerror("ERROR", "PLEASE SELECT A FILE")

                button['state'] = ACTIVE


    window = Tk()

    window.title("AUTOMATION FILE")

    window.geometry("800x450")

    Label(window, text="Click the Button to Select a File", font='Aerial 18 bold').pack(pady=20)

    button = Button(text="Open", command=openFile)

    button.pack()

    window.mainloop()




