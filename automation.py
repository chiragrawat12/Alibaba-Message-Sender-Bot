import os
from selenium import webdriver
from tkinter import Tk, Label, Entry, StringVar, Frame, Checkbutton, IntVar, Button, messagebox, END, Scrollbar, Toplevel, Scrollbar, Canvas, ttk
from tkinter.scrolledtext import ScrolledText
from threading import Thread
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import sqlite3
import webbrowser
import openpyxl
import datetime

dir_path = os.path.dirname(os.path.realpath(__file__))
DataBase_File = dir_path + '\Suppliers_DataBase.db'


class DatabaseConnection:
    def __init__(self, host):
        self.connection = None
        self.host = host

    def __enter__(self):
        self.connection = sqlite3.connect(
            self.host, detect_types=sqlite3.PARSE_DECLTYPES)
        return self.connection

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type or exc_val or exc_tb:
            self.connection.close()
        else:
            self.connection.commit()
            self.connection.close()


class login_window(Tk):
    def __init__(self):
        super().__init__()
        self.ws = self.winfo_screenwidth()
        self.hs = self.winfo_screenheight()
        self.height = 330
        self.width = 400
        self.x = int(self.ws)-int(self.width)-25
        self.y = int(self.hs)-int(self.height)-90
        self.geometry("{}x{}+{}+{}".format(self.width,
                                           self.height, self.x, self.y))
        self.update()
        self.title("Alibaba.com")
        self.resizable(width=False, height=False)
        self.iconbitmap(os.path.dirname(
            os.path.realpath(__file__))+"\Image\\send.ico")
        self.run = True
        self.path = os.path.dirname(
            os.path.realpath(__file__)) + '\chromedriver'
        self.options = webdriver.ChromeOptions()
        self.options.add_argument('--ignore-certificate-errors')
        self.options.add_argument('--ignore-ssl-errors')
        self.options.add_experimental_option(
            'excludeSwitches', ['enable-logging'])
        self.driver = webdriver.Chrome(self.path, options=self.options)
        self.driver.maximize_window()
        self.driver.get(
            'https://passport.alibaba.com/icbu_login.htm?spm=a2700.8293689.scGlobalHomeHeader.6.2ce267afnWsnY2&tracelog=hd_signin')

    def on_closing(self):
        self.Show_Data_Button.config(state="normal")
        self.top.destroy()

    def save_excel(self):
        wb = openpyxl.Workbook()
        sheet = wb.active
        with DatabaseConnection(DataBase_File) as connection:
            cursor = connection.cursor()
            cursor.execute("""SELECT * FROM suppliers_contacted""")
            suppliers = cursor.fetchall()
            connection.commit()

        if suppliers == []:
            messagebox.showwarning('No data!', 'Data is not available!')
        else:
            sheet['A1'] = "Suppliers"
            sheet['B1'] = "Suppliers Link"
            sheet['C1'] = "Main Products"
            sheet['D1'] = "Country/Region"

            sheet['A1'].font = openpyxl.styles.Font(bold=True)
            sheet['B1'].font = openpyxl.styles.Font(bold=True)
            sheet['C1'].font = openpyxl.styles.Font(bold=True)
            sheet['D1'].font = openpyxl.styles.Font(bold=True)
            sheet.column_dimensions['A'].width = 55
            sheet.column_dimensions['B'].width = 72
            sheet.column_dimensions['C'].width = 72
            sheet.column_dimensions['D'].width = 30

            for row in suppliers:
                sheet.append(row)
            try:
                wb.save("{}/Excel Files/{}.xlsx".format(
                    os.path.dirname(os.path.realpath(__file__)), datetime.datetime.now().strftime("date_%d_%m_%y__time_%H_%M_%S")))
                messagebox.showinfo('Task Completed!',
                                    'Data is successfully converted to Excel sheet!')
            except:
                pass

    def open_web(self, url):
        webbrowser.open(url, new='new')

    def show_data_func(self):
        self.Show_Data_Button.config(state="disabled")
        self.top = Toplevel()
        self.top.geometry("{}x{}".format(1100, 550))
        self.top.wm_title("Suppliers Data")
        self.top.grab_set()
        self.top.minsize(1118, 550)
        self.top.iconbitmap(os.path.dirname(
            os.path.realpath(__file__))+"\Image\\send.ico")
        self.top_frame = Frame(self.top)
        self.top_frame.pack(fill="both", expand=True)

        self.canvas = Canvas(self.top_frame)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scroll_data = ttk.Scrollbar(
            self.top_frame, orient="vertical", command=self.canvas.yview)
        self.scroll_data.pack(side="right", fill="y")

        self.canvas.configure(yscrollcommand=self.scroll_data.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))

        self.second_frame = Frame(self.canvas)
        self.canvas.create_window(
            (0, 0), window=self.second_frame, anchor="nw")

        self.scroll_data.config(command=self.canvas.yview)
        Label(self.second_frame,  width=45, text="Suppliers", bd=2, relief="raised", font=(
            "helvatica", 10, "bold")).grid(row=0, column=0)
        Label(self.second_frame,  width=65, text="Main Products", bd=2, relief="raised", font=(
            "helvatica", 10, "bold")).grid(row=0, column=1)
        Label(self.second_frame,  width=25, text="Country/Region", bd=2, relief="raised",
              font=("helvatica", 10, "bold")).grid(row=0, column=2)

        with DatabaseConnection(DataBase_File) as connection:
            cursor = connection.cursor()
            cursor.execute("""SELECT * FROM suppliers_contacted""")
            suppliers = cursor.fetchall()
            connection.commit()

        rows = len(suppliers)
        for i in range(rows):
            Button(self.second_frame, width=51, cursor="hand2", text=str(i+1)+"."+suppliers[i][0], anchor="w", bd=2, relief="groove", font=(
                "helvatica", 9), command=lambda x=suppliers[i][1]: self.open_web(x)).grid(row=i+1, column=0)

            Label(self.second_frame, width=65, text=suppliers[i][2], anchor="w", bd=2, relief="groove", font=(
                "helvatica", 10)).grid(row=i+1, column=1)
            Label(self.second_frame,  width=25, text=suppliers[i][3], bd=2, relief="groove",
                  font=("helvatica", 10)).grid(row=i+1, column=2)
        self.top.protocol('WM_DELETE_WINDOW', self.on_closing)

    def stop(self):
        stop = messagebox.askyesno("Are you sure?", "Do you want to stop?")
        if stop:
            self.run = False
            self.Start_Button.config(state="normal")
            self. Stop_Button.config(state="disabled")
            self.link_entry.configure(state="normal")
            self.message_box.configure(state="normal")
            self.Clear_Button.configure(state="normal")

    def Create_Database(self):
        with DatabaseConnection(DataBase_File) as connection:
            cursor = connection.cursor()
            cursor.execute("""CREATE TABLE IF NOT EXISTS suppliers_contacted(
                        suplier text NOT NULL,
                        supplier_link text NOT NULL,
                        main_products text NOT NULL,
                        country_region text NOT NULL);""")
            connection.commit()

    def main(self):
        with DatabaseConnection(DataBase_File) as connection:
            cursor = connection.cursor()
            cursor.execute("""DELETE From suppliers_contacted""")
            connection.commit()

        while(self.run == True):
            if not self.run:
                break
            try:
                wait = WebDriverWait(self.driver, 4).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "wait_here"))
                )
            except:
                pass

            self.driver.execute_script(
                "window.scrollTo(0,document.body.scrollHeight)")
            try:
                wait = WebDriverWait(self.driver, 4).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "wait_here_again"))
                )
            except:
                pass
            self.driver.execute_script(
                "window.scrollTo(0,document.body.scrollHeight)")
            try:
                contact_btn = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "//a[@class='button csp']"))
                )
            except:
                pass

            try:
                supplier_name = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "//div[@class='title-wrap']/h2[@class='title ellipsis']/a"))
                )
            except:
                pass

            try:
                main_products = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "//div[@class='right']/div[@class='attrs']/div[@class='attr']/div[@class='value ellipsis ph']"))
                )
            except:
                pass

            try:
                country_region = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "//div[@class='right']/div[@class='attrs']/div[@class='attr']/div[@class='value']/img[@class='flag']/following-sibling::span"))
                )
            except:
                pass

            try:
                wait = WebDriverWait(self.driver, 2).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "wait_here_again_for_none"))
                )
            except:
                pass
            self.driver.execute_script(
                "window.scrollTo(document.body.scrollHeight,0)")
            try:
                wait = WebDriverWait(self.driver, 2).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "wait_here_again_for_none"))
                )
            except:
                pass

            for contact, supp_name, supp_prod, supp_country in zip(contact_btn, supplier_name, main_products, country_region):
                if not self.run:
                    break
                with DatabaseConnection(DataBase_File) as connection:
                    cursor = connection.cursor()
                    cursor.execute("""INSERT INTO suppliers_contacted VALUES(?,?,?,?)""",
                                   (supp_name.text, supp_name.get_attribute('href'), supp_prod.text, supp_country.text))
                    connection.commit()

                contact.click()
                main_window = self.driver.window_handles
                self.driver.switch_to.window(main_window[1])
                try:
                    message = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located(
                            (By.ID, 'inquiry-content'))
                    )
                    message.send_keys(self.message_box.get(1.0, END))
                except:
                    pass

                try:
                    send_msg_btn = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//input[@type='submit']"))
                    )
                    send_msg_btn.click()
                except:
                    pass

                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//div[@class='ui2-feedback-title']"))
                    )
                except:
                    pass
                self.driver.close()
                self.driver.switch_to.window(main_window[0])
                try:
                    wait = WebDriverWait(self.driver, 2).until(
                        EC.presence_of_all_elements_located(
                            (By.XPATH, "wait_here_156"))
                    )
                except:
                    pass
            if self.run:
                try:
                    wait = WebDriverWait(self.driver, 4).until(
                        EC.presence_of_all_elements_located(
                            (By.XPATH, "wait_here_1"))
                    )
                except:
                    pass
                self.driver.execute_script(
                    "window.scrollTo(0,document.body.scrollHeight)")
                try:
                    wait = WebDriverWait(self.driver, 2).until(
                        EC.presence_of_all_elements_located(
                            (By.XPATH, "wait_here_123"))
                    )
                except:
                    pass
                try:
                    self.driver.execute_script(
                        "document.getElementById('J-m-pagination').scrollIntoView();")
                except:
                    pass

                try:
                    Next_btn = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//a[@class='next']"))
                    )
                    Next_btn.click()
                except:
                    self.run = False
        self.Start_Button.config(state="normal")
        self. Stop_Button.config(state="disabled")
        self.link_entry.configure(state="normal")
        self.message_box.configure(state="normal")
        self.Clear_Button.configure(state="normal")

    def start_thread_func(self):
        self.start_thread = Thread(target=self.start)
        self.start_thread.start()

    def stop_thread_func(self):
        self.stop_thread = Thread(target=self.stop)
        self.stop_thread.start()

    def clear(self):
        self.message_box.delete(1.0, END)

    def start(self):
        self.run = True
        self.link_entry.configure(state="disable")
        self.message_box.config(state="disabled")
        self.Clear_Button.configure(state="disabled")
        self.Start_Button.config(state="disabled")
        self.Stop_Button.config(state="normal")
        if self.link.get() == "":
            messagebox.showwarning(
                'Empty Fields!', 'You have to fill all the fields.')
            self.Start_Button.config(state="normal")
            self. Stop_Button.config(state="disabled")
            self.link_entry.configure(state="normal")
            self.message_box.configure(state="normal")
            self.Clear_Button.configure(state="normal")

        else:
            try:
                self.driver.get(self.link.get())
                self.main()
            except:
                try:
                    self.driver = webdriver.Chrome(
                        self.path, options=self.options)
                    self.driver.maximize_window()
                    self.driver.get(self.link.get())
                    self.main()
                except:
                    self.driver.quit()
                    messagebox.showwarning(
                        'Invalid Field!', 'You have to Enter the valid Url.')
                    self.Start_Button.config(state="normal")
                    self. Stop_Button.config(state="disabled")
                    self.link_entry.configure(state="normal")
                    self.message_box.configure(state="normal")
                    self.Clear_Button.configure(state="normal")

    def scrape_window(self):
        self.link_label = Label(
            self, text="Link:", font=("helvatica", 10, "bold"))
        self.link = StringVar()
        self.link_label.pack(anchor='w', padx=27)
        self.link_entry = Entry(
            self, textvariable=self.link, width=45, font=("helvatica", 9,), bd=3)
        self.link_entry.pack(padx=29, anchor='w')

        self.message_box_label = Label(
            self, text="Message:", font=("helvatica", 9, "bold"))
        self.message_box_label.pack(anchor='w', padx=27)

        self.message_box = ScrolledText(
            self, width=35, height=6, font=("helvatica", 12,), bd=3)
        self.message_box.pack()

        self.Clear_Button = Button(self, text="Clear All", fg="white",
                                   relief="raised", bg="#6600ff", bd=2,
                                   font=("helvetica", 9, "bold"),
                                   activeforeground="#6600ff", cursor="hand2", activebackground="#b489f5", command=self.clear)
        self.Clear_Button.pack(padx=45, anchor='e', pady=5)

        self.Start_Button = Button(self, text="Start", fg="white",
                                   relief="raised", bg="#0c8f00", bd=2,
                                   font=("helvetica", 9, "bold"),
                                   activeforeground="#0c8f00", activebackground="#15ff00", cursor="hand2", command=self.start_thread_func)
        self.Start_Button.pack(fill="x", pady=5, padx=20)

        self.Stop_Button = Button(self, text="Stop", fg="white",
                                  relief="raised", bg="#b50000", bd=2,
                                  font=("helvetica", 9, "bold"),
                                  state='disabled', activeforeground="#b50000", cursor="hand2", activebackground="red", command=self.stop_thread_func)
        self.Stop_Button.pack(fill="x", pady=5, padx=20)

        self.show_frame = Frame(self)
        self.show_frame.pack(expand=1, fill='x')

        self.Show_Data_Button = Button(self.show_frame, text="Show Data", fg="white",
                                       relief="raised", bg="#ba9600", bd=2,
                                       font=("helvetica", 9, "bold"),
                                       activeforeground="#ba9600", activebackground="#fce06d", cursor="hand2", command=self.show_data_func)
        self.Show_Data_Button.pack(padx=20, fill='x', side="left", expand=1)
        self.Excel_Button = Button(self.show_frame, text="Convert to Excel", fg="white",
                                   relief="raised", bg="#0035ba", bd=2, cursor="hand2",
                                   font=("helvetica", 9, "bold"),
                                   activeforeground="#0035ba", activebackground="#335cff", command=self.save_excel)
        self.Excel_Button.pack(padx=20, side="left", fill="x", expand=1)


if __name__ == "__main__":
    login = login_window()
    login.Create_Database()
    login.scrape_window()

    def cross():
        quit = messagebox.askyesno('Are you sure?', 'Do you want to quit?')
        if quit:
            login.destroy()
    login.protocol('WM_DELETE_WINDOW', cross)
    login.mainloop()
