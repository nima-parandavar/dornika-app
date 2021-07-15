from openpyxl import *
from tkinter import messagebox, ttk, IntVar, Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename


class RepeatedQr:
    def __init__(self, main_file, second_file):

        self.main_file = main_file
        self.sec_file = second_file


    def load_columns(self, main, second, m1, m2, m3, m4, m5, m6, s1, s2, s3, s4):
        self.main_columns = {
            "main": main.upper(),
            "m1": m1.upper(),
            "m2": m2.upper(),
            "m3": m3.upper(),
            "m4": m4.upper(),
            "m5": m5.upper(),
            "m6": m6.upper()
        }

        self.sec_columns = {
            "sec": second.upper(),
            "s1": s1.upper(),
            "s2": s2.upper(),
            "s3": s3.upper(),
            "s4": s4.upper(),
        }

    def find_similer(self, file_dir, state):
        # Columns from Main File
        main_columns = self.main_columns
        main = main_columns.get("main")
        m1 = main_columns.get("m1")
        m2 = main_columns.get("m2")
        m3 = main_columns.get("m3")
        m4 = main_columns.get("m4")
        m5 = main_columns.get("m5")
        m6 = main_columns.get("m6")
        
        # Columns from second File
        sec_columns = self.sec_columns
        sec = sec_columns.get("sec")
        s1 = sec_columns.get("s1")
        s2 = sec_columns.get("s2")
        s3 = sec_columns.get("s3")
        s4 = sec_columns.get("s4")
        
        # Load files
        main_file = self.main_file
        sec_file = self.sec_file 
        # parent file
        pwb = load_workbook(main_file)
        pws = pwb.active
        pw = pws[main]
        # Childe File
        cwb = load_workbook(sec_file)
        cws = cwb.active
        cw = cws[sec]

        main_dic = {}
        for counter, val in enumerate(pw):
            main_dic[val.value] = counter + 1


        sec_code_list = []
        for sec_code in cw:
            sec_code_list.append(sec_code.value)

        # Load new excell
        workbook = Workbook()
        worksheet = workbook.active

        row = 1
        sf_row = 1

        if state == 1:
            for code in sec_code_list:
                val = main_dic.get(code)
                if val:
                    worksheet[f"A{row}"] = code
                    worksheet[f"B{row}"] = pws[f"{m1}{val}"].value
                    worksheet[f"C{row}"] = pws[f"{m2}{val}"].value
                    worksheet[f"D{row}"] = pws[f"{m3}{val}"].value
                    worksheet[f"E{row}"] = pws[f"{m4}{val}"].value
                    worksheet[f"F{row}"] = pws[f"{m5}{val}"].value
                    worksheet[f"G{row}"] = pws[f"{m6}{val}"].value

                    row += 1
            # save file
            workbook.save(filename=f"{file_dir}.xlsx")
        elif state == 2:
            for code in sec_code_list:
                val = main_dic.get(code)
                if val:
                    worksheet[f"A{row}"] = code
                    worksheet[f"B{row}"] = cws[f"{s1}{sf_row}"].value
                    worksheet[f"C{row}"] = cws[f"{s2}{sf_row}"].value
                    worksheet[f"D{row}"] = cws[f"{s3}{sf_row}"].value
                    worksheet[f"E{row}"] = cws[f"{s4}{sf_row}"].value

                    row += 1
                sf_row +=1
            # save file
            workbook.save(filename=f"{file_dir}.xlsx")
        elif state == 3:
            for code in sec_code_list:
                val = main_dic.get(code)
                if val:
                    worksheet[f"A{row}"] = code
                    worksheet[f"B{row}"] = pws[f"{m1}{val}"].value
                    worksheet[f"C{row}"] = pws[f"{m2}{val}"].value
                    worksheet[f"D{row}"] = pws[f"{m3}{val}"].value
                    worksheet[f"E{row}"] = pws[f"{m4}{val}"].value
                    worksheet[f"F{row}"] = pws[f"{m5}{val}"].value
                    worksheet[f"G{row}"] = pws[f"{m6}{val}"].value
                    worksheet[f"H{row}"] = cws[f"{s1}{sf_row}"].value
                    worksheet[f"I{row}"] = cws[f"{s2}{sf_row}"].value
                    worksheet[f"J{row}"] = cws[f"{s3}{sf_row}"].value
                    worksheet[f"K{row}"] = cws[f"{s4}{sf_row}"].value

                    row += 1
                sf_row += 1
            # save file
            workbook.save(filename=f"{file_dir}.xlsx")
        else:
            messagebox.showerror("خطا", "بین گزینه ها باید یکی را انتخاب کنید")
            return 

    def find_unsimiler(self, file_dir):

         # Columns from Main File
        main_columns = self.main_columns
        main = main_columns.get("main")
        
        # Columns from second File
        sec_columns = self.sec_columns
        sec = sec_columns.get("sec")
        s1 = sec_columns.get("s1")
        s2 = sec_columns.get("s2")
        s3 = sec_columns.get("s3")
        s4 = sec_columns.get("s4")
    
        # Load files
        main_file = self.main_file
        sec_file = self.sec_file 
        # parent file
        pwb = load_workbook(main_file)
        pws = pwb.active
        pw = pws[main]
        # Childe File
        cwb = load_workbook(sec_file)
        cws = cwb.active
        cw = cws[sec]

        main_dic = {}
        for index, val in enumerate(pw):
            main_dic[val.value] = index + 1

        sec_code_list = []
        for code in cw:
            sec_code_list.append(code.value)

        workbook = Workbook()
        worksheet = workbook.active


        row = 1
        sf_row = 1

        for code in sec_code_list:
            val = main_dic.get(code)
            if not val:
                worksheet[f"A{row}"] = code
                worksheet[f"B{row}"] = cws[f"{s1}{sf_row}"].value
                worksheet[f"C{row}"] = cws[f"{s2}{sf_row}"].value
                worksheet[f"D{row}"] = cws[f"{s3}{sf_row}"].value
                worksheet[f"E{row}"] = cws[f"{s4}{sf_row}"].value
                row += 1
            sf_row += 1
            

        # save file
        workbook.save(filename=f"{file_dir}.xlsx")
        


class GUI:
    def __init__(self):
        self.int_var = IntVar()
        self.state = IntVar()

        self.main_file_adrs:str
        self.second_file_adrs:str

    def set_column(self, master):
        ttk.Label(master, text="ستون های فایل اصلی و فایل ثانویه برای مقایسه مشخص کنید").place(x=10, y=10)

        ttk.Label(master, text="بارکد فایل اصلی").place(x=10, y=50)
        self.main = ttk.Entry(master, width=3)
        self.main.place(x=150, y=50)

        ttk.Label(master, text="بارکد فایل اصلی").place(x=10, y=100)
        self.sec = ttk.Entry(master, width=3)
        self.sec.place(x=150, y=100)

        ttk.Label(master, text = "دیگر ستون های فایل اصلی را مشخص کنید (در صورت موجود)").place(x=10, y=150)

        ttk.Label(master, text = "ستون ۱").place(x=10, y=200)
        self.m1 = ttk.Entry(master, width=3)
        self.m1.place(x=10, y=250)
        self.m1.insert(0, "U")

        ttk.Label(master, text = "ستون ۲").place(x=100, y=200)
        self.m2 = ttk.Entry(master, width=3)
        self.m2.place(x=100, y=250)
        self.m2.insert(0, "U")

        ttk.Label(master, text = "ستون ۳").place(x=200, y=200)
        self.m3 = ttk.Entry(master, width=3)
        self.m3.place(x=200, y=250)
        self.m3.insert(0, "U")

        ttk.Label(master, text = "ستون ۴").place(x=300, y=200)
        self.m4 = ttk.Entry(master, width=3)
        self.m4.place(x=300, y=250)
        self.m4.insert(0, "U")

        ttk.Label(master, text = "ستون ۵").place(x=400, y=200)
        self.m5 = ttk.Entry(master, width=3)
        self.m5.place(x=400, y=250)
        self.m5.insert(0, "U")

        ttk.Label(master, text = "ستون ۶").place(x=500, y=200)
        self.m6 = ttk.Entry(master, width=3)
        self.m6.place(x=500, y=250)
        self.m6.insert(0, "U")


        ttk.Label(master, text = "دیگر ستون های فایل ثانویه را مشخص کنید (در صورت موجود)").place(x=10, y=300)

        ttk.Label(master, text = "ستون ۱").place(x=10, y=350)
        self.s1 = ttk.Entry(master, width=3)
        self.s1.place(x=10, y=400)
        self.s1.insert(0, "U")

        ttk.Label(master, text = "ستون ۲").place(x=100, y=350)
        self.s2 = ttk.Entry(master, width=3)
        self.s2.place(x=100, y=400)
        self.s2.insert(0, "U")

        ttk.Label(master, text = "ستون ۳").place(x=200, y=350)
        self.s3 = ttk.Entry(master, width=3)
        self.s3.place(x=200, y=400)
        self.s3.insert(0, "U")

        ttk.Label(master, text = "ستون ۴").place(x=300, y=350)
        self.s4 = ttk.Entry(master, width=3)
        self.s4.place(x=300, y=400)
        self.s4.insert(0, "U")

    def load_files(self, master):
        ttk.Label(master, text="انتخاب فایل اصلی").place(x=10, y=10)

        mfal = ttk.Label(master)
        mfal.place(x=200, y=50)
        def open_main_file():
            main_file = askopenfilename(title="انتخاب ...", filetypes=(("Excel", "*.xlsx"), ("All", "*.*")))
            mfal.config(text=main_file)
            self.main_file_adrs = main_file
        
        ttk.Button(master, text="انتخاب", command=open_main_file).place(x=10, y=50)

        ttk.Label(master, text="انتخاب فایل ثانویه").place(x=10, y=100)
        sfal = ttk.Label(master)
        sfal.place(x=200, y=150)

        def choose_sec_file():
            sec_file = askopenfilename(title="انتخاب ...", filetypes=(("Excel", "*.xlsx"), ("All", "*.*")))
            sfal.config(text=sec_file)
            self.second_file_adrs = sec_file


        ttk.Button(master, text="انتخاب", command=choose_sec_file).place(x=10, y=150)

    def seprate_file_gui(self, master):
        ttk.Label(master, text="نوع عملیات را مشخص کنید").place(x=10, y=10)

        ttk.Radiobutton(master, text="فقط فایل های مشابه", value=1, variable=self.int_var).place(x=10, y=50)
        ttk.Radiobutton(master, text="فقط فایل های غیر مشابه", value=2, variable=self.int_var).place(x=10, y=100)

        ttk.Label(master, text="انتخاب ستون های مورد نیاز").place(x=10, y=150)

        ttk.Radiobutton(master, text="ستون های فایل اصلی", value=1, variable=self.state).place(x=10, y=200)
        ttk.Radiobutton(master, text="ستون های فایل قانویه", value=2, variable=self.state).place(x=10, y=230)
        ttk.Radiobutton(master, text="ستون های هردو فایل را می خواهم", value=3, variable=self.state).place(x=10, y=260)


        ttk.Button(master, text="انجام عملیات" ,command=self.do_func).place(x=10,y=300)

    def do_func(self):
        try:
            rqr = RepeatedQr(self.main_file_adrs, self.second_file_adrs)
            rqr.load_columns(
                self.main.get(),
                self.sec.get(),
                self.m1.get(),
                self.m2.get(),
                self.m3.get(),
                self.m4.get(),
                self.m5.get(),
                self.m6.get(),
                self.s1.get(),
                self.s2.get(),
                self.s3.get(),
                self.s4.get()
            )

            if self.int_var.get() == 1:
                save_file_adrs = asksaveasfilename(title="ذخیره در ...", filetypes=(("Excell", "*.xlsx"), ("All", "*.*")))
                rqr.find_similer(save_file_adrs, self.state.get())
                messagebox.showinfo("عملیات انجام شد", "فایل \n {} \n ذخیره شد".format(save_file_adrs))

            elif self.int_var.get() == 2:
                save_file_adrs = asksaveasfilename(title="ذخیره در ...", filetypes=(("Excell", "*.xlsx"), ("All", "*.*")))
                rqr.find_unsimiler(save_file_adrs)
                messagebox.showinfo("عملیات انجام شد", "فایل \n {} \n ذخیره شد".format(save_file_adrs))

            else:
                messagebox.showerror("خطا", "بین گزینه مشابه و غیر مشابه یکی را انتخاب کنید")

        except AttributeError:
            messagebox.showerror("خطا", "فایل ها را انتخاب کنید")
    
