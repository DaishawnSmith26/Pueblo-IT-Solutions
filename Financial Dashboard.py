#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Feb  9 16:36:45 2022

@author: terrydennison
"""


from tkinter import Tk, Button, Text, Label, CENTER, INSERT
from PIL import ImageTk, Image
from tkinter import filedialog
import os
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pyautogui
import win32print
import win32api
import os, sys



# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Main Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates main window of program
main_window = Tk()
main_window.attributes('-fullscreen', True)
main_window.title("Pueblo Cooperative Care")


# Open file explorer method
def browseFiles():
    global filepath

    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select Excel File",
                                          filetypes=(("Excel",
                                                      "*.xlsx"),
                                                     ("all files",
                                                      "*.*")))

    filepath = os.path.abspath(filename)

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ AP Summary Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates AP Summmary window


def create_AP_Summary_frame():
    # Creates GUI object and sets size of GUI
    ap_summary_window = Tk()
    ap_summary_window.attributes('-fullscreen', True)
    ap_summary_window.title("AP Summary")

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'AP Summary File Upload' GUI
    ap_summary_window.title('AP Summary File Upload')

    ap_label = Label(ap_summary_window, text="AP Summary")
    ap_label.config(font=("Courier", 34))
    ap_label.place(x=100, y=100)

    ap_descr = Label(ap_summary_window,
                     text="Please select 'AP Summary' file for analyzing")
    ap_descr.config(font=("Courier", 16))
    ap_descr.place(x=100, y=200)

    # Creates textbox for 'AP Summary Frame' GUI and sets location on GUI
    txt_file = Text(ap_summary_window, height=4,
                    width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(ap_summary_window, text="Back to Main", command=lambda: [ap_summary_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'AP Summary' file for upload and location on 'AP Summary Frame' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(ap_summary_window, text="Find File", command=lambda: [txt_file.configure(state='normal'), txt_file.delete('1.0', "end"),
                                                                                   browseFiles(), txt_file.insert(INSERT, filepath),
                                                                                   txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    # Nested fuction for ETL of 'AP Summary' .xslx file
    def anaylze_AP_summary():
        df = pd.read_excel(
            filepath, sheet_name=1)

        df['Unnamed: 0'] = np.nan
        df.dropna(how='all', axis=1, inplace=True)

        df = pd.melt(df, id_vars='Unnamed: 1',
                     var_name='Date Ranges', value_name='Amount')

        # Nested function for creating figure for 'grouped bar chart'

        def create_plot():
            axs = sns.catplot(x='Date Ranges', y='Amount',
                              hue='Unnamed: 1', data=df, kind='bar')
            axs.fig.suptitle('AP Summary')
            axs.fig.set_size_inches(10, 10)

            return axs.fig

        # Creates window for graphical display of data and size of GUI frame
        ap_analyze_window = Tk()
        ap_analyze_window.attributes('-fullscreen', True)

        ap_analyze_window.title("AP Summary")
        # Calls create_plot() function to put 'grouped bar chart' in 'Canvas' frame
        figure = create_plot()
        canvas = FigureCanvasTkAgg(figure, master=ap_analyze_window)
        canvas.draw()
        # Sets position of created 'Canvas'

        canvas.get_tk_widget().place(relx=0.01, rely=.1, )

        fig = plt.figure(figsize=(10, 10))
        fig.set_size_inches(10, 10)
        plt.axis('equal')
        pie_labels = (df['Unnamed: 1'][25], df['Unnamed: 1'][26],
                      df['Unnamed: 1'][27], df['Unnamed: 1'][28])
        pie_data = (df['Amount'][25], df['Amount'][26],
                    df['Amount'][27], df['Amount'][28])
        colors = sns.color_palette('bright')[0:4]

        plt.title("AP Summary Totals")
        plt.pie(pie_data, labels=pie_labels, colors=colors,
                autopct='%.0f%%', pctdistance=.5, explode=[0.1]*4, textprops={"fontsize": 14}, shadow=True)
        plt.legend()

        canvas = FigureCanvasTkAgg(fig, master=ap_analyze_window)
        canvas.draw()
        canvas.get_tk_widget().place(relx=.4, rely=.1,)

        # Hides 'AP Summary GUI'
        ap_summary_window.withdraw()

        # WIP

        def pdf_Print():
                    printer_name = win32print.GetDeafaultPrint
                

                    #Grab Filename
                    filename =filedialog.askopenfilename(initialdir="/",
                                                title="Select Excel File",
                                                filetypes=(("Excel",
                                                            "*.xlsx"),
                                                            ("all files",
                                                            "*.*")))
                    if filename:
                        win32api.ShellExecute(0, "print",filename,None,",",0)
        # WIP Button for Printing 'Canvas' and button position
        btn_print = Button(ap_analyze_window, text="Print as PDF", command=lambda: [pdf_Print()],
                           width=20, height=3, fg='green')

        btn_print.place(x=1700, y=100)

        # WIP Button for saving file as 'PDF' and button position

        btn_main = Button(ap_analyze_window, text="Back to Main", command=lambda: [ap_analyze_window.withdraw(), main_window.deiconify(), plt.close('all')],
                          width=20, height=3, fg='green')
        btn_main.place(x=1700, y=300)

        btn2_exit = Button(ap_analyze_window, text="Exit",
                           command=lambda: [os._exit(0)], width=20, height=3, fg='green')
        btn2_exit.place(x=1700, y=400)

        

        # method to capture GUI image, saves as png, and converts to pdf
        def save_data():

            screenshottaker = pyautogui.screenshot()
            save_path = filedialog.asksaveasfilename()
            screenshottaker.save(save_path+".png")

            saved_image = Image.open(save_path+".png")
            im_1 = saved_image.convert('RGB')
            im_1.save(save_path+".pdf")
            os.remove(save_path+".png")

        # Saves 'Analyze' GUI frame
        btn_pdf = Button(ap_analyze_window, text="Save as PDF", command=lambda: [save_data()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=1700, y=200)

        ap_analyze_window.mainloop()

    # Button event handler for calling 'analyze_ap_summary' function and button position
    btn_analyze = Button(ap_summary_window, text="Analyze File", command=anaylze_AP_summary,
                         width=20, height=3, fg='green')
    btn_analyze.place(x=700, y=550)

    # Displays 'AP Summary' GUI
    ap_summary_window.mainloop()


# Creates Button on main window for 'AP Summary' Option
btn_AP_Summary = Button(main_window, text="AP Summary", command=create_AP_Summary_frame,
                        width=50, height=3, fg='green')
btn_AP_Summary.place(x=50, y=100)

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ AR Summary Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates AR Summmary window


def create_AR_Summary_frame():

    # Creates GUI object and sets size of GUI
    ar_summary_window = Tk()
    ar_summary_window.attributes('-fullscreen', True)

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'AR Summary File Upload' GUI
    ar_summary_window.title('AR Summary File Upload')

    # Creates textbox for 'AR Summary Frame' GUI and sets location on GUI
    ar_label = Label(ar_summary_window, text="AR Summary")
    ar_label.config(font=("Courier", 34))
    ar_label.place(x=100, y=100)

    ar_descr = Label(ar_summary_window,
                     text="Please select 'AR Summary' file for analyzing")
    ar_descr.config(font=("Courier", 16))
    ar_descr.place(x=100, y=200)

    ar_txt_file = Text(ar_summary_window, height=4,
                       width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    ar_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(ar_summary_window, text="Back to Main", command=lambda: [ar_summary_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'AR Summary' file for upload and location on 'AR Summary Frame' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(ar_summary_window, text="Find File", command=lambda: [ar_txt_file.configure(state='normal'), ar_txt_file.delete('1.0', "end"),
                                                                                   browseFiles(), ar_txt_file.insert(INSERT, filepath),
                                                                                   ar_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    # Nested fuction for ETL of 'AR Summary' .xslx file
    def anaylze_AR_summary():
        df = pd.read_excel(
            filepath, sheet_name=1)

        df['Unnamed: 0'] = np.nan
        df.dropna(how='all', axis=1, inplace=True)

        df = pd.melt(df, id_vars='Unnamed: 1',
                     var_name='Date Ranges', value_name='Amount')

        # Nested function for creating figure for 'grouped bar chart'
        def create_plot():
            axs = sns.catplot(x='Date Ranges', y='Amount',
                              hue='Unnamed: 1', data=df, kind='bar')
            axs.fig.suptitle('AR Summary')
            axs.fig.set_size_inches(10, 10)
            return axs.fig
        # Creates window for grARhical display of data and size of GUI frame
        ar_analyze_window = Tk()
        ar_analyze_window.attributes('-fullscreen', True)

        ar_analyze_window.title("AR Summary")
        # Calls create_plot() function to put 'grouped bar chart' in 'Canvas' frame
        figure = create_plot()
        canvas = FigureCanvasTkAgg(figure, master=ar_analyze_window)
        canvas.draw()

        # Sets position of created 'Canvas'

        canvas.get_tk_widget().place(relx=0.01, rely=.1,)
        fig = plt.figure(figsize=(10, 10))
        fig.set_size_inches(11.5, 10)
        plt.axis('equal')

        pie_labels = (df['Unnamed: 1'][20:23])
        pie_data = (df['Amount'][20:23])

        colors = sns.color_palette('bright')[0:4]

        plt.title("AR Summary Totals")

        # Cretes pie graph
        plt.pie(pie_data, labels=pie_labels, colors=colors,
                autopct='%.0f%%', pctdistance=.4, explode=[0.04]*3, textprops={"fontsize": 10.2}, shadow=True)
        plt.legend()

        canvas = FigureCanvasTkAgg(fig, master=ar_analyze_window)
        canvas.draw()
        canvas.get_tk_widget().place(relx=.4, rely=.1)

        # Hides 'AR Summary GUI'
        ar_summary_window.withdraw()

        # WIP
        def pdf_Print():
                    printer_name = win32print.GetDeafaultPrint
                

                    #Grab Filename
                    filename =filedialog.askopenfilename(initialdir="/",
                                                title="Select Excel File",
                                                filetypes=(("Excel",
                                                            "*.xlsx"),
                                                            ("all files",
                                                            "*.*")))
                    if filename:
                        win32api.ShellExecute(0, "print",filename,None,",",0)
          # fp = tempfile.TemporaryFile()
          # fp.write(ar_analyze_window)

          # ps = canvas..postscript(colormode='color')
          # img = Image.open(io.BytesIO(ps.encode('utf-8')))
          # img.save('filename.jpg', 'jpeg')

          # postscript_file = "tmp_snARshot.ps"
          # subprocess.call(["impg", "-window", "td", postscript_file])

        # WIP Button for Printing 'Canvas' and button position
        btn_print = Button(ar_analyze_window, text="Print as PDF", command=lambda: [pdf_Print()],
                           width=20, height=3, fg='green')

        btn_print.place(x=1700, y=100)

        # WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(ar_analyze_window, text="Save as PDF", command=lambda: [save_data()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=1700, y=200)

        btn_main = Button(ar_analyze_window, text="Back to Main", command=lambda: [ar_analyze_window.withdraw(), main_window.deiconify(), plt.close('all')],
                          width=20, height=3, fg='green')
        btn_main.place(x=1700, y=300)

        btn2_exit = Button(ar_analyze_window, text="Exit",
                           command=lambda: [os._exit(0)], width=20, height=3, fg='green')
        btn2_exit.place(x=1700, y=400)

        # method to capture GUI image, saves as png, and converts to pdf
        def save_data():

            screenshottaker = pyautogui.screenshot()
            save_path = filedialog.asksaveasfilename()
            screenshottaker.save(save_path+".png")

            saved_image = Image.open(save_path+".png")
            im_1 = saved_image.convert('RGB')
            im_1.save(save_path+".pdf")
            os.remove(save_path+".png")

        # Display's ' Analyze' GUI frame
        ar_analyze_window.mainloop()

    # Button event handler for calling 'analyze_AR_summary' function and button position
    btn_analyze = Button(ar_summary_window, text="Analyze File", command=anaylze_AR_summary,
                         width=20, height=3, fg='green')
    btn_analyze.place(x=700, y=550)

    # Displays 'AR Summary' GUI
    ar_summary_window.mainloop()


# Creates Button on main window for 'AR Summary' Option
btn_AR_Summary = Button(main_window, text="AR Summary", command=create_AR_Summary_frame,
                        width=50, height=3, fg='green')
btn_AR_Summary.place(x=50, y=200)

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Assets Liabilites Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates 'Assets Liabilities' window


def create_assets_liabilities_Summary_frame():
    # Creates GUI object and sets size of GUI
    al_summary_window = Tk()
    al_summary_window.attributes('-fullscreen', True)

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of Assets Liabilities' File Upload' GUI
    al_summary_window.title('Assets and Liabilites Summary File Upload')

    al_label = Label(al_summary_window, text="Assets Liabilities Summary")
    al_label.config(font=("Courier", 34))
    al_label.place(x=100, y=100)

    al_descr = Label(
        al_summary_window, text="Please select 'Stmnt of Assets and Liabilities' file for analyzing")
    al_descr.config(font=("Courier", 16))
    al_descr.place(x=100, y=200)

    # Creates textbox for Assets Liabilities' GUI and sets location on GUI
    al_txt_file = Text(al_summary_window, height=4,
                       width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    al_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(al_summary_window, text="Back to Main", command=lambda: [al_summary_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    btn_home.place(x=100, y=550)

    # Button for selecting Assets Liabilities' file for upload and location on Assets Liabilities' GUI
    btn_file_upload = Button(al_summary_window, text="Find File", command=lambda: [al_txt_file.configure(state='normal'), al_txt_file.delete('1.0', "end"),
                                                                                   browseFiles(), al_txt_file.insert(INSERT, filepath),
                                                                                   al_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    def analyze_AL_summary():
        df = pd.read_excel(
            filepath, sheet_name=1)

        # captured data points for total assets
        prev_year_total_assets = (df['Unnamed: 8'][47])
        current_year_total_assets = (df['Unnamed: 6'][47])

        # captured data points for total liabilities
        prev_year_total_liabilities = (df['Unnamed: 8'][68])
        current_year_total_liabilities = (df['Unnamed: 6'][68])

        # captured data points for total equity
        prev_year_total_equity = (df['Unnamed: 8'][78])
        current_year_total_equity = (df['Unnamed: 6'][78])

        # Creates GUI for 'Assets Liabilites' data
        al_analyze_window = Tk()
        al_analyze_window.attributes('-fullscreen', True)
        al_analyze_window.title("Assets Liabilites Summary")

        # increases size of entire seaborn contents
        sns.set_theme(font_scale=1.5)

        # Create total assets bar graph
        def create_total_assets_bar():

            sns.set_palette(['orange', 'purple'])
            ta_axs = sns.catplot(x=['Jan 1, 22', 'Jan 1, 21'], y=[current_year_total_assets, prev_year_total_assets],
                                 data=df, kind='bar')
            ta_axs.fig.suptitle('Total Assets Summary')
            plt.xlabel("Year")
            plt.ylabel("$ Amount (In Millions) ")

            ax = ta_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            ta_axs.fig.set_size_inches(10, 7.5)
            return ta_axs.fig

        ta_fig = create_total_assets_bar()
        canvas = FigureCanvasTkAgg(ta_fig, master=al_analyze_window)
        canvas.draw()

        canvas.get_tk_widget().place(relx=.01, rely=.02)

        # creates total liabilities bar graph
        def create_total_liabilities_bar():
            sns.set_palette(['red', 'green'])
            tl_axs = sns.catplot(x=['Jan 1, 22', 'Jan 1, 21'], y=[current_year_total_liabilities, prev_year_total_liabilities],
                                 data=df, kind='bar')
            tl_axs.fig.suptitle('Total Liabilities Summary')
            plt.xlabel("Year")
            plt.ylabel("$ Amount")

            ax = tl_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            tl_axs.fig.set_size_inches(10, 7.5)
            return tl_axs.fig

        tl_fig = create_total_liabilities_bar()
        canvas = FigureCanvasTkAgg(tl_fig, master=al_analyze_window)

        canvas.draw()

        canvas.get_tk_widget().place(relx=0.4, rely=.02,)

        # create total equity bar graph
        def create_total_equity_bar():
            sns.set_palette(['teal', 'grey'])
            te_axs = sns.catplot(x=['Jan 1, 22', 'Jan 1, 21'], y=[current_year_total_equity, prev_year_total_equity],
                                 data=df, kind='bar')
            te_axs.fig.suptitle('Total Equity Summary')
            plt.xlabel("Year")
            plt.ylabel("$ Amount (In Millions) ")

            ax = te_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            te_axs.fig.set_size_inches(10, 6.8)
            return te_axs.fig

        te_fig = create_total_equity_bar()
        canvas = FigureCanvasTkAgg(te_fig, master=al_analyze_window)
        canvas.draw()
        canvas.get_tk_widget().place(relx=.01, rely=.53)

        
        # method to capture GUI image, saves as png, and converts to pdf
        def save_data():

            screenshottaker = pyautogui.screenshot()
            save_path = filedialog.asksaveasfilename()
            screenshottaker.save(save_path+".png")

            saved_image = Image.open(save_path+".png")
            im_1 = saved_image.convert('RGB')
            im_1.save(save_path+".pdf")
            os.remove(save_path+".png")

        # WIP
        def pdf_Print():
                    printer_name = win32print.GetDeafaultPrint
                

                    #Grab Filename
                    filename =filedialog.askopenfilename(initialdir="/",
                                                title="Select Excel File",
                                                filetypes=(("Excel",
                                                            "*.xlsx"),
                                                            ("all files",
                                                            "*.*")))
                    if filename:
                        win32api.ShellExecute(0, "print",filename,None,",",0)
        # WIP Button for Printing 'Canvas' and button position``
        btn_print = Button(al_analyze_window, text="Print as PDF", command=lambda: [pdf_Print()],
                           width=20, height=3, fg='green')

        btn_print.place(x=1700, y=100)

        # WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(al_analyze_window, text="Save as PDF", command=lambda: [save_data()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=1700, y=200)

        btn_main = Button(al_analyze_window, text="Back to Main", command=lambda: [al_analyze_window.withdraw(), main_window.deiconify(), plt.close('all')],
                          width=20, height=3, fg='green')
        btn_main.place(x=1700, y=300)

        btn_exit = Button(al_analyze_window, text="Exit",
                          command=lambda: [os._exit(0)], width=20, height=3, fg='green')
        btn_exit.place(x=1700, y=400)

        al_summary_window.withdraw()
        al_analyze_window.mainloop()

    btn_analyze = Button(al_summary_window, text="Analyze File", command=analyze_AL_summary,
                         width=20, height=3, fg='green')
    btn_analyze.place(x=700, y=550)

    al_summary_window.mainloop()


# Creates Button on main window for 'Statement of Assets-Liabilities' Option
btn_Assets_Liabilities = Button(main_window, text="Statement of Assets-Liabilities",
                                command=create_assets_liabilities_Summary_frame, width=50, height=3, fg='green')
btn_Assets_Liabilities.place(x=50, y=300)


# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Revenue Expenses Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

# Creates 'Revenue and Expenses' window
def create_Revenue_Expenses_frame():

    revenues_expsenses_window = Tk()
    revenues_expsenses_window.attributes('-fullscreen', True)

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'Revenue and Expenses' GUI
    revenues_expsenses_window.title('Revenue and Expneses Summary File Upload')

    re_label = Label(revenues_expsenses_window,
                     text="Revenue and Expenses Summary")
    re_label.config(font=("Courier", 34))
    re_label.place(x=100, y=100)

    re_descr = Label(revenues_expsenses_window,
                     text="Please select 'Stmnt of Revenues and Expenses' file for analyzing")
    re_descr.config(font=("Courier", 16))
    re_descr.place(x=100, y=200)

    # Creates textbox for 'Revenue and Expenses' GUI and sets location on GUI
    re_txt_file = Text(revenues_expsenses_window, height=4,
                       width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    re_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(revenues_expsenses_window, text="Back to Main", command=lambda: [revenues_expsenses_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'Revenue and Expenses' file for upload and location on ''Revenue and Expenses' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(revenues_expsenses_window, text="Find File", command=lambda: [re_txt_file.configure(state='normal'), re_txt_file.delete('1.0', "end"),
                                                                                           browseFiles(), re_txt_file.insert(INSERT, filepath),
                                                                                           re_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    def analyze_revenue_expenses():
        df = pd.read_excel(
            filepath, sheet_name=1)

        revenues_expsenses_window.withdraw()

        # Income from Mobile Shower Program from Excel sheet
        current_month = df['Unnamed: 6'][0]
        prev_year = df['Unnamed: 8'][0]

        currrent_month_income_mshower = (df['Unnamed: 6'][37])
        prev_year_income_mshower = (df['Unnamed: 8'][37])

        # Capture 'Expenses' datapoints for Excel
        currrent_month_expenses_food_sack = (df['Unnamed: 6'][88])
        prev_year_expenses_food_sack = (df['Unnamed: 8'][88])

        currrent_month_expenses_now = (df['Unnamed: 6'][103])
        prev_year_expenses_now = (df['Unnamed: 8'][103])

        currrent_month_expenses_mshower = (df['Unnamed: 6'][115])
        prev_year_expenses_mshower = (df['Unnamed: 8'][115])

        fig, ax = plt.subplots(1, 1)

        data = [['Food Sack', currrent_month_expenses_food_sack, prev_year_expenses_food_sack],
                ['Nutrition on Wheels', currrent_month_expenses_now,
                    prev_year_expenses_now],
                ['Mobile Shower', currrent_month_expenses_mshower,
                    prev_year_expenses_mshower]
                ]

        df2 = pd.DataFrame(data=data)

        # Creates 'Revenue and Expness Current Year' bar grpah
        def create_plot():
            axs = sns.catplot(x=0, y=1,
                              data=df2, kind='bar')
            axs.fig.suptitle('Current Program Expenses')
            axs.fig.set_size_inches(10, 10)
            plt.ylabel('Amount ($)')
            plt.xlabel('Programs')
            return axs.fig

        # Creates 'Revenue and Expness Previous Year' bar grpah
        def create_prev_year_plot():

            axs2 = sns.catplot(x=0, y=2,
                               data=df2, kind='bar')
            axs2.fig.suptitle('Previous Year Program Expenses')
            axs2.fig.set_size_inches(10, 10)
            plt.ylabel('Amount ($)')
            plt.xlabel('Programs')
            return axs2.fig

        current_fig = create_plot()
        prev_fig = create_prev_year_plot()

        rev_exp_analyze_window = Tk()
        rev_exp_analyze_window.attributes('-fullscreen', True)
        rev_exp_analyze_window.title("Revenue and Expneses Summary")

        canvas = FigureCanvasTkAgg(current_fig, master=rev_exp_analyze_window)
        canvas2 = FigureCanvasTkAgg(prev_fig, master=rev_exp_analyze_window)

        canvas.draw()
        canvas2.draw()

        canvas.get_tk_widget().place(relx=.01, rely=.02, )

        canvas2.get_tk_widget().place(relx=.4, rely=.02, )

        btn_main = Button(rev_exp_analyze_window, text="Back to Main", command=lambda: [rev_exp_analyze_window.withdraw(), main_window.deiconify(), plt.close('all')],
                          width=20, height=3, fg='green')
        btn_main.place(x=1700, y=300)

        # method to capture GUI image, saves as png, and converts to pdf
        def save_data():

            screenshottaker = pyautogui.screenshot()
            save_path = filedialog.asksaveasfilename()
            screenshottaker.save(save_path+".png")

            saved_image = Image.open(save_path+".png")
            im_1 = saved_image.convert('RGB')
            im_1.save(save_path+".pdf")
            os.remove(save_path+".png")

        
        
        
        # WIP
        def pdf_Print():
                    printer_name = win32print.GetDeafaultPrint
                

                    #Grab Filename
                    filename =filedialog.askopenfilename(initialdir="/",
                                                title="Select Excel File",
                                                filetypes=(("Excel",
                                                            "*.xlsx"),
                                                            ("all files",
                                                            "*.*")))
                    if filename:
                        win32api.ShellExecute(0, "print",filename,None,",",0)
        # WIP Button for Printing 'Canvas' and button position
        btn_print = Button(rev_exp_analyze_window, text="Print as PDF",  command=lambda: [pdf_Print()],
                           width=20, height=3, fg='green')

        btn_print.place(x=1700, y=100)

        # WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(rev_exp_analyze_window, text="Save as PDF", command=lambda: [save_data()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=1700, y=200)

        btn_exit = Button(rev_exp_analyze_window, text="Exit",
                          command=lambda: [os._exit(0)], width=20, height=3, fg='green')
        btn_exit.place(x=1700, y=400)

        rev_exp_analyze_window.mainloop()

    btn_rev_exp_analyze = Button(revenues_expsenses_window, text="Analyze File", command=analyze_revenue_expenses,
                                 width=20, height=3, fg='green')
    btn_rev_exp_analyze.place(x=700, y=550)

    revenues_expsenses_window.mainloop()


# Creates Button on main window for 'Statement Revenue and Expenses' Option
btn_Revenue_Expenses = Button(main_window, text="Statement of Revenue and Expenses",
                              command=create_Revenue_Expenses_frame, width=50, height=3, fg='green')
btn_Revenue_Expenses.place(x=50, y=400)


# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Revenue and Expenses Comparison to Year Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates 'Revenue and Expenses Comparison to Prev Year' window

def create_Revenue_Expenses_Comparison_frame():

    revenues_expsenses_comparison_window = Tk()
    revenues_expsenses_comparison_window .attributes('-fullscreen', True)

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'Revenue and Expenses Comparison' GUI
    revenues_expsenses_comparison_window.title(
        'Revenue Expneses Comparison Summary File Upload')

    recy_label = Label(revenues_expsenses_comparison_window,
                       text="Revenue and Expenses Compared to Previous Year Summary")
    recy_label.config(font=("Courier", 32))
    recy_label.place(x=100, y=100)

    recy_descr = Label(revenues_expsenses_comparison_window,
                       text="Please select 'Stmnt of Revenues and Expenses Compared to Prev Year' file for analyzing")
    recy_descr.config(font=("Courier", 16))
    recy_descr.place(x=100, y=200)

    # Creates textbox for 'Revenue and Expenses Comparison' GUI and sets location on GUI
    re_C_txt_file = Text(revenues_expsenses_comparison_window, height=4,
                         width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    re_C_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(revenues_expsenses_comparison_window, text="Back to Main", command=lambda: [revenues_expsenses_comparison_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'Revenue and Expenses Comparison' file for upload and location on 'Revenue and Expenses Comparison' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(revenues_expsenses_comparison_window, text="Find File", command=lambda: [re_C_txt_file.configure(state='normal'), re_C_txt_file.delete('1.0', "end"),
                                                                                                      browseFiles(), re_C_txt_file.insert(INSERT, filepath),
                                                                                                      re_C_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    # Function for analyzing "Revenue and Expenses" Excel FIle
    def analyze_revenue_expenses_comparision():
        df = pd.read_excel(
            filepath, sheet_name=1)

        revenues_expsenses_comparison_window.withdraw()

        # Food Sack
        prev_year_expenses_food_sack = (df['Unnamed: 9'][103])
        current_year_expenses_food_sack = (df['Unnamed: 7'][103])
        delta_food_sack = (df['Unnamed: 11'][103])

        # NOW Program
        prev_year_expenses_now = (df['Unnamed: 9'][118])
        current_year_expenses_now = (df['Unnamed: 7'][118])
        delta_expenses_now = (df['Unnamed: 11'][118])

        # Mobile Shower
        prev_year_expenses_mshower = (df['Unnamed: 9'][130])
        current_expenses_mshower = (df['Unnamed: 7'][130])
        delta_expenses_mshower = (df['Unnamed: 11'][130])

        repy_analyze_window = Tk()
        repy_analyze_window.attributes('-fullscreen', True)
        repy_analyze_window.title(
            "Revenue and Expenses Comparison by Year Summary")

        sns.set_theme(font_scale=1.5)

        # Creates 'Total Assets' Bar Graph
        def create_food_sack_bar():

            sns.set_palette(['orange', 'purple', 'grey'])
            fs_axs = sns.catplot(x=['Current Year', 'Previous Year', '$ Change'], y=[current_year_expenses_food_sack, prev_year_expenses_food_sack, delta_food_sack],
                                 data=df, kind='bar')
            fs_axs.fig.suptitle('Food Sack Expenses')
            plt.xlabel("Dates")
            plt.ylabel("$ Amount")
            plt.ylim(0, None)
            ax = fs_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            fs_axs.fig.set_size_inches(10, 7)
            return fs_axs.fig

        fs_fig = create_food_sack_bar()
        canvas = FigureCanvasTkAgg(fs_fig, master=repy_analyze_window)
        canvas.draw()

        canvas.get_tk_widget().place(relx=.01, rely=.01, )

        # Creates NOW bar graph
        def create_now_bar():
            sns.set_palette(['red', 'green', 'teal'])
            now_axs = sns.catplot(x=['Current Year', 'Previous Year', '$ Change'], y=[current_year_expenses_now, prev_year_expenses_now, delta_expenses_now],
                                  data=df, kind='bar')
            now_axs.fig.suptitle('NOW Expenses')
            plt.xlabel("Programs")
            plt.ylabel("$ Amount")
            plt.ylim(0, None)
            ax = now_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            now_axs.fig.set_size_inches(10, 7)
            return now_axs.fig

        now_fig = create_now_bar()
        canvas = FigureCanvasTkAgg(now_fig, master=repy_analyze_window)
        canvas.draw()

        canvas.get_tk_widget().place(relx=.4, rely=.01, )

        # Creates 'Mobile Shower' Bar graph
        def create_mobile_shower_bar():
            sns.set_palette(['teal', 'grey', 'brown'])
            ms_axs = sns.catplot(x=['Current Year', 'Previous Year', '$ Change'], y=[current_expenses_mshower, prev_year_expenses_mshower, delta_expenses_mshower],
                                 data=df, kind='bar')
            ms_axs.fig.suptitle('Net Income by Program')
            plt.xlabel("Programs")
            plt.ylabel("$ Amount")
            plt.ylim(0, None)
            ax = ms_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            ms_axs.fig.set_size_inches(10, 7)
            return ms_axs.fig

        ms_fig = create_mobile_shower_bar()
        canvas = FigureCanvasTkAgg(ms_fig, master=repy_analyze_window)
        canvas.draw()
        canvas.get_tk_widget().place(relx=.01, rely=.5,)

        btn_main = Button(repy_analyze_window, text="Back to Main", command=lambda: [repy_analyze_window.withdraw(), main_window.deiconify(), plt.close('all')],
                          width=20, height=3, fg='green')
        btn_main.place(x=1700, y=300)

        btn_exit = Button(repy_analyze_window, text="Exit",
                          command=lambda: [os._exit(0)], width=20, height=3, fg='green')
        btn_exit.place(x=1700, y=400)

        # method to capture GUI image, saves as png, and converts to pdf
        def save_data():

            screenshottaker = pyautogui.screenshot()
            save_path = filedialog.asksaveasfilename()
            screenshottaker.save(save_path+".png")

            saved_image = Image.open(save_path+".png")
            im_1 = saved_image.convert('RGB')
            im_1.save(save_path+".pdf")
            os.remove(save_path+".png")

        # WIP
        def pdf_Print():
                    printer_name = win32print.GetDeafaultPrint
                

                    #Grab Filename
                    filename =filedialog.askopenfilename(initialdir="/",
                                                title="Select Excel File",
                                                filetypes=(("Excel",
                                                            "*.xlsx"),
                                                            ("all files",
                                                            "*.*")))
                    if filename:
                        win32api.ShellExecute(0, "print",filename,None,",",0)
        # WIP Button for Printing 'Canvas' and button position
        btn_print = Button(repy_analyze_window, text="Print as PDF", command=lambda: [pdf_Print()],
                           width=20, height=3, fg='green')

        btn_print.place(x=1700, y=100)

        # WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(repy_analyze_window, text="Save as PDF", command=lambda: [save_data()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=1700, y=200)

        repy_analyze_window.mainloop()

    btn_rec_analyze = Button(revenues_expsenses_comparison_window, text="Analyze File", command=analyze_revenue_expenses_comparision,
                             width=20, height=3, fg='green')
    btn_rec_analyze.place(x=700, y=550)

    revenues_expsenses_comparison_window.mainloop()


# Creates Button on main window for 'Statement of Assets-Liabilities' Option
btn_Revenue_Expenses_Comparison = Button(main_window, text="Statement of Revenue and Expenses Comparison by Year",
                                         command=create_Revenue_Expenses_Comparison_frame, width=50, height=3, fg='green')
btn_Revenue_Expenses_Comparison.place(x=50, y=500)

# Creates 'Revenue and Expenses Compared to Budget' window

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Revenue Expenses Budget Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

# Mayble include table for negative values??????????????????????????


def create_Revenue_Expenses_Budget_frame():

    revenues_expsenses_budget_window = Tk()
    revenues_expsenses_budget_window .attributes('-fullscreen', True)

   # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'Revenue and Expenses Compared to Budget' GUI
    revenues_expsenses_budget_window.title(
        'Revenue and Expenses Compared to Budget Summary File Upload')

    recb_label = Label(revenues_expsenses_budget_window,
                       text="Revenue and Expenses Compared to Budget Summary")
    recb_label.config(font=("Courier", 34))
    recb_label.place(x=100, y=100)

    recb_descr = Label(revenues_expsenses_budget_window,
                       text="Please select 'Stmnt of Revenues and Expenses Compared to Budget' file for analyzing")
    recb_descr.config(font=("Courier", 16))
    recb_descr.place(x=100, y=200)

    # Creates textbox for 'Revenue and Expenses Compared to Budget' GUI and sets location on GUI
    re_B_txt_file = Text(revenues_expsenses_budget_window, height=4,
                         width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    re_B_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(revenues_expsenses_budget_window, text="Back to Main", command=lambda: [revenues_expsenses_budget_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'Revenue and Expenses Compared to Budget' file for upload and location on 'Revenue and Expenses Compared to Budget' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(revenues_expsenses_budget_window, text="Find File", command=lambda: [re_B_txt_file.configure(state='normal'), re_B_txt_file.delete('1.0', "end"),
                                                                                                  browseFiles(), re_B_txt_file.insert(INSERT, filepath),
                                                                                                  re_B_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    def analyze_revenue_expneses_budget():

        df = pd.read_excel(
            filepath, sheet_name=1)

        # Capture of datapoints from Excel file
        current_month_fs = (df['Unnamed: 6'][93])
        current_month_budget_fs = (df['Unnamed: 8'][93])
        current_month_obudget_fs = (df['Unnamed: 10'][93])

        current_months_budget_fs = (df['Unnamed: 12'][93])
        current_ytd_budget_fs = (df['Unnamed: 14'][93])
        current_ytd_obudget_fs = (df['Unnamed: 16'][93])
        current_ytd_abudget_fs = (df['Unnamed: 18'][93])

        current_month_now = (df['Unnamed: 6'][108])
        current_month_budget_now = (df['Unnamed: 8'][108])
        current_month_obudget_now = (df['Unnamed: 10'][108])

        current_year_budget_now = (df['Unnamed: 12'][108])
        current_ytd_budget_now = (df['Unnamed: 14'][108])
        current_ytd_obudget_now = (df['Unnamed: 16'][108])
        current_ytd_abudget_now = (df['Unnamed: 18'][108])

        current_month_ms = (df['Unnamed: 6'][121])
        current_month_budget_ms = (df['Unnamed: 8'][121])
        current_month_obudget_ms = (df['Unnamed: 10'][121])

        current_year_budget_ms = (df['Unnamed: 12'][121])
        current_ytd_budget_ms = (df['Unnamed: 14'][121])
        current_ytd_obudget_ms = (df['Unnamed: 16'][121])
        current_ytd_abudget_ms = (df['Unnamed: 18'][121])

        # Creation of Food Shelter Current Table
        fs_current_data = [[df['Unnamed: 6'][0], current_month_fs],
                           [df['Unnamed: 8'][0], current_month_budget_fs],
                           [df['Unnamed: 10'][0], current_month_obudget_fs]
                           ]

        reb_exp_analyze_window = Tk()
        reb_exp_analyze_window.attributes('-fullscreen', True)
        reb_exp_analyze_window.title(
            "Revenue and Expenses Compared with Budget Summary")

        fig, ax = plt.subplots(1, 1)
        ax.axis('off')
        ax.axis('tight')
        plt.title("Current Food Shelter Budget", fontsize=16)

        fig.set_size_inches(6, 6)
        cfs_table = ax.table(fs_current_data, loc='center')
        cfs_table.scale(1, 7)
        cfs_table.set_fontsize(16)
        canvas = FigureCanvasTkAgg(fig, master=reb_exp_analyze_window)
        canvas.draw()
        canvas.get_tk_widget().place(relx=.12, rely=.22, anchor=CENTER)

        # Creation of Food Shelter YTD Table
        fs_ytd_data = [[df['Unnamed: 12'][0], current_months_budget_fs],
                       [df['Unnamed: 14'][0], current_ytd_budget_fs],
                       [df['Unnamed: 16'][0], current_ytd_obudget_fs],
                       [df['Unnamed: 18'][0], current_ytd_abudget_fs]]

        fig, ax = plt.subplots(1, 1)
        ax.axis('off')
        ax.axis('tight')
        plt.title("YTD Food Shelter Budget", fontsize=16)

        ytd_fs_table = ax.table(fs_ytd_data, loc='center')
        fig.set_size_inches(6, 7)
        ytd_fs_table.scale(1, 4.5)
        ytd_fs_table.set_fontsize(16)
        canvas2 = FigureCanvasTkAgg(fig, master=reb_exp_analyze_window)
        canvas2.draw()
        canvas2.get_tk_widget().place(relx=.12, rely=.7, anchor=CENTER)

        # Creation of NOW Current Table

        now_current_data = [[df['Unnamed: 6'][0], current_month_now],
                            [df['Unnamed: 8'][0], current_month_budget_now],
                            [df['Unnamed: 10'][0], current_month_obudget_now]
                            ]

        fig, ax = plt.subplots(1, 1)
        ax.axis('off')
        ax.axis('tight')
        plt.title("Current NOW Budget", fontsize=16)
        fig.set_size_inches(6, 6)
        table_data = ax.table(now_current_data, loc='center')
        table_data.set_fontsize(16)
        table_data.scale(1, 7)
        canvas3 = FigureCanvasTkAgg(fig, master=reb_exp_analyze_window)
        canvas3.draw()
        canvas3.get_tk_widget().place(relx=.4, rely=.22, anchor=CENTER)

        # Create of NOW YTD Table
        now_ytd_data = [[df['Unnamed: 12'][0], current_year_budget_now],
                        [df['Unnamed: 14'][0], current_ytd_budget_now],
                        [df['Unnamed: 16'][0], current_ytd_obudget_now],
                        [df['Unnamed: 18'][0], current_ytd_abudget_now]]

        fig, ax = plt.subplots(1, 1)
        ax.axis('off')
        ax.axis('tight')
        plt.title("YTD NOW Budget", fontsize=16)
        fig.set_size_inches(6, 7)
        ytd_now_table = ax.table(now_ytd_data, loc='center')
        ytd_now_table.scale(1, 7.5)
        ytd_now_table.set_fontsize(16)
        canvas4 = FigureCanvasTkAgg(fig, master=reb_exp_analyze_window)
        canvas4.draw()
        canvas4.get_tk_widget().place(relx=.4, rely=.7, anchor=CENTER)

        # Creation of Mobile Show Current Table
        ms_current_data = [[df['Unnamed: 6'][0], current_month_ms],
                           [df['Unnamed: 8'][0], current_month_budget_ms],
                           [df['Unnamed: 10'][0], current_month_obudget_ms]
                           ]

        fig, ax = plt.subplots(1, 1)
        ax.axis('off')
        ax.axis('tight')
        plt.title("Current Mobile Shower Budget", fontsize=16)
        fig.set_size_inches(6, 6)
        cms_table = ax.table(ms_current_data, loc='center')
        cms_table.set_fontsize(16)
        cms_table.scale(1, 7)
        canvas5 = FigureCanvasTkAgg(fig, master=reb_exp_analyze_window)
        canvas5.draw()
        canvas5.get_tk_widget().place(relx=.7, rely=.22, anchor=CENTER)

        # Create YTD Mobile Shower Table
        ms_ytd_data = [[df['Unnamed: 12'][0], current_year_budget_ms],
                       [df['Unnamed: 14'][0], current_ytd_budget_ms],
                       [df['Unnamed: 16'][0], current_ytd_obudget_ms],
                       [df['Unnamed: 18'][0], current_ytd_abudget_ms]]

        fig, ax = plt.subplots(1, 1)
        ax.axis('off')
        ax.axis('tight')
        plt.title("YTD Mobile Shower Budget", fontsize=16)
        fig.set_size_inches(6, 7)
        ytd_ms_table = ax.table(ms_ytd_data, loc='center', )
        ytd_ms_table.scale(1, 7.5)
        ytd_ms_table.set_fontsize(16)
        canvas6 = FigureCanvasTkAgg(fig, master=reb_exp_analyze_window)
        canvas6.draw()
        canvas6.get_tk_widget().place(relx=.7, rely=.7, anchor=CENTER)

        revenues_expsenses_budget_window.withdraw()

        btn_main = Button(reb_exp_analyze_window, text="Back to Main", command=lambda: [reb_exp_analyze_window.withdraw(), main_window.deiconify(), plt.close('all')],
                          width=20, height=3, fg='green')
        btn_main.place(x=1700, y=200)

        
        
        
        # WIP
        def pdf_Print():
                    printer_name = win32print.GetDeafaultPrint
                

                    #Grab Filename
                    filename =filedialog.askopenfilename(initialdir="/",
                                                title="Select Excel File",
                                                filetypes=(("Excel",
                                                            "*.xlsx"),
                                                            ("all files",
                                                            "*.*")))
                    if filename:
                        win32api.ShellExecute(0, "print",filename,None,",",0)
        # WIP Button for Printing 'Canvas' and button position
        btn_print = Button(reb_exp_analyze_window, text="Print as PDF", command=lambda: [pdf_Print()],
                           width=20, height=3, fg='green')

        btn_print.place(x=1700, y=100)


        # method to capture GUI image, saves as png, and converts to pdf
        def save_data():

            screenshottaker = pyautogui.screenshot()
            save_path = filedialog.asksaveasfilename()
            screenshottaker.save(save_path+".png")

            saved_image = Image.open(save_path+".png")
            im_1 = saved_image.convert('RGB')
            im_1.save(save_path+".pdf")
            os.remove(save_path+".png")

        # WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(reb_exp_analyze_window, text="Save as PDF", command=lambda: [save_data()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=1700, y=300)

        btn_exit = Button(reb_exp_analyze_window, text="Exit",
                          command=lambda: [os._exit(0)], width=20, height=3, fg='green')
        btn_exit.place(x=1700, y=400)

        reb_exp_analyze_window.mainloop()

    btn_reb_analyze = Button(revenues_expsenses_budget_window, text="Analyze File", command=analyze_revenue_expneses_budget,
                             width=20, height=3, fg='green')
    btn_reb_analyze.place(x=700, y=550)

    revenues_expsenses_budget_window.mainloop()


# Creates Button on main window for 'Statement of Revenue and Expense Budget' Option
btn_Revenue_Expenses_Budget = Button(main_window, text="Statement of Revenue and Expenses Budget",
                                     command=create_Revenue_Expenses_Budget_frame, width=50, height=3, fg='green')
btn_Revenue_Expenses_Budget.place(x=50, y=600)


# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Revenue Expnese by Program Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates 'Revenue and Expenses by Program' window

def create_Revenue_Expenses_Program_frame():

    revenues_expsenses_program_window = Tk()
    revenues_expsenses_program_window .attributes('-fullscreen', True)

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'Revenue and Expenses by Program' GUI
    revenues_expsenses_program_window.title(
        'Revenue and Expenses by Program Summary File Upload')

    rep_label = Label(revenues_expsenses_program_window,
                      text="Revenue and Expenses by Progam Summary")
    rep_label.config(font=("Courier", 34))
    rep_label.place(x=100, y=100)

    rep_descr = Label(revenues_expsenses_program_window,
                      text="Please select 'Stmnt of Revenues and Expenses by Program' file for analyzing")
    rep_descr.config(font=("Courier", 16))
    rep_descr.place(x=100, y=200)

    # Creates textbox for 'Revenue and Expenses by Program' GUI and sets location on GUI
    re_P_txt_file = Text(revenues_expsenses_program_window, height=4,
                         width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    re_P_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(revenues_expsenses_program_window, text="Back to Main", command=lambda: [revenues_expsenses_program_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'Revenue and Expenses by Program' file for upload and location on 'AR Summary Frame' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(revenues_expsenses_program_window, text="Find File", command=lambda: [re_P_txt_file.configure(state='normal'), re_P_txt_file.delete('1.0', "end"),
                                                                                                   browseFiles(), re_P_txt_file.insert(INSERT, filepath),
                                                                                                   re_P_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    # fucntion to anayle 'Revenue and Expenses' Excel
    def analyze_revenue_expenses_program():
        df = pd.read_excel(
            filepath, sheet_name=1)

        revenues_expsenses_program_window.withdraw()

        # Capture of data points for Excel file
        total_income_food_sack = (df['Food Sack'][38])
        total_income_mobile_shower = (df['Mobile Shower Program'][38])
        total_income_now = (df['NOW'][38])

        total_expenses_food_sack = (df['Food Sack'][160])
        total_expenses_mobile_shower = (df['Mobile Shower Program'][160])
        total_expenses_now = (df['NOW'][160])

        net_income_food_sack = (df['Food Sack'][167])
        net_income_mobile_shower = (df['Mobile Shower Program'][167])
        net_income_now = (df['NOW'][167])

        rep_analyze_window = Tk()
        rep_analyze_window.attributes('-fullscreen', True)
        rep_analyze_window.title("Revenue and Expenses by Program Summary")

        sns.set_theme(font_scale=1.5)

        # Creates 'Total Income' Bar Graph
        def create_total_income_bar():

            sns.set_palette(['orange', 'purple', 'grey'])
            ti_axs = sns.catplot(x=['Food Sack', 'Mobile Shower', 'NOW'], y=[total_income_food_sack, total_income_mobile_shower, total_income_now],
                                 data=df, kind='bar')
            ti_axs.fig.suptitle('Income by Program')
            plt.xlabel("Programs")
            plt.ylabel("$ Amount")
            plt.axis('tight')
            plt.ylim(0, None)
            ax = ti_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            ti_axs.fig.set_size_inches(11, 7.5)

            return ti_axs.fig

        ta_fig = create_total_income_bar()
        canvas = FigureCanvasTkAgg(ta_fig, master=rep_analyze_window)
        canvas.draw()

        canvas.get_tk_widget().place(relx=.01, rely=.01,)

        # Creates total expenses bar graph
        def create_total_expenses_bar():
            sns.set_palette(['red', 'green', 'teal'])
            te_axs = sns.catplot(x=['Food Sack', 'Mobile Shower', 'NOW'], y=[total_expenses_food_sack, total_expenses_mobile_shower, total_expenses_now],
                                 data=df, kind='bar')
            te_axs.fig.suptitle('Expenses by Program')
            plt.xlabel("Programs")
            plt.ylabel("$ Amount")

            plt.axis('tight')
            plt.ylim(0, None)
            ax = te_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            te_axs.fig.set_size_inches(10, 7)

            return te_axs.fig

        tl_fig = create_total_expenses_bar()
        canvas = FigureCanvasTkAgg(tl_fig, master=rep_analyze_window)
        canvas.draw()

        canvas.get_tk_widget().place(relx=.45, rely=.01,)

        # Creates net income bar graph
        def create_net_income_bar():
            sns.set_palette(['teal', 'grey', 'brown'])
            ni_axs = sns.catplot(x=['Food Sack', 'Mobile Shower', 'NOW'], y=[net_income_food_sack, net_income_mobile_shower, net_income_now],
                                 data=df, kind='bar')
            ni_axs.fig.suptitle('Net Income by Program')
            plt.xlabel("Programs")
            plt.ylabel("$ Amount")
            plt.ylim(0, None)
            plt.axis('tight')
            ax = ni_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            ni_axs.fig.set_size_inches(11, 7)

            return ni_axs.fig

        te_fig = create_net_income_bar()
        canvas = FigureCanvasTkAgg(te_fig, master=rep_analyze_window)
        canvas.draw()
        canvas.get_tk_widget().place(relx=.01, rely=.525,)

        btn_main = Button(rep_analyze_window, text="Back to Main", command=lambda: [rep_analyze_window.withdraw(), main_window.deiconify(), plt.close('all')],
                          width=20, height=3, fg='green')
        btn_main.place(x=1700, y=300)

       
       
       # WIP
        def pdf_Print():
                    printer_name = win32print.GetDeafaultPrint
                

                    #Grab Filename
                    filename =filedialog.askopenfilename(initialdir="/",
                                                title="Select Excel File",
                                                filetypes=(("Excel",
                                                            "*.xlsx"),
                                                            ("all files",
                                                            "*.*")))
                    if filename:
                        win32api.ShellExecute(0, "print",filename,None,",",0)
        # WIP Button for Printing 'Canvas' and button position
        btn_print = Button(rep_analyze_window, text="Print as PDF",command=lambda: [pdf_Print()],
                           width=20, height=3, fg='green')

        btn_print.place(x=1700, y=100)
        # method to capture GUI image, saves as png, and converts to pdf
        def save_data():

            screenshottaker = pyautogui.screenshot()
            save_path = filedialog.asksaveasfilename()
            screenshottaker.save(save_path+".png")

            saved_image = Image.open(save_path+".png")
            im_1 = saved_image.convert('RGB')
            im_1.save(save_path+".pdf")
            os.remove(save_path+".png")

        # WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(rep_analyze_window, text="Save as PDF", command=lambda: [save_data()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=1700, y=200)

        btn_exit = Button(rep_analyze_window, text="Exit",
                          command=lambda: [os._exit(0)], width=20, height=3, fg='green')
        btn_exit.place(x=1700, y=400)

        rep_analyze_window.mainloop()

    btn_rep_analyze = Button(revenues_expsenses_program_window, text="Analyze File", command=analyze_revenue_expenses_program,
                             width=20, height=3, fg='green')
    btn_rep_analyze.place(x=700, y=550)

    revenues_expsenses_program_window.mainloop()


# Creates Button on main window for 'Statement oof Revenue and Expenses by Program' Option
btn_Revenue_Expenses_Program = Button(main_window, text="Statement of Revenue and Expenses by Program",
                                      command=create_Revenue_Expenses_Program_frame, width=50, height=3, fg='green')
btn_Revenue_Expenses_Program.place(x=50, y=700)


# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Main Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Exit button to close program
btn_exit = Button(main_window, text="Exit",
                  command=lambda: [os._exit(0)], width=50, height=3, fg='green')
btn_exit.place(x=1150, y=900)

# loads Pueblo Cooperative Care image
# Will need to be changed so can be accessed globally
load = Image.open(
    "Pueblo-Coop-Center-History.gif")
render = ImageTk.PhotoImage((load))
photo = Label(main_window, image=render)
photo.place(x=600, y=100)


main_window.mainloop()
