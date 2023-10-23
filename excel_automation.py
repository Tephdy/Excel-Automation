import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import ttkbootstrap as ttb
import os
import openpyxl


# functions
def enter_data():
    first_name = first_name_entry.get()
    last_name = last_name_entry.get()
    middle_name = middle_name_entry.get()
    age = age_spinbox.get()
    gender = gender_entry.get()
    nationality = nationality_entry.get()
    course = course_entry.get()
    major = major_in_entry.get()
    section = section_entry.get()

    f_first_name = "First Name: " + first_name
    f_last_name = "Last Name: " + last_name
    f_middle_name = "Middle Name: " + middle_name
    f_age = "Age: " + str(age)
    f_gender = "Gender: " + gender
    f_nationality = "Nationality: " + nationality
    f_course = "Course: " + course
    f_major = "Major: " + major
    f_section = "Section: " + section

    user_info = [
        f_first_name,
        f_last_name,
        f_middle_name,
        f_age,
        f_gender,
        f_nationality,
        f_course,
        f_major,
        f_section]

    print("Student Details\n")
    for x in user_info:
        print(x)

    messagebox.showinfo("MESSAGE:","SUBMITTED")

    # file_path = "D:\Documents\Python Pract\TkinterPract1\Student_Info.xlsx"

    file_path=filedialog.asksaveasfilename(
        defaultextension='.xlsx',
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )

    if not os.path.exists(file_path):
        workbook=openpyxl.Workbook()
        sheet=workbook.active
        heading=[
            "First Name",
            "Last Name",
            "Middle Name",
            "Age",
            "Gender",
            "Nationality",
            "Course",
            "Major",
            "Section"]
        sheet.append(heading)
        workbook.save(file_path)
    workbook=openpyxl.load_workbook(file_path)
    sheet=workbook.active
    sheet.append(
        [
            first_name,
            last_name,
            middle_name,
            age,
            gender,
            nationality,
            course,
            major,
            section
        ]
    )
    workbook.save(file_path)

def clear_data():
    first_name_entry.delete('0', 'end')
    last_name_entry.delete('0','end')
    middle_name_entry.delete('0', 'end')
    age_spinbox.delete('0','end')
    gender_entry.delete('0', 'end')
    nationality_entry.delete('0', 'end')
    course_entry.delete('0', 'end')
    major_in_entry.delete('0', 'end')
    section_entry.delete('0', 'end')


# window
window = ttb.Window(themename='darkly')
window.title('Registration Form')
window.geometry('700x550')
window.config(padx=10,pady=10)
window.iconbitmap("icons/icon.ico")

# main label
main_label = ttb.Label(window,text='Registration Form',font='Calibre 14 bold',foreground='orange')
main_label.pack()

# mainframe
main_frame = tk.Frame(window)
main_frame.pack()

# user-info frame
user_info_frame = ttb.LabelFrame(main_frame,text='User Information')
user_info_frame.grid(row=0,
                     column=0,
                     sticky='news',
                     padx=10,
                     pady=10
                     )

# labels
first_name_label = ttb.Label(user_info_frame,
                             text='First Name',
                             foreground='orange')
first_name_label.grid(row=0,column=0)
last_name_label = ttb.Label(user_info_frame,
                            text='Last Name',
                            foreground='orange')
last_name_label.grid(row=0,column=1)
middle_name_label = ttb.Label(user_info_frame,
                              text='Middle Name',
                              foreground='orange')
middle_name_label.grid(row=0,column=2)

# age label
age_label = ttb.Label(user_info_frame,
                      text='Age',
                      foreground='orange')
age_label.grid(row=2,column=0)

# gender label
gender_label = ttb.Label(user_info_frame,
                         text='Gender',
                         foreground='orange')
gender_label.grid(row=2,column=1)

# nationality label
nationality_label = ttb.Label(user_info_frame,
                              text='Nationality',
                              foreground='orange')
nationality_label.grid(row=2,column=2)

# input entry
# first_name_var=tk.StringVar()
first_name_entry = ttb.Entry(user_info_frame,
                             style='warning')
first_name_entry.grid(row=1,column=0)
last_name_var=tk.StringVar()
last_name_entry = ttb.Entry(user_info_frame,
                            style='warning',
                            textvariable=last_name_var)
last_name_entry.grid(row=1,column=1)
middle_name_var=tk.StringVar()
middle_name_entry = ttb.Entry(user_info_frame,
                              style='warning',
                              textvariable=middle_name_var)
middle_name_entry.grid(row=1,column=2)

# age entry
age_var=tk.StringVar()
age_spinbox = ttb.Spinbox(user_info_frame,
                          from_=18,
                          to=110,
                          style='warning',
                          textvariable=age_var)
age_spinbox.grid(row=3,column=0)

# gender entry
gender_var=tk.StringVar()
gender_entry = ttb.Combobox(user_info_frame,
                            values=["", "Male", "Female", "Other"],
                            style='warning',
                            textvariable=gender_var)
gender_entry.grid(row=3,column=1)

# nationality entry
nationality_var=tk.StringVar()
nationality_entry = ttb.Entry(user_info_frame,
                              style='warning',
                              textvariable=nationality_var)
nationality_entry.grid(row=3,column=2)

# configuration
for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10,
                          pady=10)

# course info frame
course_info_frame = ttb.LabelFrame(main_frame,text='Course Information')
course_info_frame.grid(row=1,
                       column=0,
                       sticky='news',
                       padx=10,
                       pady=10)

# course_label
course_label = ttb.Label(course_info_frame,
                         text='Course',
                         foreground='orange')
course_label.grid(row=0,
                  column=0)

# major in label
major_in = ttb.Label(course_info_frame,
                     text="Major",
                     foreground='orange')
major_in.grid(row=0,column=1)

# course entry
course_var=tk.StringVar()
course_entry = ttb.Combobox(course_info_frame,
                            values=["", "BS IT", "BS Computer Science", "BS Computer Engineering"],
                            style='warning',
                            textvariable=course_var)
course_entry.grid(row=1,column=0)

# major entry
major_var=tk.StringVar()
major_in_entry = ttb.Combobox(course_info_frame,
                              values=["", "Web Development", "Graphics Art", "Networking",
                                      "Cyber Security"],
                              style='warning',
                              textvariable=major_var)
major_in_entry.grid(row=1,column=1)

# section
section_label = ttb.Label(course_info_frame,
                          text='Section',
                          foreground='orange',
                          style='warning')
section_label.grid(row=0,column=2)

# section entry
section_var=tk.StringVar()
section_entry = ttb.Combobox(course_info_frame,
                             values=["", "OLSU211E005", "LF211E012"],
                             style="warning",
                             textvariable=section_var)
section_entry.grid(row=1,column=2)

for widget in course_info_frame.winfo_children():
    widget.grid_configure(padx=9,pady=9)

# button frame
button_frame=ttb.LabelFrame(main_frame, text='Submit')
button_frame.grid(row=2,column=0,sticky='news',padx=10,pady=10)

button_sub=ttb.Button(button_frame, text='Submit', command=enter_data, style='warning', width=35)
button_sub.grid(row=2, column=0, padx=10, pady=10)

button_clear=ttb.Button(button_frame, text='Clear', style='warning', width=35, command=clear_data)
button_clear.grid(row=2, column=1, padx=10, pady=10)

for widget in button_frame.winfo_children():
    widget.grid_configure(padx=10,pady=10)


# run
window.mainloop()
