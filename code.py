import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl


def enter_data():
    accepted = accept_var.get()

    if accepted == "Accepted":
        # User info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()

        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            # Course info
            registration_status = reg_status_var.get()
            numcourses = numcourses_spinbox.get()
            numsemesters = numsemesters_spinbox.get()

            print("First name: ", firstname, "Last name: ", lastname)
            print("Title: ", title, "Age: ", age, "Nationality: ", nationality)
            print("# Courses: ", numcourses, "# Semesters: ", numsemesters)
            print("Registration status", registration_status)
            print("------------------------------------------")

            filepath = "C:\\Users\\Welcome\\Downloads\\Form Details.xlsx"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First Name", "Last Name", "Title", "Age", "Nationality",
                           "# Courses", "# Semesters", "Registration status"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, title, age, nationality, numcourses,
                          numsemesters, registration_status])
            workbook.save(filepath)

        else:
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
    else:
        tkinter.messagebox.showwarning(title="Error", message="You have not accepted the terms")


window = tkinter.Tk()
window.title("Data Entry Form")

# Set the window background color
window.configure(bg="#FFFF66")  # Yellow background

frame = tkinter.Frame(window, bg="#FFFF66")
frame.pack()

# Saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="User Information", bg="#FFFF66", fg="black")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="First Name", bg="#FFFF66", fg="black")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(user_info_frame, text="Last Name", bg="#FFFF66", fg="black")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

title_label = tkinter.Label(user_info_frame, text="Title", bg="#FFFF66", fg="black")
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms.", "Dr."])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

age_label = tkinter.Label(user_info_frame, text="Age", bg="#FFFF66", fg="black")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tkinter.Label(user_info_frame, text="Nationality", bg="#FFFF66", fg="black")
nationality_combobox = ttk.Combobox(user_info_frame,
                                    values=["Afghan", "Albanian", "Algerian", "American", "Andorran", "Angolan", "Antiguan and Barbudan", "Argentine", "Armenian", "Australian", "Austrian", "Azerbaijani", "Bahamian", "Bahraini", "Bangladeshi", "Barbadian", "Belarusian", "Belgian", "Belizean", "Beninese", "Bhutanese", "Bolivian", "Bosnian", "Botswana", "Brazilian", "Bruneian", "Bulgarian", "Burkinabe", "Burundian", "Cabo Verdean", "Cambodian", "Cameroonian", "Canadian", "Central African", "Chadian", "Chilean", "Chinese", "Colombian", "Comoran", "Congolese", "Costa Rican", "Croatian", "Cuban", "Cypriot", "Czech", "Danish", "Djiboutian", "Dominican", "Dominican Republic", "Ecuadorian", "Egyptian", "Salvadoran", "Equatorial Guinean", "Eritrean", "Estonian", "Eswatini", "Ethiopian", "Fijian", "Finnish", "French", "Gabonese", "Gambian", "Georgian", "German", "Ghanaian", "Greek", "Grenadian", "Guatemalan", "Guinean", "Bissau-Guinean", "Guyanese", "Haitian", "Honduran", "Hungarian", "Icelandic", "Indian", "Indonesian", "Iranian", "Iraqi", "Irish", "Israeli", "Italian", "Ivorian", "Jamaican", "Japanese", "Jordanian", "Kazakh", "Kenyan", "Kiribati", "North Korean", "South Korean", "Kuwaiti", "Kyrgyz", "Lao", "Latvian", "Lebanese", "Lesotho", "Liberian", "Libyan", "Liechtensteiner", "Lithuanian", "Luxembourgish", "Macedonian", "Madagascan", "Malawian", "Malaysian", "Maldivian", "Malian", "Maltese", "Marshallese", "Mauritanian", "Mauritian", "Mexican", "Micronesian", "Moldovan", "Monacan", "Mongolian", "Montenegrin", "Moroccan", "Mozambican", "Namibian", "Nauruan", "Nepali", "Dutch", "New Zealander", "Nicaraguan", "Nigerien", "Nigerian", "Niuean", "Norwegian", "Omani", "Pakistani", "Palauan", "Palestinian", "Panamanian", "Papua New Guinean", "Paraguayan", "Peruvian", "Philippine", "Polish", "Portuguese", "Qatari", "Romanian", "Russian", "Rwandan", "Saint Kitts and Nevis", "Saint Lucian", "Saint Vincentian", "Samoan", "San Marino", "Sao Tomean", "Saudi Arabian", "Senegalese", "Serbian", "Seychellois", "Sierra Leonean", "Singaporean", "Slovak", "Slovenian", "Solomon Islander", "Somali", "South African", "South Sudanese", "Spanish", "Sri Lankan", "Sudanese", "Surinamese", "Swazi", "Swedish", "Swiss", "Syrian", "Taiwanese", "Tajikistani", "Tanzanian", "Thai", "Togolese", "Tongan", "Trinidadian and Tobagonian", "Tunisian", "Turkish", "Turkmen", "Tuvaluan", "Ugandan", "Ukrainian", "Uruguayan", "Uzbekistani", "Vanuatuan", "Venezuelan", "Vietnamese", "Yemeni", "Zambian", "Zimbabwean"])
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Saving Course Info
courses_frame = tkinter.LabelFrame(frame, bg="#FFFF66", fg="black")
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = tkinter.Label(courses_frame, text="Registration Status", bg="#FFFF66", fg="black")

reg_status_var = tkinter.StringVar(value="Not Registered")
registered_check = tkinter.Checkbutton(courses_frame, text="Currently Registered", bg="#FFFF66", fg="black",
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not registered")

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

numcourses_label = tkinter.Label(courses_frame, text="# Completed Courses", bg="#FFFF66", fg="black")
numcourses_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = tkinter.Label(courses_frame, text="# Semesters", bg="#FFFF66", fg="black")
numsemesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to="infinity")
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions", bg="#FFFF66", fg="black")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms and conditions.", bg="#FFFF66", fg="black",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

# Button
button = tkinter.Button(frame, text="Enter data", command=enter_data, bg="#990011", fg="white")
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

window.mainloop()
