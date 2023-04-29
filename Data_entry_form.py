import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl

window = tk.Tk()
window.title('Data Entry Form')

frame = tk.Frame(window)
frame.pack()


def submit_data():
    # Status of Terms & Conditions
    status = terms_status.get()

    if status == 'Accepted Terms & Conditions':

        # User info
        firstname = first_name_enrty.get()
        lastname = last_name_enrty.get()

        if firstname and lastname:
            title = titlecombobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            # Course and semester info
            courses = course_completed_spinbox.get()
            semester = semesters_spinbox.get()
            regstatus = reg_status.get()

            # print('Title: ' + title + '    First Name: ' + firstname + '    Last Name: ' + lastname)
            # print('Age: ' + age + '    Nationality: ' + nationality)
            # print('Courses Done: ' + courses + '    Semester Done: ' + semester)
            # print('Registration Status: ' + regstatus)
            # print('< < < < < < < < < < < < < < < < < < < < < < < > > > > > > > > > > > > > > > > > > > > > > ')

            filepath_way = r"C:\Users\Administrator\PycharmProjects\pythonProject\Dara.xlsx"

            # Opening Excel workbook
            open_workbook = openpyxl.load_workbook(filepath_way)
            active_sheet = open_workbook.active

            # Creating heading to the workbook
            active_sheet['A1'] = 'First Name'
            active_sheet['B1'] = 'Last Name'
            active_sheet['C1'] = 'Title'
            active_sheet['D1'] = 'Age'
            active_sheet['E1'] = 'Nationality'
            active_sheet['F1'] = 'Courses Done'
            active_sheet['G1'] = 'Semester'
            active_sheet['H1'] = 'Reg Status'

            # Append data to excel worksheet
            row = (firstname, lastname, title, age, nationality, courses, semester, regstatus)
            active_sheet.append(row)

            # Saving Excel workbook
            open_workbook.save(filepath_way)

            # Reset entry values
            first_name_enrty.delete(0, tk.END)
            last_name_enrty.delete(0, tk.END)
            titlecombobox.delete(0, tk.END)
            course_completed_spinbox.delete(0, tk.END)
            semesters_spinbox.delete(0, tk.END)
            nationality_combobox.current(87)
            age_spinbox.delete(0, tk.END)
            reg_status.set(value='Not Registered')
            terms_status.set(value='Not Accepted Terms & Conditions')


        else:
            tk.messagebox.showwarning(title='Error', message="Yoh‼ First Name and Last Name can't Empty")
    else:
        tk.messagebox.showwarning(title='Error', message='Invalid‼ Terms & Conditions Not Checked')


# Saving the user info
user_info_frame = tk.LabelFrame(frame, text='User Information')
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

# User info labels
first_name_label = tk.Label(user_info_frame, text='First Name')
first_name_label.grid(row=0, column=0)
last_name_label = tk.Label(user_info_frame, text='Last Name')
last_name_label.grid(row=0, column=1)
title_label = tk.Label(user_info_frame, text='Title')
title_label.grid(row=0, column=2)
age_label = tk.Label(user_info_frame, text='Age')
age_label.grid(row=2, column=0)
nationality_label = tk.Label(user_info_frame, text='Nationality')
nationality_label.grid(row=2, column=1)

# User entry widgets for first and last name
first_name_enrty = tk.Entry(user_info_frame)
first_name_enrty.grid(row=1, column=0)
last_name_enrty = tk.Entry(user_info_frame)
last_name_enrty.grid(row=1, column=1)

# Combobox for title
titlecombobox = ttk.Combobox(user_info_frame, values=['Mr', 'Mrs'])
titlecombobox.grid(row=1, column=2)

# Spin box for age
age_spinbox = tk.Spinbox(user_info_frame, from_=18, to=118)
age_spinbox.grid(row=3, column=0, padx=20, pady=0)

# Combobox for nationality
nationality_combobox = ttk.Combobox(user_info_frame)
# List of countries to be used in the combobox
countries = ['Afghanistan', 'Albania', 'Algeria', 'Andorra', 'Angola', 'Antigua & Deps', 'Argentina', 'Armenia',
             'Australia', 'Austria', 'Azerbaijan', 'Bahamas', 'Bahrain', 'Bangladesh', 'Barbados', 'Belarus', 'Belgium',
             'Belize', 'Benin', 'Bhutan', 'Bolivia', 'Bosnia Herzegovina', 'Botswana', 'Brazil', 'Brunei', 'Bulgaria',
             'Burkina', 'Burundi', 'Cambodia', 'Cameroon', 'Canada', 'Cape Verde', 'Central African Rep', 'Chad',
             'Chile', 'China', 'Colombia', 'Comoros', 'Congo', 'Congo {Democratic Rep}', 'Costa Rica', 'Croatia',
             'Cuba', 'Cyprus', 'Czech Republic', 'Denmark', 'Djibouti', 'Dominica', 'Dominican Republic', 'East Timor',
             'Ecuador', 'Egypt', 'El Salvador', 'Equatorial Guinea', 'Eritrea', 'Estonia', 'Ethiopia', 'Fiji',
             'Finland', 'France', 'Gabon', 'Gambia', 'Georgia', 'Germany', 'Ghana', 'Greece', 'Grenada', 'Guatemala',
             'Guinea', 'Guinea-Bissau', 'Guyana', 'Haiti', 'Honduras', 'Hungary', 'Iceland', 'India', 'Indonesia',
             'Iran', 'Iraq', 'Ireland {Republic}', 'Israel', 'Italy', 'Ivory Coast', 'Jamaica', 'Japan', 'Jordan',
             'Kazakhstan', 'Kenya', 'Kiribati', 'Korea North', 'Korea South', 'Kosovo', 'Kuwait', 'Kyrgyzstan', 'Laos',
             'Latvia', 'Lebanon', 'Lesotho', 'Liberia', 'Libya', 'Liechtenstein', 'Lithuania', 'Luxembourg',
             'Macedonia', 'Madagascar', 'Malawi', 'Malaysia', 'Maldives', 'Mali', 'Malta', 'Marshall Islands',
             'Mauritania', 'Mauritius', 'Mexico', 'Micronesia' 'Moldova', 'Monaco', 'Mongolia', 'Montenegro', 'Morocco',
             'Mozambique', 'Myanmar, {Burma}', 'Namibia', 'Nauru', 'Nepal', 'Netherlands', 'New Zealand', 'Nicaragua',
             'Niger', 'Nigeria', 'Norway', 'Oman', 'Pakistan', 'Palau', 'Panama', 'Papua New Guinea', 'Paraguay',
             'Peru', 'Philippines', 'Poland', 'Portugal', 'Qatar', 'Romania', 'Russian Federation', 'Rwanda',
             'St Kitts & Nevis', 'St Lucia', 'Saint Vincent & the Grenadines', 'Samoa', 'San Marino',
             'Sao Tome & Principe', 'Saudi Arabia', 'Senegal', 'Serbia', 'Seychelles', 'Sierra Leone', 'Singapore',
             'Slovakia', 'Slovenia', 'Solomon Islands', 'Somalia', 'South Africa', 'South Sudan', 'Spain', 'Sri Lanka',
             'Sudan', 'Suriname', 'Swaziland', 'Sweden', 'Switzerland', 'Syria', 'Taiwan', 'Tajikistan', 'Tanzania',
             'Thailand', 'Togo', 'Tonga', 'Trinidad & Tobago', 'Tunisia', 'Turkey', 'Turkmenistan', 'Tuvalu', 'Uganda',
             'Ukraine', 'United Arab Emirates', 'United Kingdom', 'United States', 'Uruguay', 'Uzbekistan', 'Vanuatu',
             'Vatican City', 'Venezuela', 'Vietnam', 'Yemen', 'Zambia', 'Zimbabwe']
nationality_combobox['values'] = countries
nationality_combobox.current(87)  # Default value of the combobox
nationality_combobox.grid(row=3, column=1, padx=20, pady=0)

for widgets in user_info_frame.winfo_children():
    widgets.grid_configure(padx=10, pady=5)

# Frame to organize Reg status, Courses done and number of semesters
courses_frame = tk.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky='news', padx=20, pady=10)

# Labels
registered_label = tk.Label(courses_frame, text='Registration Status')
registered_label.grid(row=0, column=0)

reg_status = tk.StringVar(value='Not Registered')
registered_checkbtn = tk.Checkbutton(courses_frame, text='Currently Registered',
                                     variable=reg_status, onvalue='Registered', offvalue='Not Registered')
registered_checkbtn.grid(row=1, column=0)

course_completed_label = tk.Label(courses_frame, text=' Courses Completed')
course_completed_label.grid(row=0, column=1)
course_completed_spinbox = tk.Spinbox(courses_frame, from_=0, to='infinity')
course_completed_spinbox.grid(row=1, column=1)

semesters_label = tk.Label(courses_frame, text='# Number of semesters')
semesters_label.grid(row=0, column=2)
semesters_spinbox = tk.Spinbox(courses_frame, from_=0, to='infinity')
semesters_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Terms and condition
terms_frame = tk.LabelFrame(frame, text='Terms & Conditions')
terms_frame.grid(row=2, column=0, sticky='news', padx=20, pady=10)

terms_status = tk.StringVar(value='Not Accepted Terms & Conditions')
terms_check = tk.Checkbutton(terms_frame, text='I accept the terms and conditions',
                             variable=terms_status, onvalue='Accepted Terms & Conditions',
                             offvalue='Not Accepted Terms & Conditions')
terms_check.grid(row=0, column=0)

# Buttons
submitbtn = tk.Button(frame, text='Enter data', command=submit_data)
submitbtn.grid(row=3, column=0, padx=20, pady=20, sticky='news')

window.mainloop()
