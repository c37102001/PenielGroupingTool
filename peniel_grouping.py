from math import ceil
from random import shuffle
from pandas import read_excel, DataFrame, ExcelWriter
from tkinter import *
from tkinter import ttk
from os import startfile

EXCEL_INPUT_FILE = './group_theme.xlsx'
EXCEL_OUTPUT_FILE = './group_result.xlsx'
LAUNCH_EXCEL_FILE = 'group_result.xlsx'
LAUNCH_PPT_FILE = 'display_template.pptx'

def start_grouping():
    LEAD_S = int(leader_from_entry.get())
    LEAD_E = int(leader_to_entry.get())
    HELP_S = int(helper_from_entry.get())
    HELP_E = int(helper_to_entry.get())
    ATTEND_S = int(attender_from_entry.get())
    ATTEND_E = int(attender_to_entry.get())
    PEOPLE_PER_GROUP = int(group_people_entry.get())
    ALLOW_ONE_MORE = bool(remainder.get())
    THEME = str(theme_combobox.get())

    # print(LEAD_S, LEAD_E)
    # print(HELP_S, HELP_E)
    # print(ATTEND_S, ATTEND_E)
    # print(PEOPLE_PER_GROUP, ALLOW_ONE_MORE, THEME)

    leaders = [i for i in range(LEAD_S, LEAD_E + 1)]
    helpers = [i for i in range(HELP_S, HELP_E + 1)]
    attenders = [i for i in range(ATTEND_S, ATTEND_E + 1)]
    shuffle(leaders)
    shuffle(helpers)
    shuffle(attenders)
    all_people = leaders + helpers + attenders
    people_num = len(all_people)
    group_num = people_num // PEOPLE_PER_GROUP if ALLOW_ONE_MORE else ceil(people_num / PEOPLE_PER_GROUP)

    # grouping
    dataset = read_excel(EXCEL_INPUT_FILE)
    theme = dataset[THEME]
    group_list = [[] for _ in range(group_num)]
    for i, people in enumerate(all_people):
        group_list[i % group_num].append([people, theme[people - 1]])

    for team_id, group in enumerate(group_list, start=1):
        group.sort(key = lambda s: s[0])
        for i in range(len(group)):
            group[i][0] = str(group[i][0])
            group[i] = '\n'.join(group[i])
        group.insert(0, "第%d組" % team_id)

    info = "\n總共 %d人, 分成 %d組" % (people_num, group_num)
    info_label.configure(text=info)

    # ================================

    df = DataFrame()
    writer = ExcelWriter(EXCEL_OUTPUT_FILE, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    cell_format = workbook.add_format()
    cell_format.set_bold()
    cell_format.set_font_size(18)
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    cell_format.set_text_wrap()

    for col, group in enumerate(group_list):
        for row, value in enumerate(group):
            worksheet.write(row, col, value, cell_format)
    writer.save()

    startfile(LAUNCH_EXCEL_FILE)
    startfile(LAUNCH_PPT_FILE)


# ===================================== GUI ======================================

root = Tk()
root.title('Peniel Youth Grouping Tool')
root.geometry('{}x{}'.format(500, 300))

# create all of the main containers
top_frame = Frame(root, width=450, padx=10, pady=10)
mid_frame = Frame(root, width=450, padx=10)    

# layout all of the main containers
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

top_frame.grid(row=0, sticky="nsew")
mid_frame.grid(row=1, sticky="nsew")

# =========== Card Info ==============
model_label = Label(top_frame, text='Card Info')
model_label.grid(row=0)

# Leaders
leader_from_label = Label(top_frame, text='Leaders from')
leader_from_label.grid(row=1, column=0)
leader_from_entry = Entry(top_frame)
leader_from_entry.insert(END, "1")
leader_from_entry.grid(row=1, column=1, padx=(10, 0))

leader_to_label = Label(top_frame, text=' to ')
leader_to_label.grid(row=1, column=2)
leader_to_entry = Entry(top_frame)
leader_to_entry.grid(row=1, column=3)

# Helpers
helper_from_label = Label(top_frame, text='Helpers from')
helper_from_label.grid(row=2, column=0)
helper_from_entry = Entry(top_frame)
helper_from_entry.grid(row=2, column=1, padx=(10, 0))

helper_to_label = Label(top_frame, text=' to ')
helper_to_label.grid(row=2, column=2)
helper_to_entry = Entry(top_frame)
helper_to_entry.insert(END, "15")
helper_to_entry.grid(row=2, column=3)

# Attenders
attender_from_label = Label(top_frame, text='Attenders from')
attender_from_label.grid(row=3, column=0)
attender_from_entry = Entry(top_frame)
attender_from_entry.insert(END, "16")
attender_from_entry.grid(row=3, column=1, padx=(10, 0))

attender_to_label = Label(top_frame, text=' to ')
attender_to_label.grid(row=3, column=2)
attender_to_entry = Entry(top_frame)
attender_to_entry.grid(row=3, column=3)

# ============= Group Info ===============

# Theme Combobox
themes = [t for t in read_excel(EXCEL_INPUT_FILE).columns[1:]]
theme_label = Label(top_frame, text='Group Theme')
theme_label.grid(row=4, column=0, pady=(15, 0))
theme_combobox = ttk.Combobox(top_frame, values=themes, state="readonly")
theme_combobox.grid(row=4, column=1, padx=(10, 0), pady=(15, 0))
theme_combobox.current(0)
# print(theme_combobox.current(), theme_combobox.get())

# Number of people per group
group_people_label = Label(top_frame, text='Number of People\nper Group')
group_people_label.grid(row=5, column=0, pady=(15, 0))
group_people_entry = Entry(top_frame)
group_people_entry.insert(END, "4")
group_people_entry.grid(row=5, column=1, pady=(15, 0))

# One more or less
remainder_label = Label(mid_frame, text='Remainder')
remainder_label.grid(row=6, column=0, padx=(15, 0), pady=(5, 0))
remainder = IntVar()
remainder_btn1 = Radiobutton(mid_frame, text="Merge into bigger", variable=remainder, value=1)
remainder_btn1.grid(row=6, column=1, padx=(30, 0), pady=(15, 0))
remainder_btn2 = Radiobutton(mid_frame, text="Split into smaller ", variable=remainder, value=0)
remainder_btn2.grid(row=6, column=2, padx=(10, 0), pady=(15, 0))
remainder.set(1)

# Show final info
info_label = Label(mid_frame)
info_label.grid(row=7, column=0, columnspan=3, padx=(15, 0), sticky=W)

# Start button
start_button = Button(mid_frame, text='Start Grouping', bg='pink', fg='red', command=start_grouping)
start_button.grid(row=7, column=4, pady=(15, 0))

root.mainloop()