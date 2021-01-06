from tkinter import *
import tkinter as Tkinter
import tkinter.ttk as ttk
import xlrd
from clusters import *
from PIL import ImageTk, Image
from tkinter import filedialog

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        #Here I declared important variables.
        self.selected_topic_list = None
        self.selected_country_list = None
        self.cleared_countries = []
        self.list_for_sorting = []
        self.selected_countries = []
        self.selected_topics = []
        self.topics = []
        self.viruscount2 = {}
        self.countries = []
        self.pack(fill=X)
        self.initUI()

    def select_file(self, type):
        #This part is for file selection due to my lack of time (yeah lack of time) I couldent able to work out
        #letter_grade_system.txt file so while selecting do not select it because it will open manuelly.
        self.filename = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(
        ("jpeg files", "*.jpg"), ("all files", "*.*"), ("txt file", "*.txt"), ("pdf file", "*.pdf"),
        ("excel file", "*.xlsx")))

        global country_data         # I used global for global using in polymorph function.
        country_data = ""
        global topic_data
        topic_data = ""

        #First Country data will saved after criterias imported listbox's will be updated
        if type == "country":
            self.country_data = self.filename[:]

        if type == "topic":
            self.topic_data = self.filename[:]
            self.polymorph(self.country_data, self.topic_data)

    def initUI(self):
        #Here I wrote the Whole GUI and This part is for Frames I sue 6 frames to make everything clean and workable.
        self.frame0 = Frame(self, relief=GROOVE, borderwidth=2)
        self.frame0.pack(side=TOP, fill=X)

        self.frame1 = Frame(self)
        self.frame1.pack(side=TOP, fill=X)

        self.frame2 = Frame(self)
        self.frame2.pack(side=TOP, padx=5)

        self.frame3 = Frame(self, relief=GROOVE, borderwidth=2)
        self.frame3.pack(side=LEFT, padx=10, pady=15)

        self.frame4 = Frame(self)
        self.frame4.pack(side=LEFT, padx=10, pady=5)

        self.frame5 = Frame(self, relief=GROOVE, borderwidth=2)
        self.frame5.pack(side=LEFT, padx=10, pady=15)

        #############################################  # Here I set scrollbars and Canvas for image importing.

        self.title = Label(self.frame0, text="Coronavirus Data Analysis Tool", bg="red", fg="white",
                           anchor=CENTER, font=('', '20'))
        self.title.pack(side=TOP, fill=X)

        self.main_box = Canvas(self.frame1, width=1030, height=330)
        self.main_box.grid(row=1, column=0, padx=(15, 0))

        self.main_box_sb_Y = Scrollbar(self.frame1, orient="vertical", command=self.main_box.yview)
        self.main_box_sb_Y.grid(row=1, column=1, sticky="ns")

        self.main_box_sb_X = Scrollbar(self.frame1, orient="horizontal", command=self.main_box.xview)
        self.main_box_sb_X.grid(row=2, column=0, sticky="ew", padx=(15, 0))

        self.main_box.configure(yscrollcommand=self.main_box_sb_Y.set, xscrollcommand=self.main_box_sb_X.set)

        ############################################# #This part for file uploading.

        self.country_data_button = Button(self.frame2, text="Upload Country Data", command=lambda: self.select_file("country"))
        self.country_data_button.grid(row=0, column=0, padx=10, pady=15)

        self.test_statistics_button = Button(self.frame2, text="Upload Test Statistics", command=lambda: self.select_file("topic"))
        self.test_statistics_button.grid(row=0, column=1, padx=10, pady=15)

        ############################################# #This part is sorting

        self.label_sort = Label(self.frame3, text="Sort Countries: ")
        self.label_sort.pack(side=TOP, padx=10, pady=10)

        self.sort_name = Button(self.frame3, text="Sort by Name", command=transportation.data_transport_but_names)
        self.sort_name.pack(side=TOP, padx=10, pady=10)

        self.sort_case = Button(self.frame3, text="Sort by Total Case", command=transportation.data_transport)
        self.sort_case.pack(side=TOP, padx=10, pady=10)

        ############################################# #In this Part there are two big listboxes for filtering.

        self.label_countries = Label(self.frame4, text="     Countries:")
        self.label_countries.pack(side=LEFT)

        self.country_list = Listbox(self.frame4, height=20, width=30, selectmode="multiple", exportselection=False)
        self.country_list.pack(side=LEFT)
        self.country_list.bind("<<ListboxSelect>>", self.on_select)

        self.country_list_scrollbar = Scrollbar(self.frame4, orient="vertical", command=self.country_list.yview)
        self.country_list_scrollbar.pack(side=LEFT, fill=Y)

        self.country_list.configure(yscrollcommand=self.country_list_scrollbar.set)

        self.label_criteria = Label(self.frame4, text="     Criterias:")
        self.label_criteria.pack(side=LEFT)

        self.criteria_list = Listbox(self.frame4, height=20, width=30, selectmode="multiple", exportselection=False)
        self.criteria_list.pack(side=LEFT)
        self.criteria_list.bind("<<ListboxSelect>>", self.on_select_for_topics)

        self.criteria_list_scrollbar = Scrollbar(self.frame4, orient="vertical", command=self.criteria_list.yview)
        self.criteria_list_scrollbar.pack(side=LEFT, fill=Y)

        self.criteria_list.configure(yscrollcommand=self.criteria_list_scrollbar.set)

        ############################################# #Here the most important thing is that I called draw_dendongram function from matrix class.

        self.analyse_label = Label(self.frame5, text="Analyse Data:")
        self.analyse_label.pack(side=TOP, padx=10, pady=10)

        self.cluster_country_button = Button(self.frame5, text="Cluster Countries", command=lambda: matrix.draw_dendogram("country"))
        self.cluster_country_button.pack(side=TOP, padx=10, pady=10)

        self.cluster_criteria_button = Button(self.frame5, text="Cluster Criterias", command=lambda: matrix.draw_dendogram("tag"))
        self.cluster_criteria_button.pack(side=TOP, padx=10, pady=10)

    def on_select(self, val):    # In select function is for selecting countries and updating the list all the time so                                             # When I
        self.selected_countries = []     # When I press any item in listbox it refresh the list and keep it updated
        sender = val.widget
        idx = sender.curselection()
        for value in idx:
            self.selected_countries.append(sender.get(value))

    def on_select_for_topics(self, val):  # This is the same thing for criterias.
        self.selected_topics = []
        sender = val.widget
        idx = sender.curselection()
        for value in idx:
            self.selected_topics.append(sender.get(value))

    def excel_reading(self, file, row):      #This part is for excel reading but it's a function used by another function
        wb = xlrd.open_workbook(file)        # I Returned values for usage of another function and I get the country
        sheet = wb.sheet_by_index(0)         # and the criterias that country own.

        direct_data = {}
        country = sheet.cell_value(row, 0)
        try:
            country = country.replace("\xa0", "")
        except:
            pass

        for column in range(1, sheet.ncols):
            direct_data[sheet.cell_value(0, column)] = sheet.cell_value(row, column)

        if country not in self.countries:
            self.countries.append(country)
        return country, direct_data

    def polymorph(self, file1, file2):      # I called this function polymorph because it's using for merging two excel
        self.viruscount1 = {}               # Files datas in my program.
        wb = xlrd.open_workbook(file1)
        sheet = wb.sheet_by_index(0)


        for i in range(1, sheet.nrows):
            country, data = self.excel_reading(file1, i)
            self.viruscount1[country] = data

        self.viruscount2 = self.viruscount1.copy()
        self.viruscount1 = {}
        wb2 = xlrd.open_workbook(file2)
        sheet2 = wb2.sheet_by_index(0)


        for i in range(1, sheet2.nrows):
            country2, data2 = self.excel_reading(file2, i)
            self.viruscount1[country2] = data2

        for key, v in self.viruscount1.items():
            if key in self.viruscount2:
                for key2, value2 in v.items():
                    self.viruscount2[key][key2] = value2   # I created 2 beautifull and usefull lists with this func.
            else:
                self.viruscount2[key] = self.viruscount1[key]

        for k, v in self.viruscount2["Turkey"].items():  # I used turkey because it have all 12 parameters just it.
            self.topics.append(k)
        transportation.start_transport()

class MatrixCreation:
    def create_matrix(self):     # This function is first funciton of matrixcreation class.
        gui.selected_topic_list = gui.selected_topics
        gui.selected_country_list = gui.cleared_countries    # Here I declared selected lists because if we dont select any
        output_file = "output_file.txt"                      # any It should be caunt ass all so that it's purpose.
        with open(output_file, "w") as output:
            output.write("virus")
            if len(gui.selected_topics) == 0:
                gui.selected_topic_list = gui.topics  # This part is for no selection in listboxes.
            if len(gui.selected_countries) == 0:
                gui.selected_country_list = gui.countries

            for word in gui.topics:
                if word in gui.selected_topic_list:
                    output.write("\t%s" % word)     # Here I write the first line of the matrix important thing is that
            output.write("\n")                      # Word should be on selected topics which means if it's not selected
                                                    # In listbox, it will not excist on the matrix as well.
            for country in gui.selected_countries:
                try:
                    a = country.split("(")
                    cleared_country = a[0]
                    gui.selected_country_list.append(cleared_country[:-1])
                except:
                    gui.selected_country_list.append(country)
            for country, data in gui.viruscount2.items():         # Basicly same thing for country and really great loop
                if country in gui.selected_country_list:          # for matrix creation.
                    output.write(country)

                    for word in gui.topics:
                        if word in gui.selected_topic_list and country in gui.selected_country_list:
                            if word in data:
                                if data[word] == "":
                                    output.write("\t0")
                                else:                             # This part is about fillin the rest of the matrix.
                                    output.write("\t%s" % data[word])  # And if it's empty it will fill with 0
                            else:
                                output.write("\t0")
                    output.write("\n")
        return output_file

    def draw_dendogram(self, type):                 # This is another function for drawing the dendogram.
        matrix_file = self.create_matrix()

        countries, tags, data = readfile(matrix_file)

        if type == "country":
            cluster = hcluster(data)
            drawdendrogram(cluster, countries, jpeg="corona_country_cluster.jpg")
            image = Image.open("corona_country_cluster.jpg")
            gui.main_box.image = ImageTk.PhotoImage(image)
            gui.main_box.create_image(0, 0, image=gui.main_box.image, anchor="nw")
            gui.main_box.configure(scrollregion=gui.main_box.bbox('all'))

        elif type == "tag":                      # Main thing in these two function is if it's country it will upload
            rotated_matrix = rotatematrix(data)  # Normal matrix's dendrongram if its a tag (criteria) it will upload
            rotated_hcluster = hcluster(rotated_matrix)          # reverse matrix's dendogram
            drawdendrogram(rotated_hcluster, tags, jpeg="corona_tag_cluster.jpg")
            image = Image.open("corona_tag_cluster.jpg")
            gui.main_box.image = ImageTk.PhotoImage(image)
            gui.main_box.create_image(0, 0, image=gui.main_box.image, anchor="nw")
            gui.main_box.configure(scrollregion=gui.main_box.bbox('all'))


class Transport:                # Transport class contains transport functions and purpose of these functions are
    def start_transport(self):  # Filling the listboxes with countries and their cases. And using for sorting
        for i in gui.topics:
            gui.criteria_list.insert(END, i)
            self.data_transport()

    def data_transport(self):
        gui.country_list.delete(0, 'end')
        for country in gui.countries:
            try:
                gui.country_list.insert(END, country + " (" + str(gui.viruscount2[country]['Total Cases']) + ")")
            except:
                gui.country_list.insert(END, country)

    def data_transport_but_names(self):
        gui.country_list.delete(0, 'end')

        sorted_list = sorted(gui.countries)    # ----------------> Sorted part is using for sorting in the GUI
        for country in sorted_list:
            try:
                gui.country_list.insert(END, country + "(" + str(gui.viruscount2[country]['Total Cases']) + ")")
            except:
                gui.country_list.insert(END, country)

def main():           # This prt I made a main function to start everything and globalled the objects for global usage.
    global matrix
    matrix = MatrixCreation()
    global transportation
    transportation = Transport()
    root = Tk()
    root.geometry("1080x620+220+100")
    global gui
    gui = GUI(root)
    gui.mainloop()

main()
