# Import modules
from tkinter import *
from tkinter import ttk, filedialog
import pandas as pd
import os
from PDFWriter import FillablePdfWriter
import subprocess


class PDFWriterApplication(Tk):
    """
    This class is used to create a GUI for FillablePdfWriter class
    """

    def __init__(self):
        Tk.__init__(self)

        # Configure root window
        self.geometry('850x400')
        self.title('PDF Writer Application')
        self.style = ttk.Style().theme_use('clam')
        self.text_font = 'Calibre 10 bold'
        self.button_font = 'Calibre 12 bold'

        # Create input path label
        self.input_path_label = Label(self, text='Input Path:', font=self.text_font)
        self.input_path_label.grid(row=0, column=0, pady=5)
        # Create input path Entry box
        self.input_path_box = Entry(self, width=70)
        self.input_path_box.grid(row=0, column=1, columnspan=4, sticky='ew')

        # Create EFS database filename label
        self.excel_filename_label = Label(self, text='EFS Database Filename (end with .xlsx):', font=self.text_font)
        self.excel_filename_label.grid(row=1, column=0, pady=5)
        # Create EFS database filename Entry box
        self.excel_filename_box = Entry(self, width=70)
        self.excel_filename_box.grid(row=1, column=1, columnspan=4, sticky='ew')

        # Create EFS template PDF file label
        self.template_filename_label = Label(self, text='EFS Template Filename (end with .pdf):', font=self.text_font)
        self.template_filename_label.grid(row=2, column=0, pady=5)
        self.default_template_filename = 'Trade EFS Template.pdf'
        # Create EFS template PDF file Entry box
        self.template_filename_box = Entry(self, width=70)
        self.template_filename_box.insert(0, self.default_template_filename)
        self.template_filename_box.grid(row=2, column=1, columnspan=4, sticky='ew')

        # Create EFS PDF path label
        self.output_path_label = Label(self, text='EFS PDF Output Path:', font=self.text_font)
        self.output_path_label.grid(row=3, column=0, pady=5)
        # Create EFS PDF path Entry box
        self.output_path_box = Entry(self, width=70)
        self.output_path_box.grid(row=3, column=1, columnspan=4, sticky='ew')

        # Create EFS PDF filename label
        self.output_filename_label = Label(self, text='EFS PDF Output Filename (without .pdf):', font=self.text_font)
        self.output_filename_label.grid(row=4, column=0, pady=5, sticky='ew')
        # Create EFS PDF filename Entry box
        self.output_filename_box = Entry(self, width=70)
        self.output_filename_box.grid(row=4, column=1, columnspan=4, sticky='ew')

        # Create open EFS data excel file button
        self.open_excel_button = Button(text='Select excel file', font=self.button_font, fg='white', bg='violet',
                                        command=lambda: self.open_excel_file())
        self.open_excel_button.grid(row=11, column=0, pady=5, ipadx=5)

        # Create a Treeview widget to show the EFS data excel file content if user select the file with
        # open_excel_button
        self.tree = ttk.Treeview(self, height=4)

        # Create open EFS template pds file button
        self.open_template_button = Button(text='Select template file', font=self.button_font, fg='white', bg='purple',
                                           command=lambda: self.open_template_file())
        self.open_template_button.grid(row=11, column=1, pady=5, ipadx=5)

        # Create pdfwriter button
        self.create_pdf_button = Button(text='Create PDF', font=self.button_font, fg='white', bg='green',
                                        command=lambda: self.fillable_pdf_writer())
        self.create_pdf_button.grid(row=11, column=2, pady=5, ipadx=5)

    def open_excel_file(self):
        """
        This function is used in open_excel_button for user to select the EFS data excel file
        :return: create a Treeview (pd.DataFrame) of EFS data excel file if the file exists & automatically fill in the
        Entry box of input path and input filename
        """
        # Get the EFS database excel file location
        input_excel_file = filedialog.askopenfilename(title='Select EFS Data Excel File',
                                                      filetype=(('Excel files', '*.xlsx'), ('All files', '*.')))

        # If selected excel file exists, create a raw string for the filename, else raise the errors
        if input_excel_file:
            try:
                input_excel_file = r'{}'.format(input_excel_file)
                df = pd.read_excel(input_excel_file)
            except ValueError:
                Label(text='File could not be opened')
            except FileNotFoundError:
                Label(text='File not found')

        # Clear all the previous data in tree
        self.clear_treeview()

        # Add new data in Treeview widget
        self.tree['column'] = list(df.columns)
        self.tree['show'] = 'headings'

        # Set headings by iterating over the dataframe columns
        for col in self.tree['column']:
            self.tree.heading(col, text=col)

        # Put data in rows
        df_rows = df.to_numpy().tolist()
        for index, row in enumerate(df_rows):
            self.tree.insert('', 'end', values=row)
            self.tree.column(index, stretch=NO, width=100)

        # Set the grid of Treeview widget
        self.tree.grid(row=5, rowspan=5, columnspan=10, pady=5, padx=10, sticky='nsew')

        # Delete previous filename
        self.input_path_box.delete(0, END)
        self.excel_filename_box.delete(0, END)

        # Split the file location into path and filename
        path = os.path.split(input_excel_file)[0]
        input_excel_file = os.path.split(input_excel_file)[1]

        # Insert path & filename into input_path Entry box & excel_filename Entry box respectively
        self.input_path_box.insert(0, path)
        self.excel_filename_box.insert(0, input_excel_file)

    def clear_treeview(self):
        """
        This function is used in open_excel_function to clear previous data in the tree
        """
        self.tree.delete(*self.tree.get_children())

    def open_template_file(self):
        """
        This function is used in open_template_button for user to select the EFS template pdf file, the selected
        filename will overwrite the default filename in the Entry box
        :return: automatically fill in the Entry box of template filename
        """
        # Get the EFS template pdf file location
        input_template_file = filedialog.askopenfilename(title='Select EFS Data Excel File',
                                                         filetype=(('PDF files', '*.pdf'), ('All files', '*.')))

        # If selected pdf file exists, create a raw string for the filename, else raise the errors
        if input_template_file:
            try:
                input_template_file = r'{}'.format(input_template_file)
            except ValueError:
                Label(text='File could not be opened')
            except FileNotFoundError:
                Label(text='File not found')

        # Delete default filename
        self.template_filename_box.delete(0, END)

        # Extract the filename from the file location
        input_template_file = os.path.split(input_template_file)[1]

        # Insert path & filename into template_filename Entry box
        self.template_filename_box.insert(0, input_template_file)

    def fillable_pdf_writer(self):
        """
        :return: create individual EFS PDF file based on the number of data rows in EFS data excel file
        """
        # Instantiate class FillablePdfWriter
        pdf_writer = FillablePdfWriter()

        # Get all the values from 5 Entry boxes and pass into 5 variables
        path, efs_data_excel_file, efs_template_pdf, output_path, output_file_name = self.get_value()

        # Create several EFS PDF files
        pdf_writer.run_fillable_pdf_writer(path, efs_data_excel_file, efs_template_pdf, output_path, output_file_name)

        # Prompt the feedback if create_pdf_button managed to run successfully
        Label(text='PDF successfully created! Click the button to view the file  ============>', font=self.button_font,
              fg='green').grid(row=10, columnspan=3, pady=5)

        # Create open created PDF window button for user to view the PDF
        open_created_file_button = Button(text='Open created file', font=self.button_font, fg='white', bg='blue',
                                          command=lambda: self.view_created_file())
        open_created_file_button.grid(row=10, column=3, pady=5, ipadx=5)

    def get_value(self):
        """
        This function is used in fillable_pdf_writer function to get the values in 5 Entry boxes
        :return: input values from 5 Entry boxes
        """
        return self.input_path_box.get(), self.excel_filename_box.get(), self.template_filename_box.get(), \
               self.output_path_box.get(), self.output_filename_box.get()

    def view_created_file(self):
        """
        :return: pop out a window for user to choose which created PDF to view
        """

        # Get the filename of created PDF filename
        output_pdf_file = filedialog.askopenfilename(title='Select PDF file to view',
                                                     filetype=(('PDF files', '*.pdf'), ('All files', '*.')))

        # Extract the filename from the file location
        output_pdf_file = str(os.path.split(output_pdf_file)[1])

        # Open the pdf file
        subprocess.Popen([output_pdf_file], shell=True)


if __name__ == '__main__':
    root = PDFWriterApplication()
    root.mainloop()
