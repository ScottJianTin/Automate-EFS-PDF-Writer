import unittest
import pandas as pd
from PDFWriter import FillablePdfWriter
from fillpdf import fillpdfs
import os
import fnmatch
import datetime


class TestPdfEfs(unittest.TestCase):

    def test_number_of_pdf_created(self):
        """
        Check if the number of output pdf created is same as the number of rows in EFS excel file
        """
        # Initalize variables
        path = str(os.getcwd())
        output_path = str(os.getcwd())
        efs_template_pdf = 'Trade EFS Template.pdf'  # should be pdf template with all fields values deleted as the
        # values will be overwritten
        output_file_name = 'transaction'  # without '.pdf'
        efs_data_excel_file = 'Example EFS data - with_commodity_code_and_transaction_type.xlsx'

        dirpath = 'C:/Users/jiantin/PycharmProjects/PDF'
        output_pdf_count = 0

        # Call class - FillablePdfWriter
        fillable_pdf_writer = FillablePdfWriter()

        # Call method - run_FillablePdfWriter()
        fillable_pdf_writer.run_fillable_pdf_writer(path, efs_data_excel_file, efs_template_pdf, output_path,
                                                    output_file_name)

        # Find the filename string that matches the output_file_name pdf pattern
        output_pdf_files = fnmatch.filter(os.listdir(dirpath), f'{output_file_name}_?.pdf')

        # Count the files if matches the output_file_name pattern
        for output_pdf_file in output_pdf_files:
            output_pdf_count += 1

        # Read excel file
        efs_data_df = pd.read_excel(os.path.join(os.path.abspath(dirpath), efs_data_excel_file))

        # Count the number of rows in dataframe
        efs_data_df_count = len(efs_data_df)

        # Test if number of rows in excel file equals to the number of output pdf files created
        self.assertEqual(efs_data_df_count, output_pdf_count)

    def test_input_efs_data_excel_file(self):
        """
        Manually create a unit test excel file with only 2 rows, compare the output in output pdf with input dictionary
        """
        # Initalize variables
        path = str(os.getcwd())
        # print(f'path: {path}')
        output_path = str(os.getcwd())
        # print(f'output_path: {output_path}')
        efs_template_pdf = 'Trade EFS Template.pdf'  # should be pdf template with all fields values deleted as the
        # values will be overwritten
        output_file_name = 'unit_test'  # use different filename than the one in previous method, first filename
        # created will be unit_test_1.pdf'
        unit_test_efs_data_excel_file = 'data_source_for_unit_test_excel_file.xlsx'

        # Call class - FillablePdfWriter
        fillable_pdf_writer = FillablePdfWriter()

        # # Call method - run_fillable_pdf_writer()
        fillable_pdf_writer.run_fillable_pdf_writer(path, unit_test_efs_data_excel_file, efs_template_pdf,
                                                    output_path, output_file_name)

        # Get the fillable field names from the first output_file_pdf
        pdf_field_name = fillpdfs.get_form_fields(f'{output_file_name}_1.pdf')

        # Create current date and current time variables
        current_date = datetime.datetime.now()
        # print(f'CURRENT DATETIME IN TEST: {current_date}')
        current_date_str = current_date.strftime('%d/%m/%Y')
        current_time_str = current_date.strftime("%I.%M %p")

        # Create input dictionary to compare with field name in output pdf
        input_dictionary = {'Date': current_date_str,
                            'Swaps': '/On',
                            'Member Code': 'S111',
                            'Customer  Account Number': 'SRITRANG INTL / S72449G',
                            'Date  Time of EFP  EFS transaction': f'{current_date_str} at {current_time_str}',
                            'Seller': '/On',
                            'undefined': 'Sicom TSR', # Commodity Description
                            'Commodity Code  Contract Month': 'ORN21', # 'OR' + Month letter code + Year
                            'undefined_2': 'JUL', # Shipment Month
                            'undefined_3': '80', # Total Quantity (lots)
                            'undefined_4': '180.1', # Rounded strike => Price
                            'Fixed Rate Payer  Floating Rate Receiver': 'SRITRANG',
                            'SWAPS': '/On',
                            'Fixed Rate Receiver  Floating Rate Payer': 'BANK',
                            'undefined_12': '29/01/2021', # Start Date
                            'undefined_13': '31/08/2021', # Expirty Date
                            'undefined_14': '180.1', # Rounded strike => Fixed Rate
                            'undefined_16': '720,400', # Notional = Rounded strike * Total Quantity (lots) * 5 * 10
                            }

        for field_name in input_dictionary.keys():
            self.assertEqual(pdf_field_name[field_name], input_dictionary[field_name])

