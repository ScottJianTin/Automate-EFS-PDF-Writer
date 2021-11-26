# # Import libraries
import os
from fillpdf import fillpdfs
import datetime
import pandas as pd


class FillablePdfWriter:
    """
    This class is used to automatically create several PDF from EFS data excel file
    """

    def run_fillable_pdf_writer(self, path: str, efs_data_excel_file: str, efs_template_pdf: str, output_path: str,
                                output_file_name: str):
        """
        Execute all the functions created in FillablePdfWriter in the class
        :param path: The path location where you store EFS data excel file & EFS template
        (make sure these two files store in the same path)
        :param efs_data_excel_file: The excel file name of EFS data (end with .xlsx)
        :param efs_template_pdf: The pdf file name of EFS template (end with .pdf)
        :param output_path: The directory where you want to create the new PDF files
        :param output_file_name: The pdf file name of EFS that you want to create (*** without '.pdf')
        :return: create individual EFS PDF file based on the number of data rows in EFS data excel file
        """
        # Call method - import_efs_data_excel
        efs_data_df = FillablePdfWriter.create_df_from_import_efs_excel(self, path, efs_data_excel_file)

        # Call method - create_efs_data_list_of_dict_from_df
        efs_data_list = FillablePdfWriter.create_efs_data_list_of_dict_from_df(self, efs_data_df)

        # Call method - fill_pdf
        FillablePdfWriter.fill_pdfs(self, efs_template_pdf, output_path, output_file_name, efs_data_list)

        print('Run successfully! New PDF files created!')

    def create_df_from_import_efs_excel(self, path: str, efs_data_excel_file: str):
        """
        A function that import EFS data excel file, and return a EFS data dataframe after preprocessing
        :param path: The path location where you store EFS data excel file
        :param efs_data_excel_file: The excel file name of EFS data (end with .xlsx)
        :return: A dataframe of EFS data after preprocessing
        """

        # Create efs data dataframe
        # todo: in the future use ExcelHelper and full path
        efs_data_df = pd.read_excel(os.path.join(os.path.abspath(path), efs_data_excel_file))

        # Change start date and expiry date column format
        efs_data_df['Start Date'] = efs_data_df['Start Date'].apply(lambda x: x.strftime('%d/%m/%Y'))
        efs_data_df['Expiry Date'] = efs_data_df['Expiry Date'].apply(lambda x: x.strftime('%d/%m/%Y'))

        # Create shipment month and shipment year column based on the value from shipment column
        efs_data_df['Shipment Month'] = efs_data_df['Shipment'].apply(lambda x: x.strftime('%b').upper())
        efs_data_df['Shipment Year'] = efs_data_df['Shipment'].apply(lambda x: x.strftime('%y'))

        # If notional column does not exist in dataframe, create a notional column with the formula below:
        # Notional = Rounded strike * Quantity * 5 * 10
        if 'Notional' not in efs_data_df.columns:
            efs_data_df['Notional'] = efs_data_df['Rounded strike'] * efs_data_df['Total Quantity (lots)'] * 5 * 10
            efs_data_df['Notional'] = efs_data_df['Notional'].apply(lambda x: "{:,.0f}".format(x))
        else:
            efs_data_df['Notional'] = efs_data_df['Notional'].apply(lambda x: "{:,.0f}".format(x))

        # If commodity code column does not exist in dataframe, create a commodity code column using the formula below:
        # Commodity Code = 'OR' + Shipment Month Letter + last 2 digit of Shipment Year
        if 'Commodity Code' not in efs_data_df.columns:
            # create commodity codes mapper
            commodity_codes_mapper = {'JAN': 'F', 'FEB': 'G', 'MAR': 'H', 'APR': 'J', 'MAY': 'K', 'JUN': 'M',
                                      'JUL': 'N', 'AUG': 'Q', 'SEP': 'U', 'OCT': 'V', 'NOV': 'X', 'DEC': 'Z'}
            # create shipment month letter column from commodity_codes_mapper based on the value in shipment month
            # column
            efs_data_df['Shipment Month Letter'] = efs_data_df['Shipment Month'].map(commodity_codes_mapper)
            # create commonity code column by combining 'OR' + month letter + shipment year
            efs_data_df['Commodity Code'] = 'OR' + efs_data_df['Shipment Month Letter'] + efs_data_df['Shipment Year']

        # If transaction type column exist in dataframe, create a seller and a buyer column with on/off value that is
        # used to tick the checbox in EFS pdf else there is no value for seller/buyer and the checkbox in pdf will
        # not be ticked
        if 'Transaction Type' in efs_data_df.columns:
            # create transaction type mapper
            transaction_type_mapper = {'Short': 'Seller', 'Long': 'Buyer'}
            # create seller mapper
            seller_mapper = {'Seller': 'On', 'Buyer': 'Off'}
            # create buyer mapper
            buyer_mapper = {'Seller': 'Off', 'Buyer': 'On'}
            # create buyer/seller column from transaction_type_mapper based on the value in transaction type column
            efs_data_df['Buyer/Seller'] = efs_data_df['Transaction Type'].map(transaction_type_mapper)
            # create seller column from seller_mapper based on the value in buyer/seller column
            efs_data_df['Seller'] = efs_data_df['Buyer/Seller'].map(seller_mapper)
            # create buyer column from buyer_mapper based on the value in buyer/seller buyer
            efs_data_df['Buyer'] = efs_data_df['Buyer/Seller'].map(buyer_mapper)

        # Create a column names mapper to change dataframe column names according to the field names of template pdf
        # Because most of the field names in template pdf return 'undefined' <= check with fillpdfs.get_form_fields(
        # template_pdf_file)
        mapper = {'Start Date': 'undefined_12',
                  'Expiry Date': 'undefined_13',
                  'Rounded strike': 'undefined_4',
                  'Total Quantity (lots)': 'undefined_3',
                  'Shipment Month': 'undefined_2',
                  'Notional': 'undefined_16',
                  'Commodity Code': 'Commodity Code  Contract Month'}

        # Rename some columns in dataframe according to the mapper
        efs_data_df = efs_data_df.rename(mapper=mapper, axis='columns')

        # Create a new column - undefined_14 with the same value as rounded strike to be filled in the EFS pdf
        # Because there are two fillable fields in pdf that need to be filled with price value ('undefined_4' &
        # 'undefined_14')
        efs_data_df['undefined_14'] = efs_data_df['undefined_4']

        return efs_data_df

    def create_efs_data_list_of_dict_from_df(self, efs_data_df: pd.DataFrame):
        """
        A function that create a list of EFS data dictionaries from EFS dataframe
        :param efs_data_df: The dataframe variable created from create_efs_data_list_of_dict_from_df function
        :return: A list of EFS data dictionaries
        """
        efs_data_list_of_dict = []
        # iterate over EFS data dataframe rows as (index, Series) pairs
        for _, row in efs_data_df.iterrows():
            # convert each dataframe row into dictionary and append the dictionary into a list
            efs_data_list_of_dict.append(dict(row))
            # print(dict(row))

        return efs_data_list_of_dict

    def fill_pdfs(self, efs_template_pdf: str, output_path: str, output_file_name: str, efs_data_list_of_dict: list):
        """
        A function that create individual PDF file based on efs data list of dictionaries
        :param efs_template_pdf: The pdf file name of EFS template (end with .pdf)
        :param output_path: The directory where you want to create the new PDF files
        :param output_file_name: The pdf file name of EFS that you want to create (*** without '.pdf'),
        the suffix of filename created will be increase by 1 with the first filename as 'xxxxx_1.pdf' and
        second filename as 'xxxxx_2.pdf'
        :param efs_data_list_of_dict: The name of list variable created from create_efs_data_list_of_dict_from_df
        function
        :return create individual EFS PDF file based on the number of data rows in EFS data excel file
        """
        # Create current date and current time variables
        current_date = datetime.datetime.now()
        # print(f'CURRENT DATETIME: {current_date}')
        current_date_str = current_date.strftime('%d/%m/%Y')
        current_time_str = current_date.strftime("%I.%M %p")

        # Create default field value dictionary to be filled in the EFS pdf
        default_dict = {'Date': current_date_str,
                        'Swaps': 'On',  # always tick
                        'Member Code': 'S111',
                        'Customer  Account Number': 'SRITRANG INTL / S72449G',
                        'Date  Time of EFP  EFS transaction': f'{current_date_str} at {current_time_str}',
                        'undefined': 'Sicom TSR',
                        'SWAPS': 'On',  # always tick
                        'Fixed Rate Payer  Floating Rate Receiver': 'SRITRANG',
                        'Fixed Rate Receiver  Floating Rate Payer': 'BANK'}

        # Check default dictionary and excel data dictionary have different keys
        # Merge default field value dictionary with efs data list of dictionaries & create new pdf file for each row
        for row in range(len(efs_data_list_of_dict)):
            if not set(default_dict.keys()).isdisjoint(efs_data_list_of_dict[row].keys()):
                raise Exception(
                    'Default dictionary and initial data column should have different keys to avoid any conflicts')
            # merge default field value dictionary with each dictionary in the efs data list
            final_value_dict = {**default_dict, **efs_data_list_of_dict[row]}
            # print(final_value_dict)
            # create pdf file for each data dictionary and update each pdf file with efs data dictionary
            fillpdfs.write_fillable_pdf(efs_template_pdf, f'{output_path}/{output_file_name}_{row + 1}.pdf',
                                        final_value_dict)
            # (optional) flatten the pdf file to make it uneditable
            # fillpdfs.flatten_pdf('new.pdf', 'newflat.pdf', as_image=False)
