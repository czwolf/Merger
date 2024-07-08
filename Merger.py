# https://pypdf2.readthedocs.io/en/1.28.4/
# PyPDF2 1.28.6.
# Pandas 2.0.3

import PyPDF2
import pandas as pd
import glob
import os

class Merger:
    def __init__(self, file_type, output_filename="merged", work_folder=os.getcwd(), output_folder=os.getcwd() ):
        self.file_type = file_type
        self.work_folder = work_folder
        self.output_folder = output_folder
        self.output_filename = output_filename
        self.supported_files = ["pdf", "PDF", "csv", "CSV", "xls", "XLS", "xlsx", "XLSX"]
        self.supported_files_print = ["pdf", "csv", "xls", "xlsx"]
        self.input_files = self.setInputFiles()

# Check file type. If supported then check if files exists in folder and prepare list of files for constructor
    def setInputFiles(self):
        if self.file_type in self.supported_files:
            input_files = glob.glob(self.work_folder + "\*."+self.file_type)
            # input_files = []
            # for file in os.listdir(self.work_folder):
            #     if file.endswith('.'+self.file_type):
            #         input_files.append(file)
            return input_files
        else:
            return f"Unsupported file type. You can merge only these file types {self.supported_files_print}"

# You can use function print for getting input files in list input_files
    def getInputFiles(self):
        if self.input_files:
            return self.input_files
        else:
            return f"Selected folder doesn't contain {self.file_type} files."

# Check if output folder exists. If yes, then return path to output folder. If no, then make folder.
    def checkOutputFolder(self):
        if os.path.exists(self.output_folder):
            return self.output_folder
        else:
            os.mkdir(self.output_folder)

# Merge method for merging all supported file types. It used next methods for every file type.
# You can use no mandatory parameters as duplicity_keep(bool)
# CSV and XLSX parametr: duplicity_keep=True(False)
    def merge(self, **kwargs):
        self.checkOutputFolder()
        if not self.input_files:
            return f"Selected folder doesn't contain {self.file_type} files."
        elif self.input_files == f"Unsupported file type. You can merge only these file types {self.supported_files_print}":
            return f"Unsupported file type. You can merge only these file types {self.supported_files_print}"
        else:
            if self.file_type == "pdf" or self.file_type == "PDF":
                return self.pdfMerger()

            elif self.file_type == "csv" or self.file_type == "CSV":
                if kwargs:
                    for key, value in kwargs.items():
                        if key == "duplicity_keep":
                            return self.csvMerger(duplicity_keep=value)
                else:
                    return self.csvMerger()

            elif self.file_type == "xlsx" or self.file_type == "xls" or self.file_type == "XLSX" or self.file_type == "XLS":
                if kwargs:
                    for key, value in kwargs.items():
                        if key == "duplicity_keep":
                            return self.xlsxMerger(duplicity_keep=value)
                else:
                    return self.xlsxMerger()

# Method for merging PDF files
    def pdfMerger(self):
        output_filename = self.output_filename + "." + self.file_type
        if output_filename in os.listdir(self.work_folder):
            self.input_files.remove(output_filename)

        merger = PyPDF2.PdfFileMerger()
        for file in self.input_files:
            merger.append(file)

        merger.write(self.output_folder + "/" + output_filename)
        merger.close()

        return f"Output file saved to: {self.output_folder}"

# Method for merging CSV files.
# You can remove duplicities through the parameter in merge method
    def csvMerger(self, sep=";", encode="utf8", duplicity_keep=True):
        duplicity_keep = duplicity_keep
        sep = sep
        encode = encode
        df_append = pd.DataFrame()
        output_filename = self.output_filename + "." + self.file_type
        output = self.output_folder + "/" + output_filename

        if output_filename in os.listdir(self.work_folder):
            self.input_files.remove(output_filename)

        for csv in self.input_files:
            df = pd.read_csv(csv, sep=sep, encoding=encode)
            df_append = df_append._append(df, ignore_index=True)

        if duplicity_keep:
            df_append.to_csv(output, index=False, sep=';')
            return f"Output file saved to: {self.output_folder}"
        else:
            df_append = df_append.drop_duplicates()
            df_append.to_csv(output, index=False, sep=';')
            return f"Output file saved to: {self.output_folder}"

# Method for merging XLSX and XLS files.
# You can remove duplicities through the parameter in merge method
    def xlsxMerger(self, duplicity_keep=True):
        duplicity_keep = duplicity_keep
        output_filename = self.output_filename + "." + self.file_type
        output = self.output_folder + "/" + output_filename
        all_dfs = pd.DataFrame()

        if output_filename in os.listdir(self.work_folder):
            self.input_files.remove(output_filename)

        for xlsx in self.input_files:
            df = pd.read_excel(xlsx, engine="openpyxl")
            all_dfs = pd.concat([all_dfs, df], ignore_index=True, sort=False)

        if duplicity_keep:
            all_dfs.to_excel(output, index=None, engine="openpyxl")
            return f"Output file saved to: {self.output_folder}"
        else:
            all_dfs = all_dfs.drop_duplicates()
            all_dfs.to_excel(output, index=None, engine="openpyxl")
            return f"Output file saved to: {self.output_folder}"