import os, time, ctypes, re
from datetime import datetime, timedelta
from pathlib import Path
import win32com.client as win32


def Convert_xls_xlsb(files_list: list, location):
    """convert file"""
    xlExcel12 = 50
    all_details_files = []
    xl = win32.Dispatch('Excel.Application')
    # xl.Visible = False
    for file in files_list:
        filename = re.search(r"Unet (.*)\.xls\.xls", file.name).group(1)
        newfilename = str(Path.joinpath(location, str(filename) + ".xlsb"))
        wb = xl.Workbooks.Open(file)
        set_sensitiviy_label(wb)
        wb.SaveAs(newfilename, FileFormat=xlExcel12)
        wb.Close()
    

# setting sensitivity label
def set_sensitiviy_label(wbook):
    """set sensitivity label"""
    label = wbook.SensitivityLabel.CreateLabelInfo()
    label.AssignmentMethod = 1  #MsoAssignmentMethod.PRIVILEGED
    # label id and site id can be find in excel through vba sub routine
    label.LabelId = "a8a73c85-e524-44a6-bd58-7df7ef87be8f"
    label.SiteId = "6c15903a-880e-4e17-818a-6cb4f7935615"
    wbook.SensitivityLabel.SetLabel(label, label)

# list of file path that need to be converted to xlsb extension
files_list = ['file1', 'file2', 'file3']

# destination for where converted file to be saved
documents_directory = r'\path\destination'

# call the function
Convert_xls_xlsb(files_list, documents_directory)