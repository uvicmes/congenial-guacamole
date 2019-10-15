'''
This is an initial approach to the automation of CTD data processing via excel/csv
data parsing and manipulation. This developing software will enable Ocean Labs
researchers to automatically inject metadata from an excel file to each generated
ctd xml casts.

Author: Andy Breton (Bachelor of Computer Science, University of the Virgin Islands)
'''
import xlrd
import os
import lxml.etree as ET
from datetime import time
import datetime
import subprocess
import shutil
from tkinter import ttk
from tkinter import filedialog
import tkinter as tk
import html
import time as tm
#from tkinter import Image
#from tkinter.ttk import Progressbar
import sys
import pandas as pd
import numpy as np
import csv
#import matplotlib.pyplot as plt
#import matplotlib.ticker as mtick
#import math
from code import InteractiveConsole

Error = (
        "Error 0: The Headers in the Excel file cannot be found. Make sure that the headers are\nproperley labeled \
as seen in the example datasheet.",

        "Error 1: There are more casts files on the raw data folder than entries on the excel table\nentries. It is \
most likely that you may have an (accidental) extra cast on the (Raw CTD) folder that\nyou may have to delete and try \
again.",

        "Error 2: There are more entries in the excel file than available cast files in the folder,\nplease check your \
excel table entries for discrepancies and try again and\or make sure that you have\nselected the right (Raw CTD) folder \
that contains.",

        "Error 3: The format of the excel file is not in the correct format, please format of the\nexcel file as shown \
in the example ctd format sheet located in the server."

        "Error 4: The file does not exist. Make sure that you've selected the correct file.")

def update_Status(text):
    status_message.set(text)

def find_header(totalrows, header_name, expected_column, excel_file_path):
    workbook = xlrd.open_workbook(excel_file_path)
    sheet = workbook.sheet_by_index(0)
    index = 0
    for j in range(totalrows):
        if (sheet.cell_value(j, expected_column) == header_name):
            print("Cast header found in row " + str(j+1))
            index = j+1
            break
    return index

def excel_n_files_Checker(excel_file_path, ctd_raw_folder_path):

    #def date_to_Julian():
    source_raw_xmlhex_files = sorted(os.listdir(ctd_raw_folder_path))
    workbook = xlrd.open_workbook(excel_file_path)
    sheet = workbook.sheet_by_index(0)
    total_rows = sheet.nrows - 1
    print("inside checker")
    j = find_header(total_rows, "Cast #", 0, excel_file_path)
    if j == 0:
        return 0

    else:
        #this represents the rows that contains the variables values after the headers
        rows_to_read = sheet.nrows-j
        raw_xml_files_count = len(source_raw_xmlhex_files)
        print (str(rows_to_read) + " rows to read in excel.")
        print (str(raw_xml_files_count) + " files in folder. \n")

        # Error Handling statements
        if raw_xml_files_count > rows_to_read:
            return 1

        elif rows_to_read > raw_xml_files_count:
            return 2

        elif rows_to_read == raw_xml_files_count:
            print("proceed")
            return "proceed"

def xml_Metadata_SBE25Plus(project_name, ctd_model, ctd_serial, year, julian_day, excel_file_path, ctd_raw_folder_path, output_folder_path):

    #def date_to_Julian():
    source_raw_xml_files = sorted(os.listdir(ctd_raw_folder_path))
    workbook = xlrd.open_workbook(excel_file_path)
    sheet = workbook.sheet_by_index(0)
    total_rows = sheet.nrows - 1
    total_columns = sheet.ncols - 1

    j = find_header(total_rows, "Cast #", 0, excel_file_path)
    notes_header_pos = find_header(total_rows, "Notes", 13, excel_file_path)

    #this represents the rows that contains the variables values after the headers
    rows_to_read = sheet.nrows-j
    raw_xml_files_count = len(source_raw_xml_files)
    print (str(rows_to_read) + " rows to read in excel.")
    print (str(raw_xml_files_count) + " files in folder. \n")

    # Error Handling statements
    if raw_xml_files_count == rows_to_read:

        for i in range(rows_to_read):

            #variables
            cast_number = sheet.cell_value(j, 0)
            site_name = sheet.cell_value(j, 1)
            site_code = sheet.cell_value(j, 2)
            station_code = sheet.cell_value(j, 3)
            secchi_disk = sheet.cell_value(j, 4)
            latitude_in = sheet.cell_value(j, 5)
            longitude_in = sheet.cell_value(j, 6)
            raw_timein_hours = sheet.cell_value(j, 7)
            propertimein = xlrd.xldate_as_tuple(raw_timein_hours, workbook.datemode)
            time_in = time(*propertimein[3:])
            depth = sheet.cell_value(j, 8)
            line_out = sheet.cell_value(j, 9)
            raw_timeout_hours = sheet.cell_value(j, 10)
            propertimeout = xlrd.xldate_as_tuple(raw_timeout_hours, workbook.datemode)
            time_out = time(*propertimeout[3:])
            latitude_out = sheet.cell_value(j, 11)
            longitude_out = sheet.cell_value(j, 12)

            if total_columns == 13:

                notes = sheet.cell_value(j, 13)

                CTD_format_toString = "\n<![CDATA["+"\n"+"** Location Name "+str(site_name)+"\n"\
                +"** Station Code "+str(station_code).zfill(3)+"\n"+"** Secchi Disk (ft) "+str(secchi_disk)+"\n"\
                "** Latitude In "+str(latitude_in)+"\n"+"** Longitude In "+\
                str(longitude_in)+"\n"+"** Time In "+str(time_in)+"\n"+"** Depth (ft) "+str(depth)+\
                "\n"+"** Line Out (ft) "+str(line_out)+"\n"+"** Time Out "+str(time_out)+\
                "\n"+"** Latitude Out "+str(latitude_out)+"\n"+"** Longitude Out "+\
                str(longitude_out)+"\n"+"** Notes "+str(notes)+"\n"+"]]>\n"

                print (CTD_format_toString)
                xml_file_path = ctd_raw_folder_path+source_raw_xml_files[i]
                xml_tree = ET.parse(xml_file_path)
                xml_root = xml_tree.getroot()
                xml_root.text = CTD_format_toString
                print (xml_file_path + "\n")

                new_xml_file_path = output_folder_path+str(year)+'_'+str(julian_day).zfill(3)+'_'+str(int(cast_number)).zfill(3)\
                               +'_'+str(project_name)+'_'+str(site_code)+'_'+str(ctd_model).upper()+str(int(ctd_serial)).zfill(4)+'.xml'
                f = open(new_xml_file_path, "w")
                f.write(html.unescape(ET.tostring(xml_root, encoding="unicode")))
                f.close()

                j = j+1
                i = i+1

            if total_columns == 18:

                par1 = sheet.cell_value(j, 13)
                D1 = sheet.cell_value(j, 14)
                par2 = sheet.cell_value(j, 15)
                D2 = sheet.cell_value(j, 16)
                par3 = sheet.cell_value(j, 17)
                D3 = sheet.cell_value(j, 18)

                CTD_format_toString = "\n<![CDATA["+"\n"+"** Location Name "+str(site_name)+"\n"\
                +"** Station Code "+str(station_code).zfill(3)+"\n"+"** Secchi Disk (ft) "+str(secchi_disk)+"\n"\
                "** Latitude In "+str(latitude_in)+"\n"+"** Longitude In "+\
                str(longitude_in)+"\n"+"** Time In "+str(time_in)+"\n"+"** Depth (ft) "+str(depth)+\
                "\n"+"** Line Out (ft) "+str(line_out)+"\n"+"** Time Out "+str(time_out)+\
                "\n"+"** Latitude Out "+str(latitude_out)+"\n"+"** Longitude Out "+\
                str(longitude_out)+"\n"+"** PAR (umol s-1 m-2 per uA) | Depth (ft) "+str(par1)+" | "+str(D1)+", "+str(par2)+" | "+str(D2)+", "+str(par3)+" | "+str(D3)+"\n]]>\n"

                print (CTD_format_toString)

                xml_file_path = ctd_raw_folder_path+source_raw_xml_files[i]
                print (xml_file_path + "\n")
                xml_tree = ET.parse(xml_file_path)
                xml_root = xml_tree.getroot()
                xml_root.text = CTD_format_toString

                new_xml_file_path = output_folder_path+str(year)+'_'+str(julian_day).zfill(3)+'_'+str(int(cast_number)).zfill(3)\
                               +'_'+str(project_name)+'_'+str(site_code)+'_'+str(ctd_model).upper()+str(int(ctd_serial)).zfill(4)+'.xml'

                f = open(new_xml_file_path, "w")
                f.write((html.unescape(ET.tostring(xml_root, encoding="unicode"))))
                f.close()
                j = j+1
                i = i+1

def hex_Metadata_SBE25(project_name, ctd_model, ctd_serial, year, julian_day, excel_file_path, ctd_raw_folder_path, output_folder_path):

    #def date_to_Julian():
    source_raw_hex_files = sorted(os.listdir(ctd_raw_folder_path))
    workbook = xlrd.open_workbook(excel_file_path)
    sheet = workbook.sheet_by_index(0)
    total_rows = sheet.nrows - 1
    total_columns = sheet.ncols - 1

    j = find_header(total_rows, "Cast #", 0, excel_file_path)
    notes_header_pos = find_header(total_rows, "Notes", 13, excel_file_path)

    #this represents the rows that contains the variables values after the headers
    rows_to_read = sheet.nrows-j
    raw_xml_files_count = len(source_raw_hex_files)
    print (str(rows_to_read) + " rows to read in excel.")
    print (str(raw_xml_files_count) + " files in folder. \n")

    # Error Handling statements
    if raw_xml_files_count == rows_to_read:

        for i in range(rows_to_read):

            #variables
            cast_number = sheet.cell_value(j, 0)
            site_name = sheet.cell_value(j, 1)
            site_code = sheet.cell_value(j, 2)
            station_code = sheet.cell_value(j, 3)
            secchi_disk = sheet.cell_value(j, 4)
            latitude_in = sheet.cell_value(j, 5)
            longitude_in = sheet.cell_value(j, 6)
            raw_timein_hours = sheet.cell_value(j, 7)
            propertimein = xlrd.xldate_as_tuple(raw_timein_hours, workbook.datemode)
            time_in = time(*propertimein[3:])
            depth = sheet.cell_value(j, 8)
            line_out = sheet.cell_value(j, 9)
            raw_timeout_hours = sheet.cell_value(j, 10)
            propertimeout = xlrd.xldate_as_tuple(raw_timeout_hours, workbook.datemode)
            time_out = time(*propertimeout[3:])
            latitude_out = sheet.cell_value(j, 11)
            longitude_out = sheet.cell_value(j, 12)

            print("here we go")

            notes = sheet.cell_value(j, 13)

            CTD_format_toString = "** Location Name "+str(site_name)+"\n"\
            +"** Station Code "+str(station_code).zfill(3)+"\n"+"** Secchi Disk (ft) "+str(secchi_disk)+"\n"\
            "** Latitude In "+str(latitude_in)+"\n"+"** Longitude In "+\
            str(longitude_in)+"\n"+"** Time In "+str(time_in)+"\n"+"** Depth (ft) "+str(depth)+\
            "\n"+"** Line Out (ft) "+str(line_out)+"\n"+"** Time Out "+str(time_out)+\
            "\n"+"** Latitude Out "+str(latitude_out)+"\n"+"** Longitude Out "+\
            str(longitude_out)+"\n"+"** Notes "+str(notes)+"\n"

            print (CTD_format_toString)
            rawhex_file_path = ctd_raw_folder_path+source_raw_hex_files[i]
            newhex_file_path = output_folder_path+str(year)+'_'+str(julian_day).zfill(3)+'_'+str(int(cast_number)).zfill(3)\
                               +' '+str(project_name)+'_'+str(site_code)+'_'+str(ctd_model).upper()+str(int(ctd_serial)).zfill(4)+'.hex'


            with open(rawhex_file_path) as f_raw, open(newhex_file_path, "w") as f_new:
                for line in f_raw:
                    f_new.write(line)
                    if 'System UpLoad Time' in line:
                        f_new.write(str(CTD_format_toString))

            f_raw.close()
            f_new.close()

            j = j+1
            i = i+1

def exit_program():
    print ("exiting program...")
    root.destroy()
    sys.exit()

def fetch_File(filetype, target):
    filename = filedialog.askopenfilename(title='Choose a file', filetypes=[(filetype)])
    target.set(filename)

def fetch_Directory(target):
    filename = filedialog.askdirectory(parent=root, title='Choose a folder')
    target.set(filename)

def copytree(src, dst, symlinks=False, ignore=None):
    if not os.path.exists(dst):
        os.makedirs(dst)
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            copytree(s, d, symlinks, ignore)
        else:
            if not os.path.exists(d) or os.stat(s).st_mtime - os.stat(d).st_mtime > 1:
                shutil.copy2(s, d)

def ascii_to_CSVRaw(in_src, out_src):
    fl = os.listdir(in_src)
    for f in fl:
        if f[-3:] == "asc":
            # converts asc file to csv file
            intext = csv.reader(open(in_src+f, "r"), delimiter='\t')
            outcsv = csv.writer(open(out_src + f[:-3] + "csv", "w", encoding='utf-8', newline=''))
            outcsv.writerows(intext)

def CSVRaw_CSVClean(in_src, out_src, timecut, surface_Limit):

    fl2 = os.listdir(in_src)
    for f in fl2:
        # import csv file as a dataframe in pandas
        x = pd.read_csv(in_src + "\\" + f)
        # replaces all bad data -9.99e-29 with nan or null
        y = x.replace(to_replace=-9.99e-29, value=np.nan)
        index = None
        z = None
        # Drops the Flag Column
        for col in y:
            if 'Flag' in col:
                z = y.drop('Flag', 1)
            else:
                z = y
        # Removes all rows that contain NaN values - Helpful may be removed in later versions
        y = z.dropna()
        # Selects the depth column out and finds the index value of the max depth
        # print("Here 7 here was crashing")
        d = y.ix[:, 'DepSM']
        c = d.idxmax()
        # extracts the downcast using the above cut off for the max depth
        cc = y.ix[:c, :]
        # creation of final cast (fc)
        fc = cc[cc.TimeS > timecut]
        fc2 = fc[cc.DepSM > surface_Limit]

        # export clean dataframe to new csv
        fc2.to_csv(out_src + "\\" + f, na_rep='NaN', index=False, encoding='utf-8')


def Grapth_Casts(ascii_folder, rawcsv_folder, cleancsv_folder, graphout_folder, soaktime, surface_Limit):
    ascii_to_CSVRaw(ascii_folder, rawcsv_folder)
    CSVRaw_CSVClean(rawcsv_folder, cleancsv_folder, soaktime, surface_Limit)

'''

    fl = os.listdir(cleancsv_folder)

    for f in fl:
        print(f, "namestring")
        #import csv file as a dataframe in pandas
        data = pd.read_csv(cleancsv_folder + "\\" + f)
        #figure created with 9 subplots
        fig = plt.figure(1)
        p1 = fig.add_subplot(3,3,1)
        p2 = fig.add_subplot(3,3,2)
        p3 = fig.add_subplot(3,3,3)
        p4 = fig.add_subplot(3,3,4)
        p5 = fig.add_subplot(3,3,5)
        p6 = fig.add_subplot(3,3,6)
        p7 = fig.add_subplot(3,3,7)
        p8 = fig.add_subplot(3,3,8)
        p9 = fig.add_subplot(3,3,9)

        #Depth column (y-axis)
        d = data.ix[:,"DepSM"]
        dep_mi=min(d)
        dep_ma=max(d)
        dep_me1=(dep_mi+dep_ma)/2
        dep_me0=(dep_mi+dep_me1)/2
        dep_me2=(dep_ma+dep_me1)/2

        def findCSVHeader(variable1, variable2):
            for col in data:
                if variable1 in col:
                    return variable1
                elif variable2 in data:
                    return variable2
            return 0

        def CSV_grapher(p, columnName1, columnName2, colorS, mainlabelS, xlabelS, formatStringS):

            variable = findCSVHeader(columnName1, columnName2)

            if variable == 0:
                print("graph", mainlabelS,"not found")
                p.set_title(mainlabelS + " N/A", fontsize=10)
                p.set_xlabel(xlabelS, fontsize=8)
                p.set_xticklabels("0", fontsize=8)
                p.set_yticklabels("0", fontsize=8)

            else:
                print("Graphing", variable)
                t = data.ix[:, variable]
                mi = min(t)
                ma = max(t)
                me1 = (mi + ma) / 2
                me0 = (mi + me1) / 2
                me2 = (ma + me1) / 2

                p.plot(t, d, color=colorS)
                p.set_title(mainlabelS, fontsize=10)
                p.set_xlabel(xlabelS, fontsize=8)
                if mainlabelS == "Temperature":
                    p.set_ylabel("Depth (m)", fontsize=8)
                if mainlabelS == "Chlorophyll":
                    p.set_ylabel("Depth (m)", fontsize=8)
                if mainlabelS == "Oxygen Saturation":
                    p.set_ylabel("Depth (m)", fontsize=8)
                if not variable == "Density00":
                    p.set_xticks([mi, me0, me1, me2, ma])
                    p.set_xticklabels([mi, me0, me1, me2, ma], fontsize=8)
                else:
                    p.set_xticks([mi,me1,ma])
                    p.set_xticklabels([mi,me1,ma], fontsize=8)
                p.set_yticks([dep_mi, dep_me0, dep_me1, dep_me2, dep_ma])
                p.set_yticklabels([dep_mi, dep_me0, dep_me1, dep_me2, dep_ma], fontsize=8)
                p.invert_yaxis()
                p.xaxis.set_major_formatter(mtick.FormatStrFormatter(formatStringS))
                p.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.1f'))

        #Plot 1 - Temperature
        CSV_grapher(p1, "T090C", "placeholder", "r", "Temperature", "Temp ($^\circ$C)", "%.2f")
        CSV_grapher(p2, "Sal00", "placeholder", "c", "Salinity", "Salinity (PSU)", "%.2f")
        CSV_grapher(p3, "Density00", "placeholder", "b", "Density", "Density (kg/m$^3\!$)", "%.2f")
        CSV_grapher(p4, "FlECO-AFL", "placeholder", "g", "Chlorophyll", "Chlorophyll (mg/m$^3\!$)", "%.2f")
        CSV_grapher(p5, "TurbWETntu0", "Upoly0", "k", "Turbidity", "Turbidity (NTU)", "%.2f")
        CSV_grapher(p6, "C0S/m", "placeholder", "m", "Conductivity", "Conductivity (s/m)", "%.2f")
        CSV_grapher(p7, "OxsatMg/L", "Sbeox0Mg/L", "b", "Oxygen Saturation", "Oxygen Saturation (mg/L)", "%.2f")
        CSV_grapher(p8, "pH", "Ph", "c", "pH", "pH", "%.2f")
        CSV_grapher(p9, "Par", "PAR", "y", "PAR", "PAR (µmol s$^-\!$$^1\!$ m$^-\!$$^2\!$ per µA)", "%.2f")

        #Adjusting the white space between subplots
        fig.subplots_adjust(hspace=.3,wspace=.2)
        fig.set_size_inches(8,11)
        fsplit = f.split("_")
        ctdsplit = fsplit[5]
        fig.suptitle("Project: " + fsplit[3]+", CTD: "+ctdsplit[:-8]+", Site: "+fsplit[4]+", Julian Date: "+fsplit[0]+" "+
                     fsplit[1]+", Cast: "+fsplit[2], fontsize=12)

        #Exports figure to folder and saves as pdf - this can be changed
        plt.savefig(graphout_folder + "\\" + f[:-3] + "pdf",dpi=400)
        plt.close(fig)
    '''
def start_button():

    global pre_processed_input
    try:
        print("on_button")
        start = tm.time()
        excel_file_input = excelFilePath.get()
        config_file = ctd_config_path.get()
        ctd_raw_folder_path = ctdRawFolderPath.get() + "/"
        output_folder_path = outputFolderPath.get() + "/"

        workbook = xlrd.open_workbook(excel_file_input)
        sheet = workbook.sheet_by_index(0)
        project_name = sheet.cell_value(0, 0)
        ctd_caster = sheet.cell_value(0, 2)
        ctd_serial = sheet.cell_value(0, 5)
        ctd_castdate = sheet.cell_value(0, 7)
        ctd_castdateGregorian = xlrd.xldate_as_tuple(ctd_castdate, datemode=0)

        print(project_name)
        print(ctd_caster)
        print(ctd_serial)
        year = int(ctd_castdateGregorian[0])
        month = int(ctd_castdateGregorian[1])
        day = int(ctd_castdateGregorian[2])
        julian_day = (datetime.datetime(int(year), int(month), int(day)).timetuple().tm_yday)

        ctd_model = ctdUsedEntry.get()
        soak = int(soaktime.get())
        surface_Limit = scale_depth.get()


        # all folders paths
        root_dir = os.getcwd() + "\\"
        main_dir = output_folder_path + project_name+" "+str(year)+"-"+str(month).zfill(2)+"-"+str(day).zfill(2)+" "+str(ctd_model)\
                   +str(int(ctd_serial)).zfill(4)+"\\"
        analyzed_dir = main_dir + "analyzed_csv\\"
        asc_in_dir = main_dir + "asc_in\\"
        ascii_dir = main_dir + "ascii\\"
        binary_dir = main_dir + "binary\\"
        clean_csv_dir = main_dir + "clean_csv\\"
        config_dir = main_dir + "config\\"
        data_sheet_dir = main_dir + "data_sheet\\"
        extracted_raw_xml_hex_dir = main_dir + "extracted_raw_xml_hex\\"
        graphs_dir = main_dir + "graphs\\"
        pre_processed_dir = main_dir + "pre_processed_xml_hex\\"
        processed_dir = main_dir + "processed\\"
        raw_csv_dir = main_dir + "raw_csv\\"
        xmlhex_to_txtfile = main_dir + "xmlhex_txt"
        SBE_Dependencies_dir = root_dir + "SBE_Dependencies\\"

        def process_data_derive(process, input_files, output_folder):
            print("Processing", process + "...")
            SBE_Executable = process + "W"  # "DatCnvW.exe"
            configuration = config_file  # "ctd_1107_2016_07_pH_xml_BOT.xmlcon"
            PSA_file = SBE_Dependencies_dir + process + ".psa"  # "DatCnv.psa"
            toString = str('"' + SBE_Executable + '"') + " /c" + str('"' + configuration + '"') + " /p" + str(
                '"' + PSA_file + '"') + " /i" + str(
                '"' + input_files + '"') + " /o" + str('"' + output_folder + '"') + " /s /m"
            subprocess.call(toString)

        def process_fawcla(process, input_files, output_folder):
            print("Processing", process + "...")
            SBE_Executable = process + "W"  # "DatCnvW.exe"
            PSA_file = SBE_Dependencies_dir + process + ".psa"  # "DatCnv.psa"
            toString = str('"' + SBE_Executable + '"') + " /p" + str('"' + PSA_file + '"') + " /i" + str(
                '"' + input_files + '"') + " /o" + str('"' +
                                                       output_folder + '"') + " /s /m"
            subprocess.call(toString)

        if project_name == "":
            raise ValueError("Warning: Please fill out the project name blank textbox.")
        elif ctd_serial == "":
            raise ValueError("Warning: Please fill out the CTD Used blank textbox.")
        elif excel_file_input == "":
            raise ValueError("Warning: Browse and Select excel file to process.")
        elif os.path.exists(excel_file_input) is False:
            raise ValueError("Error: The excel file doesn't exists, select the correct file.")
        elif config_file == "":
            raise ValueError("Warning: Browse and Select CTD config file to process the data.")
        elif os.path.exists(config_file) is False:
            raise ValueError("Error: The CTD configuration file doesn't exist. Select the correct file.")
        elif ctd_raw_folder_path == "/":
            raise ValueError("Warning: Browse and Select a folder that contains the raw xml-hex files.")
        elif os.path.exists(ctd_raw_folder_path) is False:
            raise ValueError("Error: The path of the (Raw CTD Folder Path) doesn't exists")
        #elif not any(fname.endswith('.xml') for fname in os.listdir(ctd_raw_folder_path)):
            #raise ValueError("The folder that you've selected for (Raw CTD Folder Path) doesn't contain any .xml input\n"
                             #"files.")
        elif output_folder_path == "/":
            raise ValueError("Warning: Browse and Select a folder to put the results in.")
        elif os.path.exists(output_folder_path) is False:
            raise ValueError("Error: The path of the (Output Folder Path) doesn't exists")
        elif year == "":
            raise ValueError("Warning: Please insert the year value (YYYY) of the cast.")
        elif month == "":
            raise ValueError("Warning: Please insert the month value (MM) of the cast.")
        elif day == "":
            raise ValueError("Warning: Please insert the day value (DD) of the cast.")
        elif soak == "":
            raise ValueError("Warning: Please insert the amount of soak time of cast.")

        if excel_n_files_Checker(excel_file_input, ctd_raw_folder_path) == "proceed":

            print("Date to Julian day: " + str(julian_day))
            print("Starting metadata injection...")
            print("Creating directories...")
            if not os.path.exists(main_dir):
                os.makedirs(main_dir)
                os.makedirs(analyzed_dir)
                os.makedirs(asc_in_dir)
                os.makedirs(ascii_dir)
                os.makedirs(binary_dir)
                os.makedirs(clean_csv_dir)
                os.makedirs(config_dir)
                os.makedirs(data_sheet_dir)
                os.makedirs(extracted_raw_xml_hex_dir)
                os.makedirs(graphs_dir)
                os.makedirs(pre_processed_dir)
                os.makedirs(processed_dir)
                os.makedirs(raw_csv_dir)
                os.makedirs(xmlhex_to_txtfile)

            # all input files

            binary_input = binary_dir + "*.cnv"
            processed_input = processed_dir + "*.cnv"

            print("Injecting metedata to cast files...")

            if ctd_model == "SBE025":
                print("SBE025 model chosen")
                pre_processed_input = pre_processed_dir + "*.hex"
                hex_Metadata_SBE25(project_name, ctd_model, ctd_serial, year, julian_day, excel_file_input, ctd_raw_folder_path, pre_processed_dir)

            if ctd_model == "SBE025Plus":
                print("SBE025Plus model chosen")
                pre_processed_input = pre_processed_dir + "*.xml"
                xml_Metadata_SBE25Plus(project_name, ctd_model, ctd_serial, year, julian_day, excel_file_input, ctd_raw_folder_path, pre_processed_dir)

            #pre_processed_input = pre_processed_dir + "*.xml"
            print("Starting processing...")
            process_data_derive("DatCnv", pre_processed_input, binary_dir)
            process_fawcla("AlignCTD", binary_input, processed_dir)
            process_fawcla("WildEdit", processed_input, processed_dir)
            process_fawcla("Filter", processed_input, processed_dir)
            process_fawcla("CellTM", processed_input, processed_dir)
            process_data_derive("Derive", processed_input, processed_dir)
            process_fawcla("LoopEdit", processed_input, processed_dir)
            process_fawcla("ASCII_Out", processed_input, ascii_dir)

            print("Transfering input files used...")
            shutil.copy(config_file, config_dir)
            shutil.copy(excel_file_input, data_sheet_dir)
            copytree(ctd_raw_folder_path, extracted_raw_xml_hex_dir)

            print("Metadata injection finished!!!")

            if excel_n_files_Checker(excel_file_input, ctd_raw_folder_path) == 0:
                print("Error 0")
                raise ValueError(Error[0])
            elif excel_n_files_Checker(excel_file_input, ctd_raw_folder_path) == 1:
                print("Error 1")
                raise ValueError(Error[1])
            elif excel_n_files_Checker(excel_file_input, ctd_raw_folder_path) == 2:
                print("Error 3")
                raise ValueError(Error[2])

            print("Graphing Clean CSV Data")
            #Grapth_Casts(ascii_dir, raw_csv_dir, clean_csv_dir, graphs_dir, soak, surface_Limit)

            update_Status("Data Processed Successfully!!!")
            end = tm.time()
            print(end - start)

    except ValueError as e:
        update_Status(e)

if __name__ == "__main__":
    root_dir = os.getcwd()
    root = tk.Tk()
    root.title("CTD Metadata Injector (BETA)")
    icon_path = root_dir+"\\SBE_Dependencies\\images\\ICON_UVI_logo.ico"
    root.iconbitmap(default=icon_path)
    root.resizable(0,0)

    excel = 'Excel files', '.xlsx'
    config_file = 'XMLCON files', '.xmlcon'

    default_soak = 45
    default_surface_Limit = 1
    mainframe = tk.Frame(root,width=500,height=500)
    mainframe.grid(sticky="nw")

    label_ctd_used = ttk.Label(mainframe, text="CTD Model Used: ")
    label_excel = ttk.Label(mainframe, text="Excel Sheet File Path: ")
    label_configuration = ttk.Label(mainframe, text="CTD Config File: ")
    label_CTD_raw = ttk.Label(mainframe, text="Raw CTD Folder Path: ")
    label_Output = ttk.Label(mainframe, text="Output Folder Path: ")
    label_soaktime = ttk.Label(mainframe, text="Soak time (in seconds): ")
    label_surface_Limit = ttk.Label(mainframe, text="AVG depth of bottom cast (in meters): ")
    label_depth = ttk.Label(mainframe, text="Depth start graph(m):")

    OptionList_ctdModel = ["SBE025", "SBE025Plus"]
    ctdUsedEntry = tk.StringVar(None)
    ctdUsedEntry.set("SBE025Plus")
    ctddroplist = tk.OptionMenu(mainframe, ctdUsedEntry, *OptionList_ctdModel)
    ctddroplist.update()
    ctddroplist.focus_set()

    excel_path = tk.StringVar(None)
    excelFilePath = ttk.Entry(mainframe, width=44, textvariable=excel_path)
    excelFilePath.update()
    excelFilePath.focus_set()
    button_browse_Excel_file = ttk.Button(mainframe, text='Select...', command=lambda: fetch_File(excel, excel_path))

    ctd_config_file = tk.StringVar(None)
    ctd_config_path = ttk.Entry(mainframe, width=44, textvariable=ctd_config_file)
    ctd_config_path.update()
    ctd_config_path.focus_set()
    button_browse_ctd_config_file = ttk.Button(mainframe, text='Select...', command=lambda: fetch_File(config_file, ctd_config_file))

    raw_folderpath = tk.StringVar(None)
    ctdRawFolderPath = ttk.Entry(mainframe, width=44, textvariable=raw_folderpath)
    ctdRawFolderPath.update()
    ctdRawFolderPath.focus_set()
    button_select_raw_folder = ttk.Button(mainframe, text='Browse...', command=lambda: fetch_Directory(raw_folderpath))

    output_folderpath = tk.StringVar(None)
    outputFolderPath = ttk.Entry(mainframe, width=44, textvariable=output_folderpath)
    outputFolderPath.update()
    outputFolderPath.focus_set()
    button_select_output_folder = ttk.Button(mainframe, text='Browse...', command=lambda: fetch_Directory(output_folderpath))

    optionListsoak = [0, 45, 60, 120, 180, 240]
    soaktime = tk.StringVar(None)
    soaktime.set(120)
    entry_soaktime = tk.OptionMenu(mainframe, soaktime, *optionListsoak)
    entry_soaktime.update()
    entry_soaktime.focus_set()

    is_checked = tk.IntVar(None)
    is_checked.set(0)
    CheckMN = ttk.Checkbutton(mainframe, text="Project Mare Nostrum", onvalue=1, offvalue=0, variable=is_checked)

    scale_depth = tk.Scale(mainframe, orient= tk.HORIZONTAL, from_=0, to=5, resolution=0.1)
    scale_depth.set(0)
    scale_depth.update()
    scale_depth.focus_set()

    button_start = ttk.Button(mainframe, text="Start", command=start_button)
    button_quit = ttk.Button(mainframe, text="Quit", command=exit_program)

    status_message = tk.StringVar(None)
    status = ttk.Label(mainframe, textvariable=status_message)

    #Gui widget placement on template

    label_ctd_used.place(relx=0.02, rely=0.08)
    ctddroplist.place(relx=0.26, rely=0.06)
    label_excel.place(relx=0.02, rely=0.14)
    excelFilePath.place(relx=0.26, rely=0.14)
    button_browse_Excel_file.place(relx=0.82, rely=0.135)
    label_configuration.place(relx=0.02, rely=0.20)
    ctd_config_path.place(relx=0.26, rely=0.20)
    button_browse_ctd_config_file.place(relx=0.82, rely=0.195)
    label_CTD_raw.place(relx=0.02, rely=0.26)
    ctdRawFolderPath.place(relx=0.26, rely=0.26)
    button_select_raw_folder.place(relx=0.82, rely=0.255)
    label_Output.place(relx=0.02, rely=0.32)
    outputFolderPath.place(relx=0.26, rely=0.32)
    button_select_output_folder.place(relx=0.82, rely=0.315)
    label_soaktime.place(relx=0.02, rely=0.38)
    entry_soaktime.place(relx=0.27, rely=0.37)
    CheckMN.place(relx=0.02, rely=0.44)
    label_depth.place(relx=0.40, rely=0.38)
    scale_depth.place(relx=0.64, rely=0.37)
    button_start.place(relx=0.25, rely=0.93)
    button_quit.place(relx=0.6, rely=0.93)
    status.place(relx=0.02, rely=0.5)

    root.mainloop()
print("++++++++++++++++Program Exited++++++++++++++++++++")