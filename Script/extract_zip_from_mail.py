import glob
import os
import csv
import string
import sys
import re

# Email Package -----------------------
import extract_msg
import email
from email import policy
from email.message import EmailMessage
from email.utils import parsedate
# OS Package --------------------------
from os.path import exists, dirname, abspath, join, isfile, isdir, basename
from os import makedirs, listdir, mkdir
from shutil import copy, copy2, rmtree
from pathlib import Path
import ntpath
# Datetime Package --------------------
from time import mktime, sleep
from datetime import datetime
# Process Bar Package -----------------
from alive_progress import alive_bar
# -------------------------------------

# OPTION ------------------------------
DELETE_TEMP_FILE_FLAG = False
# -------------------------------------

MAIL_EXTENSION_LIST = ("EML", "EM")
MSG_EXTENSION_LIST = ("MSG",)
EXCEL_EXTENSION_LIST = ("ZIP", "7Z")
EXCL_FILE_EXT = (".BIN", ".BMP", ".CSV", ".DOC", ".DOCX", ".EMZ", ".GIF", ".HTML", ".HTM", 
                ".ICS", ".JPG", ".MP3", ".MP4", ".PDF", ".PNG", ".PPTX", ".PSD", ".TIFF", 
                ".RTF", ".TXT", ".WAV", ".WMZ", ".XLSX", ".XLS", ".XLSB", ".XLSM", ".XLAM")

DATE_RANGE = 'test'
PAF_LL_FOLDER = abspath(r"..\\")
MAIL_FOLDER = PAF_LL_FOLDER + r"\Email\{0}".format(DATE_RANGE)
ATT_FOLDER  = PAF_LL_FOLDER + r"\Attachment\{0}".format(DATE_RANGE)
ERROR_RECORD_FOLDER = PAF_LL_FOLDER + r"\Error Files"
ERROR_RECORD_FILE = ERROR_RECORD_FOLDER + r"\error_files{0}.csv".format("_" + DATE_RANGE)
EXTRACTED_RECORD_FOLDER = PAF_LL_FOLDER + r"\Extract Files"
EXTRACTED_RECORD_FILE = EXTRACTED_RECORD_FOLDER + r"\extracted_files{0}.csv".format("_" + DATE_RANGE)
RENAME_EMAIL_RECORD_FOLDER = PAF_LL_FOLDER + r"\Rename Files"
RENAME_EMAIL_RECORD_FILE = RENAME_EMAIL_RECORD_FOLDER + r"\renamed_mails{0}.csv".format("_" + DATE_RANGE)
TEMP_EMAIL_FOLDER = PAF_LL_FOLDER + r"\Temp Email\{0}".format(DATE_RANGE)
ERROR_EMAIL_FOLDER = PAF_LL_FOLDER + r"\Error Email\{0}".format(DATE_RANGE)

CUR_PATH = dirname(abspath(__file__))
RENAME_TABLE = set(string.ascii_letters + " " + string.digits)

def set_path(new_date_range):
    global DATE_RANGE
    global MAIL_FOLDER
    global ATT_FOLDER
    global ERROR_RECORD_FILE
    global EXTRACTED_RECORD_FILE
    global RENAME_EMAIL_RECORD_FILE
    global TEMP_EMAIL_FOLDER
    global ERROR_EMAIL_FOLDER
    
    DATE_RANGE = new_date_range
    MAIL_FOLDER = PAF_LL_FOLDER + r"\Email\{0}".format(DATE_RANGE)
    ATT_FOLDER  = PAF_LL_FOLDER + r"\Attachment\{0}".format(DATE_RANGE)
    ERROR_RECORD_FOLDER = PAF_LL_FOLDER + r"\Error Files"
    ERROR_RECORD_FILE = ERROR_RECORD_FOLDER + r"\error_files{0}.csv".format("_" + DATE_RANGE)
    EXTRACTED_RECORD_FOLDER = PAF_LL_FOLDER + r"\Extract Files"
    EXTRACTED_RECORD_FILE = EXTRACTED_RECORD_FOLDER + r"\extracted_files{0}.csv".format("_" + DATE_RANGE)
    RENAME_EMAIL_RECORD_FOLDER = PAF_LL_FOLDER + r"\Rename Files"
    RENAME_EMAIL_RECORD_FILE = RENAME_EMAIL_RECORD_FOLDER + r"\renamed_mail{0}.csv".format("_" + DATE_RANGE)
    TEMP_EMAIL_FOLDER = PAF_LL_FOLDER + r"\Temp Email\{0}".format(DATE_RANGE)
    ERROR_EMAIL_FOLDER = PAF_LL_FOLDER + r"\Error Email\{0}".format(DATE_RANGE)

def rename_email():
    files = [file for file in listdir(MAIL_FOLDER) if isfile(join(MAIL_FOLDER, file))]

    with open(RENAME_EMAIL_RECORD_FILE, "w", encoding='utf-8', newline='') as of:
        of_writer = csv.writer(of)
        for file in files:
            split_files = file.split(".")
            new_filename = "unknown"
            
            if len(split_files) > 1:
                ext = split_files[-1]
                filename = ".".join(split_files[:-1])
                old_filename = filename
                filename = ''.join(filter(lambda x: x in RENAME_TABLE, filename))
                filename = re.sub(' +', ' ', filename)
                filename = filename.strip()
                
                if filename != old_filename:
                    if filename == "":
                        filename = "new_name"
                    new_filename = filename + "." + ext
                    index = 0
                    while exists(join(MAIL_FOLDER, new_filename)):
                        new_filename = filename + " ({0})".format(index) + "." + ext
                        index += 1
                    os.rename(join(MAIL_FOLDER, file), join(MAIL_FOLDER, new_filename))
                    of_writer.writerow([file,new_filename,join(MAIL_FOLDER, file), join(MAIL_FOLDER, new_filename)])
            else:
                filename = file
                old_filename = filename
                filename = ''.join(filter(lambda x: x in RENAME_TABLE, filename))
                filename = re.sub(' +', ' ', filename)
                filename = filename.strip()
                
                if filename != old_filename:
                    if filename == "":
                        filename = "new_name"
                    new_filename = filename
                    index = 0
                    while exists(join(MAIL_FOLDER, new_filename)):
                        new_filename = filename + " ({0})".format(index)
                        index += 1
                    os.rename(join(MAIL_FOLDER, file), join(MAIL_FOLDER, new_filename))
                    of_writer.writerow([file,new_filename,join(MAIL_FOLDER, file), join(MAIL_FOLDER, new_filename)])

def rename_file(output_filename, folder_name):
    output_filename_split = output_filename.split(".")
    if len(output_filename_split) > 1:
        file_ext = "." + ''.join(e for e in output_filename_split[-1] if e in RENAME_TABLE)
        att_file_name = "".join(output_filename_split[:-1])
    else:
        file_ext = ""
        att_file_name = output_filename

    new_filename = ''.join(e for e in att_file_name if e in RENAME_TABLE)
    new_filename = re.sub(' +', ' ', new_filename)
    new_filename = new_filename.strip()
    
    if new_filename == "":
        new_filename = "noname"
    else:
        if len(new_filename) > (200 - len(folder_name)):
            new_filename = new_filename[:(200 - len(folder_name))]
            new_filename = new_filename.strip()
    
    check_dup_filename = new_filename
    index_num = 1
    while exists(join(folder_name,check_dup_filename + file_ext)):
        check_dup_filename = new_filename + " ({0})".format(index_num)
        index_num += 1

    check_dup_filename += file_ext

    return folder_name, check_dup_filename

def rename_temp_folder(temp_folder, temp_name):
    temp_name = ''.join(e for e in temp_name if e in RENAME_TABLE)
    temp_name = re.sub(' +', ' ', temp_name)
    temp_name = temp_name.strip()

    if len(temp_name) >= 20:
        temp_name = temp_name[:20].strip()
    new_temp_name = temp_name
    index_num = 1
    while exists(join(temp_folder, new_temp_name)):
        new_temp_name = temp_name + " ({0})".format(index_num)
        index_num += 1

    return temp_folder, temp_name

def extract(root_mail, filename):
    extracted_files = []
    error_files = []
    
    if filename.upper().endswith(EXCL_FILE_EXT):
        pass
    else:
        try:
            if filename.upper().endswith(MSG_EXTENSION_LIST):
                msg = extract_msg.Message(filename)
                if len(msg.attachments) > 0:
                    temp_folder, temp_name = rename_temp_folder(TEMP_EMAIL_FOLDER, basename(filename))
                    output_temp_path = join(temp_folder, temp_name)

                    Path(output_temp_path).mkdir(parents=True, exist_ok=True)
                    try:
                        sent_date = datetime.strptime(" ".join(msg.date.split(", ")[-1].split(" ")[:-1]), '%d %b %Y %H:%M:%S')
                    except AttributeError:
                        sent_date = ""
                    for att in msg.attachments:
                        att_name = att.getFilename()
                        output_temp_path, att_name = rename_file(att_name, output_temp_path)
                        try:
                            att.save(customPath = output_temp_path, customFilename = att_name)
                        except (FileNotFoundError, IndexError):
                            pass

                    for path, subdirs, files in os.walk(output_temp_path):
                        for output_filename in files:
                            abs_out_path = ""
                            # Extract EML, EM, E, MSG files
                            if output_filename.upper().endswith(MAIL_EXTENSION_LIST+MSG_EXTENSION_LIST):
                                try:
                                    abs_out_path = join(CUR_PATH,path,output_filename)

                                    error_temp_files, extracted_temp_files = extract(root_mail, abs_out_path)
                                    for error_temp_file in error_temp_files:
                                        error_temp_file[0] = filename
                                        error_files.append(error_temp_file)
                                    for extracted_temp_file in extracted_temp_files:
                                        extracted_file = [extracted_temp_file[0],filename,sent_date,msg.subject,msg.sender]

                                        extracted_files.append(extracted_file)
                                    extracted_temp_files.append([abs_out_path,filename,sent_date,msg.subject,msg.sender])
                                    
                                except Exception as excep_log:
                                    error_files.append([root_mail, filename, abs_out_path, excep_log])
                                    if not isfile(join(ERROR_EMAIL_FOLDER, ntpath.basename(root_mail))):
                                        copy2(root_mail, ERROR_EMAIL_FOLDER)
                                        print("\t# {0}".format(root_mail))
                                    print("\t[Error MSG-EML/MSG] {0}".format(excep_log))
                                    continue
                            # Extract EXCEL files
                            elif output_filename.upper().endswith(EXCEL_EXTENSION_LIST):
                                abs_folder_path, check_dup_filename = rename_file(output_filename, ATT_FOLDER)

                                try:
                                    old_name_path = join(CUR_PATH, path, output_filename)
                                    abs_out_path = join(abs_folder_path, check_dup_filename)
                                    
                                    copy(old_name_path, abs_out_path)
                                    extracted_files.append([abs_out_path,filename,sent_date,msg.subject,msg.sender])
                                except Exception as excep_log:
                                    error_files.append([root_mail, filename, abs_out_path, excep_log])
                                    if not isfile(join(ERROR_EMAIL_FOLDER, ntpath.basename(root_mail))):
                                        copy2(root_mail, ERROR_EMAIL_FOLDER)
                                        print("\t# {0}".format(root_mail))
                                    print("\t[Error MSG-Excel] {0}".format(excep_log))
                                    continue
                            # Extract all (handle exception)
                            else:
                                try:
                                    abs_out_path = join(output_temp_path, output_filename)
                                    error_temp_files, extracted_temp_files = extract(root_mail, abs_out_path)

                                    for error_temp_file in error_temp_files:
                                        error_temp_file[0] = filename
                                        error_files.append(error_temp_file)
                                    for extracted_temp_file in extracted_temp_files:
                                        extracted_file = [extracted_temp_file[0],filename,sent_date,msg.subject,msg.sender]

                                        extracted_files.append(extracted_file)
                                    extracted_temp_files.append([abs_out_path,filename,sent_date,msg.subject,msg.sender])
                                
                                # Delete this except if you want to get all attachment file (will be in Temp Email Folder)
                                except TypeError:
                                    error_files.append([root_mail, filename, abs_out_path, "TypeError"])
                                    continue
                                except Exception as excep_log:
                                    error_files.append([root_mail, filename, abs_out_path, excep_log])
                                    if not isfile(join(ERROR_EMAIL_FOLDER, ntpath.basename(root_mail))):
                                        copy2(root_mail, ERROR_EMAIL_FOLDER)
                                        print("\t# {0}".format(root_mail))
                                    print("\t[Error MSG-Any] {0}".format(excep_log))
                                    continue
                    rmtree(output_temp_path)
                msg.close()

            else:
                with open(filename, "rb") as f:
                    try:
                        msg = email.message_from_bytes(f.read(), _class=EmailMessage)
                        for attachment in msg.get_payload():
                            if attachment != None:
                                _ = attachment.get_filename()
                    except AttributeError:
                        msg = email.message_from_bytes(f.read(), _class=EmailMessage, policy=policy.default)
                    for attachment in msg.get_payload():
                        if attachment != None:
                            try:
                                output_filename = attachment.get_filename()
                            except AttributeError:
                                continue
                            # If no attachments are found, skip this file
                            if output_filename:
                                date_str=msg.get('date')
                                if date_str:
                                    date_tuple=email.utils.parsedate_tz(date_str)
                                    if date_tuple:
                                        sent_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                                    else:
                                        sent_date = ""
                                else:
                                    sent_date = ""

                                #sent_date = datetime.fromtimestamp(mktime(parsedate(msg.get('date')))).strftime("%Y-%m-%d %H:%M:%S")
                                # Extract EML, EM, E, MSG files
                                if output_filename.upper().endswith(MAIL_EXTENSION_LIST+MSG_EXTENSION_LIST):
                                    abs_folder_path, check_dup_filename = rename_file(output_filename, TEMP_EMAIL_FOLDER)
                                    abs_out_path = join(abs_folder_path,check_dup_filename)

                                    try:
                                        att_data_list = attachment.get_payload(decode=True) or attachment.get_payload()
                                        if not isinstance(att_data_list, list):
                                            att_data_list = [att_data_list]
                                        else:
                                            att_data_list = [bytes(att_data) for att_data in att_data_list]

                                        for att_data in att_data_list:
                                            if att_data != None:
                                                with open(abs_out_path, "wb") as of:
                                                    of.write(att_data)

                                                error_temp_files, extracted_temp_files = extract(root_mail, abs_out_path)
                                                for error_temp_file in error_temp_files:
                                                    error_temp_file[0] = filename
                                                    error_files.append(error_temp_file)
                                                for extracted_temp_file in extracted_temp_files:
                                                    extracted_file = [extracted_temp_file[0],filename,sent_date,msg.get('subject'),msg.get('from')]

                                                    extracted_files.append(extracted_file)
                                                extracted_temp_files.append([abs_out_path,filename,sent_date,msg.get('subject'),msg.get('from')])
                                            
                                    except Exception as excep_log:
                                        error_files.append([root_mail, filename, abs_out_path, excep_log])
                                        if not isfile(join(ERROR_EMAIL_FOLDER, ntpath.basename(root_mail))):
                                            copy2(root_mail, ERROR_EMAIL_FOLDER)
                                            print("\t# {0}".format(root_mail))
                                        print("\t[Error EML-EML/MSG] {0}".format(excep_log))
                                        continue
                                # Extract EXCEL files
                                elif output_filename.upper().endswith(EXCEL_EXTENSION_LIST):
                                    abs_folder_path, check_dup_filename = rename_file(output_filename, ATT_FOLDER)
                                    abs_out_path = join(abs_folder_path,check_dup_filename)

                                    try:
                                        att_data = attachment.get_payload(decode=True)
                                        if att_data != None:
                                            with open(abs_out_path, "wb") as of:
                                                of.write(att_data)
                                            extracted_files.append([abs_out_path,filename,sent_date,msg.get('subject'),msg.get('from')])
                                        
                                    except Exception as excep_log:
                                        error_files.append([root_mail, filename, abs_out_path, excep_log])
                                        if not isfile(join(ERROR_EMAIL_FOLDER, ntpath.basename(root_mail))):
                                            copy2(root_mail, ERROR_EMAIL_FOLDER)
                                            print("\t# {0}".format(root_mail))
                                        print("\t[Error EML-Excel] {0}".format(excep_log))
                                        continue
                                # Extract all (handle exception)
                                else:
                                    abs_folder_path, check_dup_filename = rename_file(output_filename, TEMP_EMAIL_FOLDER)
                                    abs_out_path = join(abs_folder_path, check_dup_filename)

                                    try:
                                        att_data_list = attachment.get_payload(decode=True) or attachment.get_payload()
                                        if not isinstance(att_data_list, list):
                                            att_data_list = [att_data_list]
                                        else:
                                            att_data_list = [bytes(att_data) for att_data in att_data_list]

                                        for att_data in att_data_list:
                                            if att_data != None:
                                                try:
                                                    with open(abs_out_path, "wb") as of:
                                                        of.write(att_data)
                                                except TypeError:
                                                    with open(abs_out_path, "w") as of:
                                                        of.write(att_data)

                                                error_temp_files, extracted_temp_files = extract(root_mail, abs_out_path)
                                                for error_temp_file in error_temp_files:
                                                    error_temp_file[0] = filename
                                                    error_files.append(error_temp_file)
                                                for extracted_temp_file in extracted_temp_files:
                                                    extracted_file = [extracted_temp_file[0],filename,sent_date,msg.get('subject'),msg.get('from')]

                                                    extracted_files.append(extracted_file)
                                                extracted_temp_files.append([abs_out_path,filename,sent_date,msg.get('subject'),msg.get('from')])
                                    # Delete this except if you want to get all attachment file (will be in Temp Email Folder)
                                    except TypeError:
                                        error_files.append([root_mail, filename, abs_out_path, "TypeError"])
                                        os.remove(abs_out_path)
                                        continue
                                    except Exception as excep_log:
                                        error_files.append([root_mail, filename, abs_out_path, excep_log])
                                        if not isfile(join(ERROR_EMAIL_FOLDER, ntpath.basename(root_mail))):
                                            copy2(root_mail, ERROR_EMAIL_FOLDER)
                                            print("\t# {0}".format(root_mail))
                                        print("\t[Error EML-Any] {0}".format(excep_log))
                                        continue
                                
        # This should catch read and write errors
        except extract_msg.exceptions.StandardViolationError:
            pass
        except AttributeError:
            error_files.append([filename, "No file", 'AttributeError'])
            copy2(filename, ERROR_EMAIL_FOLDER)
            print("[Attribute Error] ", filename)
        except FileNotFoundError:
            if filename.upper().endswith(MAIL_EXTENSION_LIST + MSG_EXTENSION_LIST + EXCEL_EXTENSION_LIST):
                error_files.append([filename, "Missing file", excep_log])
        except Exception as excep_log:
            error_files.append([filename, "No file", excep_log])
                                
    return error_files, extracted_files

# ----------------------------------------------------------------------------
# MAIN -------------------

def main(argv):
    header = '''
                   ███████╗███╗░░░███╗░█████╗░██╗██╗░░░░░
                   ██╔════╝████╗░████║██╔══██╗██║██║░░░░░
                   █████╗░░██╔████╔██║███████║██║██║░░░░░
                   ██╔══╝░░██║╚██╔╝██║██╔══██║██║██║░░░░░
                   ███████╗██║░╚═╝░██║██║░░██║██║███████╗
                   ╚══════╝╚═╝░░░░░╚═╝╚═╝░░╚═╝╚═╝╚══════╝
 ███████╗██╗░░██╗████████╗██████╗░░█████╗░░█████╗░████████╗░█████╗░██████╗░
 ██╔════╝╚██╗██╔╝╚══██╔══╝██╔══██╗██╔══██╗██╔══██╗╚══██╔══╝██╔══██╗██╔══██╗
 █████╗░░░╚███╔╝░░░░██║░░░██████╔╝███████║██║░░╚═╝░░░██║░░░██║░░██║██████╔╝
 ██╔══╝░░░██╔██╗░░░░██║░░░██╔══██╗██╔══██║██║░░██╗░░░██║░░░██║░░██║██╔══██╗
 ███████╗██╔╝╚██╗░░░██║░░░██║░░██║██║░░██║╚█████╔╝░░░██║░░░╚█████╔╝██║░░██║
 ╚══════╝╚═╝░░╚═╝░░░╚═╝░░░╚═╝░░╚═╝╚═╝░░╚═╝░╚════╝░░░░╚═╝░░░░╚════╝░╚═╝░░╚═╝'''

    print(header)
    print("\n\n")
    print("Version: {0}".format("1.1"))

    if "--range" in argv:
        set_path(argv[argv.index("--range") + 1])
    else:
        set_path(input("Date range input: "))
    if "--rename" in argv:
        print("REMOVE SPECIAL CHAR IN EMAIL NAMES...")
        print("-------------------------------------")
        rename_email()
    if "--clear-temp" in argv:
        global DELETE_TEMP_FILE_FLAG
        DELETE_TEMP_FILE_FLAG = True
    
    # ensure that an output dir exists
    folder_path_list = [
                    ERROR_RECORD_FOLDER, 
                    EXTRACTED_RECORD_FOLDER, 
                    RENAME_EMAIL_RECORD_FOLDER
    ]
    
    for folder_path in folder_path_list:
        os.path.exists(folder_path) or os.makedirs(folder_path)
    
    # Clear Attachment Folder if exists
    if os.path.exists(ATT_FOLDER):
        try:
            rmtree(ATT_FOLDER)
            os.makedirs(ATT_FOLDER)
        except PermissionError:
            pass
    else:
        os.makedirs(ATT_FOLDER)
    
    # Clear Temp Email Folder (Email attach Email) folder if exists
    if os.path.exists(TEMP_EMAIL_FOLDER):
        try:
            rmtree(TEMP_EMAIL_FOLDER)
            os.makedirs(TEMP_EMAIL_FOLDER)
        except PermissionError:
            pass
    else:
        os.makedirs(TEMP_EMAIL_FOLDER)
    
    # Clear Error Email Folder folder if exists
    if os.path.exists(ERROR_EMAIL_FOLDER):
        try:
            rmtree(ERROR_EMAIL_FOLDER)
            os.makedirs(ERROR_EMAIL_FOLDER)
        except PermissionError:
            pass
    else:
        os.makedirs(ERROR_EMAIL_FOLDER)
    
    email_list = []
    for mail_ext in MAIL_EXTENSION_LIST+MSG_EXTENSION_LIST:
        email_list += glob.glob(join(MAIL_FOLDER, "*.{0}".format(mail_ext)))

    extracted_total = []
    error_total = []
    # file_info : ( Attachment file path, Email path, Sent Date, Subject, From )
    with alive_bar(len(email_list), title='Emails: ', theme='smooth') as bar:
        for email_file in email_list:
            sleep(.001)
            bar()
            error_files, extracted_files = extract(email_file, email_file)
            extracted_total.append(extracted_files)
            error_total.append(error_files)

    with open(ERROR_RECORD_FILE,"w",newline="", encoding='utf-8') as error_file_stream:
        error_writer = csv.writer(error_file_stream)
        for error_files in error_total:
            #for error_file in error_files:
                #copy2(error_file[0], ERROR_EMAIL_FOLDER)
            error_writer.writerows(error_files)

    with open(EXTRACTED_RECORD_FILE, "w", encoding="utf-8", newline="") as extracted_file_stream:
        extracted_writer = csv.writer(extracted_file_stream)
        for extracted_files in extracted_total:
            extracted_writer.writerows(extracted_files)

    if DELETE_TEMP_FILE_FLAG == True:
        sleep(7)
        print("REMOVE TEMP EMAIL FOLDER...")
        try:
            rmtree(TEMP_EMAIL_FOLDER)
        except:
            print("\tError")
        print("-------------------------------------")
    print("\t\t--- FINISH ---")

if __name__ == "__main__":
    main(sys.argv)