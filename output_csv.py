import os
import csv
from datetime import date

basic_dict = {
    '電腦主機名稱': 'ComputerName',
    '使用者名稱': 'UserName',
    '網域\工作站': 'WorkDomain',
}

mb_os_dict = {
    '製造商': 'MbManufacturer',
    '產品': 'MBType',
    '序列號': 'MBSerial',
    '版本': 'MBVersion',
    'Operating System': 'OSName',
    'Version': 'OSVersion',
    'CPU Model': 'CPUModel',
    '當地時間': 'TestDate',
    '系統最後更新時間': 'LastUpdateDate',
}

ram_dict = {
    '製造商': 'RamManufacturer',
    '速度': 'RamSpeed',
    '容量': 'RamSize',
    '型號': 'RamType',
}
basic_title = ['InsertDate', 'ComputerName', 'UserName', 'WorkDomain']
mb_os_title = ['MbManufacturer', 'MBType', 'MBSerial', 'MBVersion', 'OSName', 'OSVersion', 'CPUModel', 'TestDate', 'LastUpdateDate']
ram_title = ['RamManufacturer', 'RamSpeed', 'RamSize', 'RamType']
nic_title = ['NicName', 'MacAddress', 'NicStatus', 'IPAddress']
printer_title = ['PrintName', 'PrintIPAddr']
hdd_title = ['HDDID', 'HDDModel', 'HDDSizeGB']

def get_basic_info(file):
    basic_info = {}
    for _ in range(len(basic_dict)):
        line = file.readline().strip('\n')
        [key, value] = line.split(':', maxsplit=1)
        basic_info[basic_dict[key]] = value.strip()
    basic_info['InsertDate'] = date.today().strftime("%m/%d/%y")
    return basic_info

def get_mb_info(file, basic):
    mb_info = dict(basic)
    for _ in range(4):
        line = file.readline().strip('\n')
        [key, value] = line.split(':', maxsplit=1)
        mb_info[mb_os_dict[key]] = value.strip()
    return mb_info

def append_os_info(file, orig):
    for _ in range(2):
        line = file.readline().strip('\n')
        [key, value] = line.split(':', maxsplit=1)
        orig[mb_os_dict[key]] = value.strip()
    return orig

def append_cpu_info(file, orig):
    line = file.readline().strip('\n')
    [key, value] = line.split(':', maxsplit=1)
    orig[mb_os_dict[key]] = value.strip()
    return orig

def get_ram_info(file, basic, line):
    ram_info = dict(basic)
    for _ in range(5):
        if ':' in line:
            [key, value] = line.split(':', maxsplit=1)
            ram_info[ram_dict[key]] = value.strip()
        line = file.readline().strip('\n')
    return ram_info, line

def get_nic_info(line, basic):
    nic_info = dict(basic)
    modified = ' '.join(line.split())
    nic_ary = modified.rsplit(' ', len(nic_title) - 1)
    for i in range(len(nic_title)):
        nic_info[nic_title[i]] = nic_ary[i]
    return nic_info

def get_printer_info(line, basic):
    printer_info = dict(basic)
    printer_ary = line.split(' 的 IPAddr是 ')
    for i in range(len(printer_title)):
        printer_info[printer_title[i]] = printer_ary[i]
    return printer_info

def get_hdd_info(line, basic):
    hdd_info = dict(basic)
    hdd_ary = ' '.join(line.split()).split(' ')
    for i in range(len(hdd_title)):
        hdd_info[hdd_title[i]] = hdd_ary[i]
    return hdd_info



path = './' + date.today().strftime("%Y_%m_%d")
try:
    os.mkdir(path)
    print("創建資料夾成功")
except:
    print("創建資料夾失敗")
try:
    mb_os_csv = open(path+'/電腦主機版_作業系統資訊.csv', 'w', newline='', encoding='UTF-8')
    mb_os_writer = csv.DictWriter(mb_os_csv, fieldnames=basic_title+mb_os_title)
    mb_os_writer.writeheader()
    ram_csv = open(path+'/電腦記憶體資訊.csv', 'w', newline='', encoding='UTF-8')
    ram_writer = csv.DictWriter(ram_csv, fieldnames=basic_title+ram_title)
    ram_writer.writeheader()
    nic_csv = open(path+'/電腦網路卡資訊.csv', 'w', newline='', encoding='UTF-8')
    nic_writer = csv.DictWriter(nic_csv, fieldnames=basic_title+nic_title)
    nic_writer.writeheader()
    printer_csv = open(path+'/電腦印表機資訊.csv', 'w', newline='', encoding='UTF-8')
    printer_writer = csv.DictWriter(printer_csv, fieldnames=basic_title+printer_title)
    printer_writer.writeheader()
    hdd_csv = open(path+'/電腦硬碟資訊.csv', 'w', newline='', encoding='UTF-8')
    hdd_writer = csv.DictWriter(hdd_csv, fieldnames=basic_title+hdd_title)
    hdd_writer.writeheader()
except:
    print("創建csv檔失敗")
    mb_os_csv.close()
    ram_csv.close()
    nic_csv.close()
    printer_csv.close()
    hdd_csv.close()
    exit()


for filename in os.listdir(os.getcwd()):
    if filename.endswith('.txt') and ('作業系統更新狀況' in filename or 'ComputerInfo' in filename):
        # open file
        with open(os.path.join(os.getcwd(), filename), 'r') as f:
            print(filename)
            line = f.readline().strip('\n')
            while line is not None and line != '':
                if '取得電腦主機名稱及當前登入的使用者' in line:
                    basic_info = get_basic_info(f)
                elif '取得電腦主機的主機板資訊' in line:
                    mb_os_info = get_mb_info(f, basic_info)
                elif '作業系統資訊' in line:
                    mb_os_info = append_os_info(f, mb_os_info)
                elif 'CPU資訊' in line or '取得當地時間' in line or '取得最後一次更新的資訊' in line:
                    mb_os_info = append_cpu_info(f, mb_os_info)
                elif '記憶體資訊' in line:
                    line = f.readline().strip('\n')
                    while '製造商' in line:
                        ram_info, line = get_ram_info(f, basic_info, line)
                        ram_writer.writerow(ram_info)
                elif '網路卡IP資訊' in line:
                    # skip above lines
                    while '---' not in line:
                        line = f.readline().strip('\n')
                    # read nic section
                    line = f.readline().strip('\n')
                    while line != '':
                        nic_info = get_nic_info(line, basic_info)
                        nic_writer.writerow(nic_info)
                        line = f.readline().strip('\n')
                elif '取得網路印表機資訊及硬碟資訊' in line:
                    line = f.readline().strip('\n')
                    while line != '':
                        printer_info = get_printer_info(line, basic_info)
                        printer_writer.writerow(printer_info)
                        line = f.readline().strip('\n')
                    # skip above lines
                    while '---' not in line:
                        line = f.readline().strip('\n')
                    # read hdd section
                    line = f.readline().strip('\n')
                    while line != '':
                        hdd_info = get_hdd_info(line, basic_info)
                        hdd_writer.writerow(hdd_info)
                        line = f.readline().strip('\n')

                line = f.readline()
                # print(line, end='')
            mb_os_writer.writerow(mb_os_info)

mb_os_csv.close()
ram_csv.close()
nic_csv.close()
printer_csv.close()
hdd_csv.close()