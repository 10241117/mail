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

def get_basic_info(lines, idx):
    basic_info = {}
    try:
        for _ in range(3):
            line = lines[idx].strip('\n')
            idx += 1
            [key, value] = line.split(':', maxsplit=1)
            basic_info[basic_dict[key]] = value.strip()
        basic_info['InsertDate'] = date.today().strftime("%m/%d/%y")
    except:
        print(f'\033[91m取得基本資訊時發生錯誤\033[0m')
    return basic_info

def get_mb_info(lines, idx, basic):
    mb_info = dict(basic)
    try:
        for _ in range(4):
            line = lines[idx].strip('\n')
            idx += 1
            [key, value] = line.split(':', maxsplit=1)
            mb_info[mb_os_dict[key]] = value.strip()
    except:
        print(f'\033[91m取得主機板資訊時發生錯誤\033[0m')
    return mb_info

def append_os_info(lines, idx, orig):
    try:
        for _ in range(2):
            line = lines[idx].strip('\n')
            idx += 1
            [key, value] = line.split(':', maxsplit=1)
            orig[mb_os_dict[key]] = value.strip()
    except:
        print(f'\033[91m取得作業系統資訊時發生錯誤\033[0m')
    return orig

def append_single_info(line, orig):
    line = line.strip('\n')
    [key, value] = line.split(':', maxsplit=1)
    orig[mb_os_dict[key]] = value.strip()
    return orig

def get_ram_info(lines, idx, basic):
    ram_info = dict(basic)
    try:
        for _ in range(5):
            line = lines[idx].strip('\n')
            idx += 1
            if ':' in line:
                [key, value] = line.split(':', maxsplit=1)
                ram_info[ram_dict[key]] = value.strip()
    except:
        print(f'\033[91m取得記憶體資訊時發生錯誤\033[0m')
    return ram_info

def get_nic_info(line, basic, ip_len):
    nic_info = dict(basic)
    try:
        nic_ary = []
        end = len(line)
        for size in ip_len:
            nic_ary.append(line[end-size:end])
            end -= size
        nic_ary.append(line[0:end])
        nic_ary.reverse()
        for i in range(len(nic_ary)):
            nic_info[nic_title[i]] = nic_ary[i].strip()
    except:
        print(f'\033[91m取得網路卡資訊時發生錯誤\033[0m')
    return nic_info

def get_printer_info(line, basic):
    printer_info = dict(basic)
    try:
        printer_ary = line.split(' 的 IPAddr是 ')
        for i in range(len(printer_title)):
            printer_info[printer_title[i]] = printer_ary[i]
    except:
        print(f'\033[91m取得印表機資訊時發生錯誤\033[0m')
    return printer_info

def get_hdd_info(line, basic):
    hdd_info = dict(basic)
    try:
        hdd_string = ' '.join(line.split())
        hdd_ary = hdd_string.split(' ', 1)
        remain = hdd_ary.pop()
        hdd_ary += remain.rsplit(' ', 1)
        for i in range(len(hdd_title)):
            hdd_info[hdd_title[i]] = hdd_ary[i]
    except:
        print(f'\033[91m取得硬碟資訊時發生錯誤\033[0m')
    return hdd_info


path = './' + date.today().strftime("%Y_%m_%d")
try:
    os.mkdir(path)
    print(f"創建資料夾成功: {date.today().strftime("%Y_%m_%d")}")
except:
    print(f"資料夾已存在: {date.today().strftime("%Y_%m_%d")}")
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

data_fields = ['電腦主機名稱及當前登入的使用者', '主機板資訊', '作業系統資訊', 'CPU資訊', '當地時間', '最後一次更新的資訊', '記憶體資訊', '網路卡IP資訊', '網路印表機資訊及硬碟資訊']
print('=====')
for filename in os.listdir(os.getcwd()):
    if filename.endswith('.txt') and ('作業系統更新狀況' in filename or 'ComputerInfo' in filename):
        # open file
        with open(os.path.join(os.getcwd(), filename), 'r') as f:
            print(filename)
            f_fileds = []
            lines = f.readlines()
            for idx, line in enumerate(lines):
                if '取得電腦主機名稱及當前登入的使用者' not in line:
                    continue
                basic_info = get_basic_info(lines, idx+1)
                f_fileds.append(data_fields[0])
                break

            for idx, line in enumerate(lines):
                if '取得電腦主機的主機板資訊' not in line:
                    continue
                mb_os_info = get_mb_info(lines, idx+1, basic_info)
                f_fileds.append(data_fields[1])
                break

            for idx, line in enumerate(lines):
                if '作業系統資訊' not in line:
                    continue
                mb_os_info = append_os_info(lines, idx+1, mb_os_info)
                f_fileds.append(data_fields[2])
                break

            for idx, line in enumerate(lines):
                if 'CPU資訊' not in line:
                    continue
                try:
                    mb_os_info = append_single_info(lines[idx+1], mb_os_info)
                except:
                    print(f'\033[91m取得CPU資訊時發生錯誤\033[0m')
                f_fileds.append(data_fields[3])
                break

            for idx, line in enumerate(lines):
                if '當地時間:' not in line:
                    continue
                try:
                    mb_os_info = append_single_info(lines[idx], mb_os_info)
                except:
                    print(f'\033[91m取得當地時間時發生錯誤\033[0m')
                f_fileds.append(data_fields[4])
                break

            for idx, line in enumerate(lines):
                if '系統最後更新時間:' not in line:
                    continue
                try:
                    mb_os_info = append_single_info(lines[idx], mb_os_info)
                except:
                    print(f'\033[91m取得最後一次更新的資訊時發生錯誤\033[0m')
                f_fileds.append(data_fields[5])
                break
            
            for idx, line in enumerate(lines):
                if '記憶體資訊' not in line:
                    continue
                idx += 1
                while '製造商' in lines[idx]:
                    ram_info = get_ram_info(lines, idx, basic_info)
                    idx += 5
                    ram_writer.writerow(ram_info)
                f_fileds.append(data_fields[6])
                break

            for idx, line in enumerate(lines):
                if '網路卡IP資訊' not in line:
                    continue
                idx += 1
                # skip above lines
                while 'Name' not in lines[idx]:
                    idx += 1
                line = lines[idx].strip('\n')
                ip_len = []
                ip_len.append(line.find('Status') - line.find('MacAddress'))
                ip_len.append(line.find('IPAddress') - line.find('Status'))
                ip_len.append(len(line) - line.find('IPAddress'))
                ip_len.reverse()
                # read nic section
                # print(ip_len)
                idx += 2
                line = lines[idx].strip('\n')
                while line != '':
                    nic_names = ('乙太網路', 'Wi-Fi', '藍牙網路連線', '區域連線', 'Ethernet')
                    if(line.startswith(nic_names)):
                        nic_info = get_nic_info(line, basic_info, ip_len)
                        # print(line)
                        # print(nic_info)
                        nic_writer.writerow(nic_info)
                    idx += 1
                    line = lines[idx].strip('\n')
                f_fileds.append(data_fields[7])
                break

            for idx, line in enumerate(lines):
                if '取得網路印表機資訊及硬碟資訊' not in line:
                    continue
                idx += 1
                line = lines[idx].strip('\n')
                while line != '':
                    printer_info = get_printer_info(line, basic_info)
                    printer_writer.writerow(printer_info)
                    idx += 1
                    line = lines[idx].strip('\n')
                # skip above lines
                while '---' not in lines[idx]:
                    idx += 1
                idx += 1
                # read hdd section
                line = lines[idx].strip('\n')
                while line != '':
                    hdd_info = get_hdd_info(line, basic_info)
                    hdd_writer.writerow(hdd_info)
                    idx += 1
                    line = lines[idx].strip('\n')
                f_fileds.append(data_fields[8])
                break

            mb_os_writer.writerow(mb_os_info)
            for field in data_fields:
                if field not in f_fileds:
                    print(f'\033[93m{field}不存在\033[0m')
            print('=====')
mb_os_csv.close()
ram_csv.close()
nic_csv.close()
printer_csv.close()
hdd_csv.close()