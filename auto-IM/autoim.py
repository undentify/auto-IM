from docxtpl import DocxTemplate
import datetime
import pandas
import openpyxl
import numpy
import re
import os

version = '1.0'


def read_props(file):
    props = {
        'author':'UNKNOWN',
        'PAC_name': 'UNKNOWN',
        'PAC_type': 'UNKNOWN',
        'OS_version': 'UNKNOWN'
    }
    f = open(file, 'r', encoding='utf-8')
    for line in f:
        variable = line[:re.search('=',line).start()]
        if variable in ['author', 'PAC_name', 'PAC_type', 'OS_version']:
            props[variable] = line[re.search('\".*\"',line).start()+1:re.search('\".*\"',line).end()-1]
        else:
            print('Variable '+variable+' not supported. Skipped.')
    f.close()
    return props

def doc_search(path):
    docPath = {
        'netpassport': 'UNKNOWN',
        'cablejournal': 'UNKNOWN',
        'specification': 'UNKNOWN',
        'raid':'UNKNOWN',
        'vmlist': 'UNKNOWN',
        'lsblk': 'UNKNOWN',
        'fstab': 'UNKNOWN',
        'report': 'UNKNOWN'
    }
    path=path+slash_type+'res'
    print('Поиск документов:', path)
    files = os.listdir(path)
    symbol=slash_type
    for filename in files:
        if '~' in filename:
            pass
        else:
            if 'passport' in filename.lower():
                docPath['netpassport']=path + symbol + filename
            if 'сеть' in filename.lower():
                docPath['cablejournal'] = path + symbol + filename
            if 'спецификация' in filename.lower():
                docPath['specification'] = path + symbol + filename
            if 'raid' in filename.lower():
                docPath['raid'] = path + symbol + filename
            if 'vm-hw' in filename.lower():
                docPath['vmlist'] = path + symbol + filename
            if 'lsblk' in filename.lower():
                docPath['lsblk'] = path + symbol + filename
            if 'fstab' in filename.lower():
                docPath['fstab'] = path + symbol + filename
            if 'report' in filename.lower():
                docPath['report'] = path + symbol + filename
    print('Сетевой паспорт найден:',docPath['netpassport'])
    print('Кабельный журнал найден:', docPath['cablejournal'])
    print('Спецификация найдена:', docPath['specification'])
    print('Настройки RAID найдены:', docPath['raid'])
    print('Данные по виртуальным машинам найдены:', docPath['vmlist'])
    print('Данные по дисковым подсистемам найдены::', docPath['lsblk'])
    print('Данные по параметрам монтирования найдены:', docPath['fstab'])
    print('Отчет по установке найден:', docPath['report'])
    return docPath
def prepare_spec(specPath='Спецификация ADB.DR_2586.xlsx'):
    print('Preparing specification:', specPath)
    df_spec = pandas.read_excel(specPath)
    colsSpec = list(df_spec)
    newColsSpec = []
    for item in colsSpec:
        val = df_spec[item].iloc[2]
        newItem = item
        if type(val) is str:
            if len(val) > 40:
                newItem = 'spec'
            if val == 'Master' or val == 'Segment':
                newItem = 'comment'
        if type(val) == numpy.float64 or type(val) == float:
            if not numpy.isnan(val):
                newItem = 'amount'
        newColsSpec.append(newItem)
    df_spec = df_spec.rename(columns=dict(zip(colsSpec, newColsSpec)))
    for item in newColsSpec:
        if 'Unnamed' in item:
            del df_spec[item]
    df_spec=df_spec.drop(labels=[0],axis=0)
    df_spec['amount']=df_spec['amount'].astype(int)
    return df_spec
def spec_to_json(df,date='2024-1'):
    print('preparing JSON for docxtpl for Table1:')
    json_spec = {}
    dict_spec = []
    for i in range(1,df.shape[0]):
        dict_spec.append({"cols": {'spec': df['spec'].iloc[i],'amount': df['amount'].iloc[i],'comment': df['comment'].iloc[i],'date': date,'articul': ''}})
    json_spec = dict_spec
    return json_spec

def prepare_cj(CJPath='ЗИ Сеть L1-L3 ПАК_1639 ECO DR v1.xlsx'):
    print('Preparing Cable Journal:', CJPath)
    df_cj = pandas.read_excel(CJPath, sheet_name='L1 links')
    del df_cj['[S] Description']
    del df_cj['[S] Model']
    del df_cj['[S] Rack']
    del df_cj['[S] Units']
    del df_cj['[S] Port Phys. Type']
    del df_cj['[D] Description']
    del df_cj['[D] Model']
    del df_cj['[D] Rack']
    del df_cj['[D] Units']
    df_cj = df_cj.fillna('')
    return df_cj

def cj_to_json(df):
    print('preparing JSON for docxtpl for Table3:')
    json_cj = {}
    dict_cj = []
    num = 0
    for i in range(df.shape[0]):
        if df['[S] Hostname'].iloc[i]=='' or df['[D] Hostname'].iloc[i]=='':
            pass
        else:
            num+=1
            portType = df['[D] Port Phys. Type'].iloc[i]
            portSpeed = ''
            if '10G' in portType:
                portSpeed = '10'
            elif '100G' in portType:
                portSpeed = '100'
            elif '1000' in portType:
                portSpeed = '1'
            elif '25' in portType:
                portSpeed = '25'
            else:
                portSpeed = '-'
            dict_cj.append({"cols": {
                    'num': num,
                    's_name': df['[S] Hostname'].iloc[i],
                    's_port': df['[S] Port'].iloc[i],
                    'd_name': df['[D] Hostname'].iloc[i],
                    'd_port': df['[D] Port'].iloc[i],
                    'speed': portSpeed,
                    'type': portType}
                })
    json_CJ = dict_cj
    return json_CJ

def vm_to_json(path='vm-hw.txt', OS_version='UNKNOWN'):
    print('preparing JSON for docxtpl for Table2:')
    with open(path, 'r') as file:
        lines = file.readlines()
    json_vm={}
    struct=[[]]
    i=0
    for line in lines:
        if 'Name: ' in line:
            i+=1
            struct.append([])
        struct[i].append(line)
    if struct[0]==[]:
        struct.pop(0)
    dict_vm = []
    for obj in struct:
        tmp_dict = {
            'name': 'UNKNOWN',
            'OS': 'UNKNOWN',
            'vCPU': 'UNKNOWN',
            'RAM': 'UNKNOWN',
            'space': 'UNKNOWN',
            'diskType': 'UNKNOWN',
            'comment': 'Использовать GPT'}
        tmp_dict['OS']=OS_version
        for line in obj:
            if 'Name: ' in line:
                vm_name = line[6:-1]
                tmp_dict['name'] = vm_name
            if 'cores=' in line:
                vm_cpu = re.split(' ', line[re.search('cores=', line).end():])[0]
                tmp_dict['vCPU'] = vm_cpu
            if 'memory ' in line:
                vm_mem = line[9:-1]
                if vm_mem[-2:] == 'Mb':
                    vm_mem = int(float(vm_mem[:-2]) / 1024)
                tmp_dict['RAM'] = vm_mem
            if 'hdd0' in line and 'Boot order' not in line:
                s = re.split(' ', line)
                for l in s:
                    if 'image' and 'hdd' in l:
                        vm_hdd_type = 'HDD'
                    if 'image' and 'ssd' in l:
                        vm_hdd_type = 'SSD'
                    if 'Mb' in l:
                        vm_hdd = int(float(l[:-2]) / 1024)
                tmp_dict['space'] = vm_hdd
                tmp_dict['diskType'] = vm_hdd_type
        dict_vm.append({"cols": tmp_dict})
    json_vm = dict_vm
    return json_vm

def find_repeats_in_dict(list):
    repeatList=[]
    for i in range(len(list)):
        for j in range(i+1,len(list)):
            if list[i] == list[j] and list[i] not in repeatList:
                repeatList.append(list[i])
    repeatDict={}
    for item in repeatList:
        count=0
        for name in list:
            if item == name:
                count+=1
        repeatDict[item]=count
    for item in repeatList:
        for i in reversed(range(len(list))):
            if item == list[i]:
                list[i]=list[i]+str(repeatDict[item])
                repeatDict[item]-=1
    new_list=list
    return new_list

def cut_netpassport(df):
    new_df = df.loc[df['DNS-имя'] != '']
    return new_df
def prepare_netpassport(PassPath='1639_DREV_ADB_DR_NET_PASSPORT_v4.xlsx'):
    #обработка листа "Серверы"
    print('Preparing NetPassport:', PassPath)
    df_srv = pandas.read_excel(PassPath, sheet_name='Серверы')
    df_srv = df_srv.fillna('')
    newCols_srv = []
    for l1, l2 in zip(df_srv.T[0].tolist(), df_srv.T[1].tolist()):
        if l2 == '':
            newCols_srv.append(l1)
        else:
            newCols_srv.append(l2)
    newCols_srv = find_repeats_in_dict(newCols_srv)
    #print('new Cols: ', newCols_srv)
    df_srv = df_srv.drop(index=[0, 1], axis=0)
    df_srv.reset_index(drop=True, inplace=True)
    df_srv = df_srv.rename(columns=dict(zip(list(df_srv), newCols_srv)))
    df_srv.replace('Не используется', '', inplace=True)
    df_srv = df_srv.loc[df_srv['DNS-имя1'] != '']
    #print(df_srv)

    #обработка листа "Коммутаторы"
    df_sw = pandas.read_excel(PassPath, sheet_name='Коммутаторы')
    df_sw = df_sw.fillna('')
    newCols_sw = []
    for l1, l2 in zip(list(df_sw),df_sw.T[0].tolist()):
        if l2 == '':
            newCols_sw.append(l1)
        else: newCols_sw.append(l2)
    newCols_sw=find_repeats_in_dict(newCols_sw)
    #print('new Cols: ',newCols_sw)
    df_sw = df_sw.drop(index=[0], axis=0)
    df_sw.reset_index(drop= True, inplace= True)
    df_sw=df_sw.rename(columns=dict(zip(list(df_sw), newCols_sw)))
    df_sw.replace('Не используется', '', inplace=True)
    df_sw = df_sw.loc[df_sw['DNS-имя'] != '']
    #print(df_sw)

    #обработка листа "Сервисный порт ПАК" (если такой есть)
    try:
        df_port = pandas.read_excel(PassPath, sheet_name='Сервисный порт ПАК')
        df_port = df_port.fillna('')
        newCols_port = []
        for l1, l2 in zip(list(df_port), df_port.T[0].tolist()):
            if l2 == '':
                newCols_port.append(l1)
            else:
                newCols_port.append(l2)
        newCols_port = find_repeats_in_dict(newCols_port)
        #print('new Cols: ', newCols_port)
        df_port = df_port.drop(index=[0], axis=0)
        df_port.reset_index(drop=True, inplace=True)
        df_port = df_port.rename(columns=dict(zip(list(df_port), newCols_port)))
        df_port.replace('Не используется', '', inplace=True)
        df_port = df_port.loc[df_port['DNS-имя'] != '']
    except ValueError:
        df_port = pandas.DataFrame({'ЦОД':['UNKNOWN','UNKNOWN'],
                                    'Зал':['UNKNOWN','UNKNOWN'],
                                    'Ряд':['UNKNOWN','UNKNOWN'],
                                    'Стойка':['UNKNOWN','UNKNOWN'],
                                    'Unit':['UNKNOWN','UNKNOWN'],
                                    'Роль':['UNKNOWN','UNKNOWN'],
                                    'DNS-имя':['UNKNOWN','UNKNOWN'],
                                    'IP':['UNKNOWN','UNKNOWN'],
                                    'Маска':['UNKNOWN','UNKNOWN']})
    #print(df_port)
    return df_srv, df_sw, df_port

def convert_netmask(s):
    mask='UNKNOWN'
    maskDict={
        '255.255.255.255':'/32',
        '255.255.255.254':'/31',
        '255.255.255.252':'/30',
        '255.255.255.248':'/29',
        '255.255.255.240':'/28',
        '255.255.255.224':'/27',
        '255.255.255.192':'/26',
        '255.255.255.128':'/25',
        '255.255.255.0':'/24',
        '255.255.254.0':'/23',
        '255.255.252.0':'/22',
        '255.255.248.0':'/21',
        '255.255.240.0':'/20'
    }
    s=s.strip()
    if len(s)<4:
        mask=s
    else:
        mask=maskDict[s]
    return mask

def convert_gw_to_subnet(s):
    s=s.strip()
    netList=s.split('.')
    netList[-1]=str(int(netList[-1])-1)
    net=''
    for item in netList:
        net=net+item+'.'
    net=net[:-1]
    return net

def netpassport_to_json(df_srv, df_sw, df_port):
    segment_name = 'UNKNOWN'
    segment_data = 'UNKNOWN'
    segment_mgmt = 'UNKNOWN'
    vlan_data = 'UNKNOWN'
    vlan_mgmt = 'UNKNOWN'
    data_net = 'UNKNOWN'
    ipmi_net = 'UNKNOWN'
    vision_net = 'UNKNOWN'
    mgmt_net = 'UNKNOWN'
    print('preparing JSON for docxtpl for Table4:')
    json_netpassport_srv = {}
    dict_netpassport_srv = []
    for i in range(0, df_srv.shape[0]):
        dict_netpassport_srv.append({"cols": {'num': str(i+1), 'name': df_srv['DNS-имя1'].iloc[i], 'role': df_srv['Роль \\ Имя ВМ'].iloc[i],
                                              'ip_d': df_srv['IP1'].iloc[i], 'ip_m': df_srv['IP2'].iloc[i],'ip_y': df_srv['IP3'].iloc[i], 'date': df_srv['Год и порядковый номер поставки'].iloc[i]}})
    json_netpassport_srv = dict_netpassport_srv
    print('preparing JSON for docxtpl for Table5:')
    json_netpassport_sw = {}
    dict_netpassport_sw = []
    for i in range(0, df_sw.shape[0]):
        dict_netpassport_sw.append({"cols": {'num': str(i+1), 'name': df_sw['DNS-имя'].iloc[i], 'type': df_sw['Функционал'].iloc[i],
                                              'ip_m': df_sw['MGMT IP'].iloc[i], 'date': df_sw['Год реализации, номер поставки'].iloc[i]}})
    json_netpassport_sw = dict_netpassport_sw
    print('preparing JSON for docxtpl for Table6:')
    json_netpassport_port = {}
    dict_netpassport_port = []
    for i in range(0, df_port.shape[0]):
        dict_netpassport_port.append({"cols": {'dc': df_port['ЦОД'].iloc[i], 'rack': df_port['Стойка'].iloc[i], 'unit': df_port['Unit'].iloc[i], 'role': df_port['Роль'].iloc[i], 'name': df_port['DNS-имя'].iloc[i], 'port': '', 'ip': df_port['IP'].iloc[i], 'mask': df_port['Маска'].iloc[i]}})
    json_netpassport_port = dict_netpassport_port
    df_nets = df_srv.loc[df_srv['Роль \\ Имя ВМ'] == 'Management']
    #print(df_nets)
    segment_data = df_nets['Сегмент1'].iloc[0]
    if 'PROD' in segment_data:
        segment_name = 'PROD'
    elif 'TEST' in segment_data:
        segment_name = 'TEST'
    else:
        segment_name = 'UNKNOWN'
    segment_mgmt = df_nets['Сегмент2'].iloc[0]
    vlan_data = df_nets['VLAN1'].iloc[0]
    vlan_mgmt = df_nets['VLAN2'].iloc[0]
    data_net = convert_gw_to_subnet(df_nets['Шлюз1'].iloc[0])+convert_netmask(df_nets['Маска1'].iloc[0])
    ipmi_net = convert_gw_to_subnet(df_nets['Шлюз2'].iloc[0])+convert_netmask(df_nets['Маска2'].iloc[0])
    vision_net = convert_gw_to_subnet(df_nets['Шлюз3'].iloc[0])+convert_netmask(df_nets['Маска3'].iloc[0])
    mgmt_net = convert_gw_to_subnet(df_sw['GW for MGMT'].iloc[1])+convert_netmask(df_sw['MGMT IP MASK'].iloc[1])
    json_netpassport={
        'srv': json_netpassport_srv,
        'sw': dict_netpassport_sw,
        'port': json_netpassport_port,
        'vars':
            {
                'segment_name': segment_name,
                'segment_data': segment_data,
                'segment_mgmt': segment_mgmt,
                'vlan_data': vlan_data,
                'vlan_mgmt': vlan_mgmt,
                'data_net': data_net,
                'ipmi_net': ipmi_net,
                'vision_net': vision_net,
                'mgmt_net': mgmt_net
            }
    }
    return json_netpassport

def prepare_lsblk(path):
    print('preparing LSBLK: ',path)
    with open(path, 'r') as file:
        lines = file.readlines()
    #распил файла lsblk в объект
    json_lsblk = {}
    hostname=''
    for line in lines:
        if len(line)>0 and len(line)<17:
            hostname=line.strip()
            json_lsblk[hostname]=[]
        else:
            if 'NAME' in line:
                coordNAME = (re.search('NAME', line).start(), re.search('MAJ:MIN', line).start())
                coordMAJMIN = (re.search('MAJ:MIN', line).start(), re.search('MAJ:MIN', line).end())
                coordRM = (re.search('RM', line).start(), re.search('RM', line).end())
                coordSIZE = (re.search('RM', line).end(), re.search('SIZE', line).end())
                coordRO = (re.search('RO', line).start(), re.search('RO', line).end())
                coordTYPE = (re.search('TYPE', line).start(), re.search('MOUNTPOINT', line).start())
                coordMOUNTPOINT = (re.search('MOUNTPOINT', line).start(), len(line))
            splittedLine = [
                line[coordNAME[0]:coordNAME[1]].rstrip(),
                line[coordMAJMIN[0]:coordMAJMIN[1]].strip(),
                line[coordRM[0]:coordRM[1]].strip(),
                line[coordSIZE[0]:coordSIZE[1]].strip(),
                line[coordRO[0]:coordRO[1]].strip(),
                line[coordTYPE[0]:coordTYPE[1]].strip(),
                line[coordMOUNTPOINT[0]:len(line)-1].strip()
            ]
            json_lsblk[hostname].append(splittedLine)
    for host in list(json_lsblk.keys()):
        json_lsblk[host]=pandas.DataFrame(data=json_lsblk[host][1:],columns=json_lsblk[host][0])
    return json_lsblk

def lsblk_to_json(json_lsblk, json_netpassport):
    print('preparing JSON for docxtpl for Table8+:')
    listRoles= []
    json_disk = {}
    for i in range(len(json_netpassport['srv'])):
        if 'xk' not in json_netpassport['srv'][i]['cols']['name']:
            if (json_netpassport['srv'][i]['cols']['role'] != 'Management') and (json_netpassport['srv'][i]['cols']['role'] not in listRoles):
                listRoles.append(json_netpassport['srv'][i]['cols']['role'])
    dictRoles={}
    for item in listRoles:
        dictRoles[item]=[]
    for i in range(len(json_netpassport['srv'])):
        if ('xk' not in json_netpassport['srv'][i]['cols']['name']) and (json_netpassport['srv'][i]['cols']['role'] != 'Management'):
            dictRoles[json_netpassport['srv'][i]['cols']['role']].append(json_netpassport['srv'][i]['cols']['name'])
    for type in list(dictRoles.keys()):
        print('Заполнение структуры дисковой подсистемы для',type,'(',dictRoles[type][0],')')
        role = json_lsblk[dictRoles[type][0]]
        dict_disk = []
        for i in range(0, role.shape[0]):
            dict_disk.append({"cols": {
                                       'name': role['NAME'].iloc[i],
                                       'size': role['SIZE'].iloc[i],
                                       'type': role['TYPE'].iloc[i],
                                       'mount': role['MOUNTPOINT'].iloc[i]
                                       }})
        json_disk[type] = dict_disk
    return json_disk

def fstab_to_json(path):
    print('preparing FSTAB: ',path)
    with open(path, 'r') as file:
        lines = file.readlines()
    #распил файла lsblk в объект
    json_fstab = {}
    hostname=''
    for line in lines:
        if line.startswith('#') or (line == '\n'):
            pass
        else:
            if len(line)>14 and len(line)<17 and ('-' in line):
                hostname=line.strip()
                json_fstab[hostname]=[]
            else:
                line=line[:-1]
                fstablist = line.split(' ')
                fstablist = [x for x in fstablist if x]
                json_fstab[hostname].append(
                    {
                        'file system': fstablist[0],
                        'mount point': fstablist[1],
                        'type': fstablist[2],
                        'options': fstablist[3],
                        'dump': fstablist[4],
                        'pass': fstablist[5]
                    }
                )
    #for host in list(json_lsblk.keys()):
    #    json_lsblk[host]=pandas.DataFrame(data=json_lsblk[host][1:],columns=json_lsblk[host][0])
    return json_fstab

def find_roles_in_netpassport(json_netpassport):
    listRoles = []
    for i in range(len(json_netpassport['srv'])):
        if 'xk' not in json_netpassport['srv'][i]['cols']['name']:
            if json_netpassport['srv'][i]['cols']['role'] not in listRoles:
                listRoles.append(json_netpassport['srv'][i]['cols']['role'])
    dictRoles = {}
    for item in listRoles:
        dictRoles[item] = []
    for i in range(len(json_netpassport['srv'])):
        if 'xk' not in json_netpassport['srv'][i]['cols']['name']:
            dictRoles[json_netpassport['srv'][i]['cols']['role']].append(json_netpassport['srv'][i]['cols']['name'])
    print('Найдены следующие роли:', dictRoles)
    return dictRoles
def raid_to_json_full(RaidPath):
    print('preparing RAID file: ',RaidPath)
    df_raid = pandas.read_excel(RaidPath)
    df_raid = df_raid.fillna('')
    newCols = []

    #df_srv = df_srv.drop(index=[0, 1], axis=0)
    #df_srv.reset_index(drop=True, inplace=True)
    #df_srv = df_srv.rename(columns=dict(zip(list(df_srv), newCols_srv)))
    #df_srv.replace('Не используется', '', inplace=True)
    #df_srv = df_srv.loc[df_srv['DNS-имя1'] != '']
    df_raid = df_raid.loc[df_raid['№'] != 'Пример']
    #print(df_raid)
    #print(list(df_raid))
    json_raid_full = {}
    json_raid_full['item'] = []
    for i in range(0, df_raid.shape[0]):
        json_raid_full['item'].append(
            {'num':df_raid['№'].iloc[i],
             'date':df_raid['Год и порядковый номер поставки'].iloc[i],
             'name': df_raid['DNS имена'].iloc[i],
             'type': df_raid['Конфигурация сервера'].iloc[i],
             'spec': df_raid['Спецификация дисковой подсистемы сервера ПАК'].iloc[i],
             'raid': df_raid['Кол-во дисков \ Тип RAID \ Назначение группы (например - ОС)'].iloc[i],
             'mount': df_raid['RAID-группа \ Тип раздела \ Раздел, Гб (объем) \тип FS \ точка монтирования (по необходимости).\nOS-default-lvm - размеры системных разделов ОС (/boot, / , /var, /root, swap) определны lvm по умолчанию (требований нет).'].iloc[i],
             'fs': df_raid['ТипFS - Параметры FS'].iloc[i],
             'params': df_raid['Точка монтированияn - параметры '].iloc[i]
             }
        )
    #print(json_raid_full)
    return json_raid_full

def raid_to_json_short(json_raid_full,dictRoles):
    json_raid = {}
    dict_raid = []
    dictHosts = []
    for item in list(dictRoles.keys()):
        dictHosts.append(dictRoles[item][0])
    num=1
    for item in json_raid_full['item']:
        if item['name'] in dictHosts:
            dict_raid.append(
                {"cols": {'num': num, 'type': item['type'], 'spec': item['spec'],
                          'raid': item['raid'], 'mount': item['mount'], 'params': item['params']}})
            num+=1
    json_raid = dict_raid
    #print('json_raid:\n',json_raid)
    return json_raid

# Main script
if __name__ == '__main__':
    OStype=os.name
    if OStype == 'posix':
        slash_type = '/'
    else:
        slash_type = '\\'
    print('Тип ОС определен как:',OStype)
    userProps = read_props('properties.ini')
    print('Найдены настройки: ',userProps)
    # поиск документов
    docPath = doc_search(os.path.dirname(os.path.abspath(__file__)))

    # Загрузка шаблона
    if userProps['PAC_type'] == 'МБД.Г':
        doc = DocxTemplate("template_ADB.docx")
    elif userProps['PAC_type'] == 'МБД.Х':
        doc = DocxTemplate("template_ADH.docx")
    else:
        print('Шаблон не найден')
    # Обработка данных из источников
    if docPath['specification'] != 'UNKNOWN':
        df_spec = prepare_spec(docPath['specification'])
        json_spec = spec_to_json(df_spec)
    else:
        json_spec = [{
            "cols": {'spec': 'UNKNOWN', 'amount': 'UNKNOWN', 'comment': 'UNKNOWN', 'date': 'UNKNOWN', 'articul': 'UNKNOWN'}
        }]
    if docPath['vmlist'] != 'UNKNOWN':
        json_vm = vm_to_json(docPath['vmlist'],userProps['OS_version'])
    else:
        json_vm = [{
            "cols": {'name': 'UNKNOWN', 'OS': 'UNKNOWN', 'vCPU': 'UNKNOWN', 'RAM': 'UNKNOWN',
                     'space': 'UNKNOWN','diskType': 'UNKNOWN','comment': 'UNKNOWN'}
        }]
    if docPath['cablejournal'] != 'UNKNOWN':
        df_cj = prepare_cj(docPath['cablejournal'])
        json_cj = cj_to_json(df_cj)
    else:
        json_cj = [{
            "cols": {'num': 'UNKNOWN', 's_name': 'UNKNOWN', 's_port': 'UNKNOWN', 'd_name': 'UNKNOWN',
                     'd_port': 'UNKNOWN', 'speed': 'UNKNOWN', 'type': 'UNKNOWN'}
        }]
    if docPath['netpassport'] != 'UNKNOWN':
        df_netpassport=prepare_netpassport(docPath['netpassport'])
        json_netpassport = netpassport_to_json(df_netpassport[0],df_netpassport[1],df_netpassport[2])
    else:
        json_netpassport = {

            'srv': [{
                "cols": {'num': 'UNKNOWN', 'name': 'UNKNOWN', 'role': 'UNKNOWN', 'ip_d': 'UNKNOWN',
                         'ip_m': 'UNKNOWN', 'ip_v': 'UNKNOWN', 'date': 'UNKNOWN'}
            }],
            'sw': [{
                "cols": {'num': 'UNKNOWN', 'name': 'UNKNOWN', 'type': 'UNKNOWN',
                         'ip_m': 'UNKNOWN', 'date': 'UNKNOWN'}
            }],
            'port': [{
                "cols": {'dc': 'UNKNOWN', 'rack': 'UNKNOWN', 'unit': 'UNKNOWN', 'role': 'UNKNOWN',
                         'name': 'UNKNOWN', 'port': 'UNKNOWN', 'ip': 'UNKNOWN', 'mask': 'UNKNOWN'}
            }],
            'vars':
            {
                'segment_name': 'UNKNOWN',
                'segment_data': 'UNKNOWN',
                'segment_mgmt': 'UNKNOWN',
                'vlan_data': 'UNKNOWN',
                'vlan_mgmt': 'UNKNOWN',
                'data_net': 'UNKNOWN',
                'ipmi_net': 'UNKNOWN',
                'vision_net': 'UNKNOWN',
                'mgmt_net': 'UNKNOWN'
            }
        }
    if docPath['lsblk'] != 'UNKNOWN':
        json_lsblk = prepare_lsblk(docPath['lsblk'])
        json_disk = lsblk_to_json(json_lsblk, json_netpassport)
    else:
        json_disk = {
            'Segment': [{
                "cols": {'name': 'UNKNOWN', 'size': 'UNKNOWN', 'type': 'UNKNOWN',
                         'mount': 'UNKNOWN'}
            }],
            'Master': [{
                "cols": {'name': 'UNKNOWN', 'size': 'UNKNOWN', 'type': 'UNKNOWN',
                         'mount': 'UNKNOWN'}
            }],
            'ADCC': [{
                "cols": {'name': 'UNKNOWN', 'size': 'UNKNOWN', 'type': 'UNKNOWN',
                         'mount': 'UNKNOWN'}
            }]
        }
    if docPath['fstab'] != 'UNKNOWN':
        json_fstab = fstab_to_json(docPath['fstab'])
    else:
        print('ошибка fstab')
    if docPath['raid'] != 'UNKNOWN':
        json_raid_full = raid_to_json_full(docPath['raid'])
        dictRoles = find_roles_in_netpassport(json_netpassport)
        json_raid = raid_to_json_short(json_raid_full, dictRoles)
    else:
        json_raid = [{
            "cols": {'num': 'UNKNOWN', 'type': 'UNKNOWN', 'spec': 'UNKNOWN', 'raid': 'UNKNOWN',
                     'mount': 'UNKNOWN', 'params': 'UNKNOWN'}
        }]
    now = datetime.datetime.now()

    # Данные для заполнения шаблона
    print('Заполнение итогового JSON в нотации JINJA2:')
    context = {
        'project_name': userProps['PAC_name'],
        'pac_type': userProps['PAC_type'],
        'OS_version': userProps['OS_version'],
        'changes' : {
            'change_date': str(now.day) + '.' + str(now.month) + '.' + str(now.year),
            'change_author': userProps['author'],
            'change_description': 'Документ сгенерирован (v' + version + ')',
        },
        'tbl1_contents': json_spec,
        'tbl2_contents': json_vm,
        'tbl3_contents': json_cj,
        'tbl4_contents': json_netpassport['srv'],
        'tbl5_contents': json_netpassport['sw'],
        'tbl6_contents': json_netpassport['port'],
        'segment_name': json_netpassport['vars']['segment_name'],
        'segment_data': json_netpassport['vars']['segment_data'],
        'segment_mgmt': json_netpassport['vars']['segment_mgmt'],
        'vlan_data': json_netpassport['vars']['vlan_data'],
        'vlan_mgmt': json_netpassport['vars']['vlan_mgmt'],
        'data_net': json_netpassport['vars']['data_net'],
        'ipmi_net': json_netpassport['vars']['ipmi_net'],
        'vision_net': json_netpassport['vars']['vision_net'],
        'mgmt_net': json_netpassport['vars']['mgmt_net'],
        'tbl7_contents': json_raid,
        'tbl8_contents': json_disk['Segment'],
        'tbl9_contents': json_disk['Master'],
        'tbl10_contents': json_disk['ADCC']
    }

    # Заполнение шаблона данными
    print('Обработка данных docxtpl')
    print(context)
    doc.render(context)

    # Сохранение документа
    PAC_type=''
    if context['pac_type'] == 'МБД.Г':
        PAC_type = 'ADB'
    elif context['pac_type'] == 'МБД.Х':
        PAC_type = 'ADH'
    elif context['pac_type'] == 'МБД.КХ':
        PAC_type = 'ADQM'
    docName = 'Инсталляционная_карта_'+PAC_type+'_'+userProps['PAC_name']+'_v'+version
    docDir = os.path.dirname(os.path.abspath(__file__))+slash_type+'result'
    os.makedirs(docDir, exist_ok=True)
    resultPath = docDir+slash_type+docName+'.docx'
    print('Сохранение документа:',resultPath)
    doc.save(resultPath)
    print('FINISHED')


