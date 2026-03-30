import threading, requests, pandas
from os import system
from time import sleep

N_SAMPLES = 60
LOCK_INFO = threading.Lock()
STOP_UPD  = threading.Event()
STOP_WR   = threading.Event()
URL       = 'https://sdcm.ru/%D0%BA%D0%BE%D0%BD%D1%82%D1%80%D0%BE%D0%BB%D1%8C/get.php?uid=2&bid=2&dt=0101-01-01'
STATIONS_INFO = {
    'Минск':           {'staId':'57', 'rcvId': '91'},
    'Иркутск':         {'staId':'86', 'rcvId':'172'},
    'Калининград':     {'staId':'26', 'rcvId': '30'},
    'Игарка':          {'staId':'72', 'rcvId':'137'},
    'Северо-Курильск': {'staId':'97', 'rcvId':'205'}
}

last_info, all_info = ['']*30, []


def myprint(txt:str, color:str, end:str=''):
    match color:
        case 'red':    color = '\033[31m'
        case 'yellow': color = '\033[33m'
        case 'green':  color = '\033[32m'
        case _:        color = '\033[00m'
    print(f'\r\033[K{color}{txt}\033[00m', end=end)

def isMyStation(info:list):
    for name, ids in STATIONS_INFO.items():
        if info['stationId']  != ids['staId']: continue
        if info['receiverId'] != ids['rcvId']: continue
        return True, name
    return False, None

def reader():
    global last_info, all_info
    myprint('Собираю данные', 'yellow', '\n')
    myprint('Ожидание обновления', 'yellow')
    while any(item == '' for item in last_info): sleep(1)
    while True:
        with LOCK_INFO:
            all_info.append(last_info.copy())
        if len(all_info) >= N_SAMPLES: break
        for s in range(30, 0, -1):
            STOP_WR.wait()
            t = (N_SAMPLES-1 - len(all_info))*30 + s
            myprint(f'Осталось {t//60:02d}:{t%60:02d}', 'yellow')
            sleep(1)
    STOP_UPD.set()
    myprint(f'Почти готово...', 'yellow')

def updater():
    global last_info
    while not STOP_UPD.is_set():
        try:
            response = requests.get(URL, timeout=10)
            response.raise_for_status()
            data = response.json()
            info = {name: [] for name in STATIONS_INFO} # {k:[], ...}
            for station in data['corrections']:
                ok, name = isMyStation(station)
                if ok: info[name].append(station)  # {k:[{k:'', ...}, ...], ...}
            info = { # {k:{k:'', ...}, ...}
                k: max(v, key=lambda x: x['gpsTime']) if v else {}
                for k, v in info.items()
            }
            info = [ # [['', ...], ...]
                [
                    v.get('gps_single_plane',  ''),
                    v.get('gps_single_height', ''),
                    v.get('gps_sbas_height',   ''),
                    v.get('gps_sbas_nsta',     ''),
                    v.get('gps_sbas_hdop',     ''),
                    v.get('gps_sbas_vdop',     '')
                ]
                for v in info.values()
            ]
            info = [v for sub in info for v in sub] # ['', ...]
            with LOCK_INFO:
                last_info = [
                    v or last_info[i]
                    for i, v in enumerate(info)
                ]
            if not STOP_WR.is_set(): STOP_WR.set()
            if not STOP_UPD.is_set(): sleep(1)
        except requests.exceptions.Timeout:
            myprint('Превышено время ожидания. Проверьте подключение (российский VPN)', 'red')
            STOP_WR.clear()
        except requests.exceptions.HTTPError as e:
            myprint(f'Ошибка сервера (код {response.status_code}): {e}', 'red')
            STOP_WR.clear()
        except Exception as e:
            myprint(f'Ошибка: {e}', 'red')
            STOP_WR.clear()

system('cls')

t1 = threading.Thread(target=reader)
t2 = threading.Thread(target=updater)

t1.start(); t2.start()
t1.join();  t2.join()

myprint('Данные получены', 'green', '\n')
myprint('Записываю в таблицу лаба.xlsx', 'yellow', '\n')

all_info = [[float(n) for n in line] for line in all_info]

header = []
for point in STATIONS_INFO:
    header += [
        (point, 'Δr'),
        (point, 'Δh, абс.'),
        (point, 'Δh, SBAS'),
        (point, 'N'),
        (point, 'HDOP'),
        (point, 'VDOP'),
    ]

while True:
    try:
        with pandas.ExcelWriter('лаба.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            pandas.DataFrame(all_info, columns=pandas.MultiIndex.from_tuples(header)).to_excel(writer, sheet_name='Данные')
        break
    except PermissionError:
        myprint('Закройте файл таблицы!', 'red')
        sleep(1)
    
myprint('ГОТОВО', 'green', '\n')
