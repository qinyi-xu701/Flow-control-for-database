import subprocess

try:
    subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/Download_external_3files.py'], check=True)
except:
    print("Error happen on download process!")
    input()
    exit()

try:
    subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/concat.py'], check=True)
except:
    print("Error happen on concat process!")
    input()
    exit()

try:
    subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/Upload_to_SQL.py'], check=True)
except:
    print("Error happen on uploadtoSQL process!")
    input()
    exit()

try:
    subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/SGTransform.py'], check=True)
except:
    print("Error happen on SG process!")
    input()
    exit()

try:
    subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/upload_PNFV_toSQL.py'], check=True)
except:
    print("Error happen on PNFV upload process!")
    input()
    exit()

try:
    subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/amend_data.py'], check=True)
except:
    print("Error happen on amend data process!")
    input()
    exit()