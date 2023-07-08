import subprocess

from memory_profiler import profile


def my_function():
    subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'C:/Users/lulo/OneDrive - HP Inc/SystemPy/autoSendPNFV.py'], check=True)
    # subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/concat.py'], check=True)
    # subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/Upload_to_SQL.py'], check=True)
    # subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/SGTransform.py'], check=True)
    # subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/upload_PNFV_toSQL.py'], check=True)
    # subprocess.run(['C:/Users/lulo/AppData/Local/Programs/Python/Python310/python.exe', 'c:/Users/lulo/GitHub/Flow-control-for-database/amend_data.py'], check=True)

if __name__ == "__main__":
    my_function()

# memory_profiler.run(my_function)
# print(memory_profiler.memory_usage())