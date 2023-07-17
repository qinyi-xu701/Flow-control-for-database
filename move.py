import os
import shutil

home = os.path.expanduser("~")
targetFolder = os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','Upload folder ( for buyer update )')
datelist = [str(i + 20230331) for i in range(31)]
datelist = [str(20230711)]
commodity = '_'
# date = '20230510'

def move(targetFolder, date):
    FD_folder = os.path.join(targetFolder, "FD_today")
    FD_archive_folder = os.path.join(targetFolder, 'FD_Archive_After_1025')

    for f in os.listdir(FD_archive_folder):
        if f.startswith(date) and (commodity in f):
            shutil.move(os.path.join(FD_archive_folder, f), os.path.join(FD_folder, f))
        else:
            pass
        
    shortage_folder = os.path.join(targetFolder ,"shortage_today")
    shortage_archive_folder = os.path.join(targetFolder ,"Shortage_Archive_After_1025")

    for f in os.listdir(shortage_archive_folder):
        if f.startswith(date) and (commodity in f):
            shutil.move(os.path.join(shortage_archive_folder, f), os.path.join(shortage_folder, f))
        else:
            pass

    PNbasedDetail_folder = os.path.join(targetFolder ,"PNbasedDetail_today")
    PNbasedDetail_archive_folder = os.path.join(targetFolder ,"PNbasedDetail_Archive_After_1025")

    for f in os.listdir(PNbasedDetail_archive_folder):
        if f.startswith(date) and (commodity in f):
            shutil.move(os.path.join(PNbasedDetail_archive_folder, f), os.path.join(PNbasedDetail_folder, f))
        else:
            pass
    
    return

for _ in datelist:
    move(targetFolder, _)