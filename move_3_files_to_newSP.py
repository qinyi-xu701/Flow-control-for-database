import os
import shutil

home = os.path.expanduser("~")
OF = os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','Upload folder ( for buyer update )')
TF = os.path.join(home, 'HP Inc','GPS TW Innovation - Documents','Project team','Project RiXin - Shortage management','Upload_folder')
datelist = [str(i + 20230331) for i in range(31)]
datelist = [str(20230711)]
commodity = '_'
# date = '20230510'

def move(originFolder, targetFolder):
    Original_FD_folder = os.path.join(originFolder, "FD_Archive_After_1025")
    target_FD_folder = os.path.join(targetFolder, 'FD_Archive')

    for f in os.listdir(Original_FD_folder):
        try:
            shutil.move(os.path.join(Original_FD_folder, f), os.path.join(target_FD_folder, f))
            print(f)
        except:
            pass

    Original_shortage_folder = os.path.join(originFolder ,"shortage_Archive_After_1025")
    target_shortage_folder = os.path.join(originFolder ,"Shortage_Archive")

    for f in os.listdir(Original_shortage_folder):
        try:
            shutil.move(os.path.join(Original_shortage_folder, f), os.path.join(target_shortage_folder, f))
            print(f)
        except:
            pass

    Original_PNbasedDetail_folder = os.path.join(originFolder ,"PNbasedDetail_Archive_After_1025")
    target_PNbasedDetail_archive_folder = os.path.join(originFolder ,"PNbasedDetail_Archive")

    for f in os.listdir(Original_PNbasedDetail_folder):
        try:
            shutil.move(os.path.join(Original_PNbasedDetail_folder, f), os.path.join(target_PNbasedDetail_archive_folder, f))
            print(f)
        except:
            pass
    
    return

# for _ in datelist:
#     move(originFolder, _)

move(OF, TF)