import os
import time

import pandas as pd

home = os.path.expanduser('~')
target = os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','Single shortage','test')



t= []
start_time = time.time()

for f in os.listdir(target):
    print(f)
    temp = pd.read_excel(os.path.join(target, f))
    t.append(temp)
    # print(print("%s seconds ---" % (time.time() - start_time)))

ans = pd.concat(t)

ans = ans[['Commodity','Single Shortage QTY','ODM','Series','HP PN','Prev_Single Shortage QTY','Procurement type','reportDate','ETA','GPS Remark','LastSGreportDate']]
ans.reset_index(drop = True)

print(ans)

ans.to_excel(os.path.join(target, 'concat.xlsx'))
