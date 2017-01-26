import os
import datetime
import shutil

dnow = datetime.datetime.now()
n=0
for root, subFolders, files in os.walk('/edi'):
    for file in files:
        dmodify = datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(root, file)))
        days_diff = (dnow-dmodify).days
        if days_diff > 270:
            n+=1
            # directory = str(os.path.join('/tmp/1'+root[4:]))
            # if not os.path.exists(directory):
            #         os.makedirs(directory)
            # shutil.copy(os.path.join(root, file), str(os.path.join('/tmp/1'+root[4:])+str(file)))
print n