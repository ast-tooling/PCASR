import subprocess
import os

def winMergeOpen(path1, path2):
    programFiles = os.environ["ProgramFiles(x86)"]
    winMergeU = programFiles + "\WinMerge\WinMergeU.exe"
    path1 = path1
    path2 = path2
    openPath ='"%s" "%s" "%s"'%(winMergeU,path1,path2)
    subprocess.Popen(openPath)

