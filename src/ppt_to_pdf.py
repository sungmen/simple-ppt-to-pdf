"""
 Purpose : Convert PowerPoint (PPT/PPTX) to Adobe PDF
 pip install -r requirements.txt
 Usage : python ppt_to_pdf.py
"""

import sys
import os.path
from glob import glob
import pathlib
import win32com.client
from multiprocessing import Pool

def convert_p(name):
    curpath = str((pathlib.Path(__file__).parent.absolute())).replace('\\', '\\\\')
    name = str(curpath) + '\\\\' + str(name)
    ext = os.path.splitext(name)
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(name, WithWindow = False)
    name = name.replace(ext[1], '.pdf')
    presentation.SaveAs(name, 32)
    powerpoint.Quit()
    presentation =  None
    powerpoint = None
    print(name + ' success convert ' + ext[1] +' to pdf ')

def convert(exten, pool):
    pool.map(convert_p, glob(str('*.')+exten))        
    
if __name__ == '__main__':
    pool = Pool(processes=4)
    p_list = ['ppt', 'pptx']
    for i in p_list:
        convert(i, pool)
