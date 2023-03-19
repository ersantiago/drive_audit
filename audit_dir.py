#!/home/xinyx/miniconda3/bin/python3
import os
import sys
import time
from os import stat
from pwd import getpwuid
from filereader import *

path_conf = '/home/emerson/flagged.list'

list_conf = open(path_conf,'r').read().splitlines()

def find_owner(filename):
    return getpwuid(stat(filename).st_uid).pw_name

dir = sys.argv[1]
print(dir)
os.chdir(dir)
cmd = 'find . * | grep -v "^./"'
listF = os.popen(cmd).read().splitlines()
del listF[0]

exclude = ['kde4', 'cache', '.config', '.local', '20120301_GF28DST']
listF_filt = [s for s in listF if not any(xs in s for xs in exclude)]

output_xlsx = "audit.xlsx"
output_wkb = openpyxl.Workbook()
outwks = output_wkb['Sheet']

headlst = ['Full_path', 'Filename', 'Owner', 'FileType', 'Timestamp', 'Check', 'Flags']
for i in range(len(headlst)):
    outwks.cell(1,i+1).value = headlst[i]
output_wkb.save(output_xlsx)

header = '^'.join(['Full_path', 'Filename', 'Owner', 'FileType', 'Timestamp', 'Check'])

for k in range(len(listF_filt)):
    file = listF_filt[k]
    print(file)
    path = os.path.abspath(file)
    basename = os.path.basename(file)



    try:
        fnsplt = path.split('/')[-1].split('.')

        if os.path.isdir(path):
            ftype = 'directory'

            try:
                user = getpwuid(stat(file).st_uid).pw_name
            except:
                user = str(stat(file).st_uid)
            stamp = time.strftime('%x %X', time.gmtime(os.path.getmtime(file)))

        elif os.path.islink(path):
            ftype = 'link'
            user  = 'na'

        else:

            try:
                user = getpwuid(stat(file).st_uid).pw_name
            except:
                user = str(stat(file).st_uid)
            stamp = time.strftime('%x %X', time.gmtime(os.path.getmtime(file)))

            if len(fnsplt) == 1:
                ftype = 'no_ext'
            else:
                ftype = fnsplt[-1]
                if ftype.isalnum():
                    ftype = ftype
                else:
                    ftype = 'invalid'

                if ftype == '0':
                    ftype = '\'0'

    except:
        ftype = 'no_ext'

    path_ext = '/home/emerson/audit_lists/ext.list'
    path_odf = '/home/emerson/audit_lists/odf.list'
    path_gen = '/home/emerson/audit_lists/gen.list'
    path_cnf = '/home/emerson/audit_lists/cnf.list'

    list_ext    = open(path_ext,'r').read().splitlines()
    odf_ext     = open(path_odf,'r').read().splitlines()
    generic_ext = open(path_gen,'r').read().splitlines()
    supext      = list_ext + odf_ext + generic_ext
    conf_ext    = open(path_cnf,'r').read().splitlines()

    if ftype in supext:

        # check file-size if > 5mb
        if os.path.getsize(path) > 5000000:
            flagchk = 'exceeded filesize limit (5mb)'
            flags = 'not checked'
        else:
            try:

                if ftype == 'docx':
                    readfile = read_docx(path)
                elif ftype == 'pptx':
                    readfile = read_pptx(path)
                elif ftype == 'xlsx':
                    readfile = read_xlsx(path)
                elif ftype == 'xls':
                    readfile = read_xls(path)
                elif ftype == 'pdf':
                    readfile = read_pdf(path)
                elif ftype in odf_ext:
                    readfile = read_odf(path)
                elif ftype in generic_ext:
                    readfile = read_generic(path)
                else:
                    readfile = 'invalid'

                flags = check_flag(readfile, list_conf)

                if flags == '':
                    flagchk = 'Pass'
                else:
                    flagchk = 'Fail'
            except:
                flagchk = 'unable to read file, possible corrupted.'
                flags = 'not-checked'

    elif ftype in conf_ext:
        flagchk = 'Fail'
        flags = 'Project/Technology-specific Files'
    elif ftype == 'directory':
        flagchk = 'NA'
        flags = 'NA'

    else:
        flagchk = 'not-checked'
        flags = 'not-checked'

    freplst = [path, basename, user, ftype, stamp, flagchk, flags]
    

    row_index = k + 2
    for j in range(len(freplst)):
       outwks.cell(row_index,j+1).value = freplst[j]

    freport = '|'.join([path, basename, user, ftype, stamp, flagchk, flags])

output_wkb.save(output_xlsx)
output_wkb.close()

