# Python script template to grid any number of files.
# 
# All files of the fileformat -> file_format in 'inpat' will be gridded by sufer.
# The output files have the same name as the input files and are placed in the
# same directory.
#
import win32com.client

# input information
inpath = "C:\\Program Files\\Golden Software\\Surfer 13\\Samples\\"
file_format = ".dat"

# get a sufer instance
app = win32com.client.Dispatch("Surfer.Application")
app.Visible = True

fail_list = []

for infile in glob.glob(inpath + "*" + file_format):
    outfile = infile.replace(file_format, ".grd")
    try:
        app.GridData(DataFile=infile, Algorithm="srfKriging", ShowReport=False, OutGrid=outfile)
        pass
    except:
        fail_list.append(infile)

app.quit

if len(fail_list) == 0:
    print("### all files gridded ###")
else: 
    for fail in fail_list:
        outstr = fail + "\n"
    print("### most files gridded, following files failed: \n" + outstr)
