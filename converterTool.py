import sys
import converter
import os
import os
import glob
import datetime
import random
import shutil

input_path = sys.argv[1]
output_path = sys.argv[2]
error_dir = sys.argv[3]
archival_dir = sys.argv[4]
env = sys.argv[5]

#find the newst file in the input directory
files = glob.glob(input_path + '\\*.[Xx][Ll][Ss][Xx]')
if len(files) == 0:
        print('Nothing to do here, no files')
        exit(0) 
input_path = max(files,key=os.path.getctime)
td = datetime.datetime.today()
output_filename = 'filename' + '.xml'

#raises if file is open
try: 
        input_path_temp = input_path + 'temp'
        os.rename(input_path, input_path_temp)
        os.rename(input_path_temp, input_path)
except OSError:
        print("Input file is open: " + input_path)
        print("Can't move input file, please instruct to close file")

        raise
###
xl_filename = output_filename[:-4] + '.xlsx'
xml_filepath = output_path+ '\\' + output_filename
try:
        with open(xml_filepath,'wb') as f:
                        
                contents = converter.convert_to_xml(input_path)
                f.write(contents)

                print('success! Moving file to :' + xml_filepath)
                shutil.copyfile(input_path,archival_dir + '\\' + xl_filename)
except:
        print('error when converting excel to xml')
        shutil.copyfile(input_path, error_dir + '\\' + xl_filename)
        print("file moved to error folder:" + error_dir)
        raise
finally:
        os.remove(input_path)
        
	
# Error handling: I am assuming that exceptions raised by any part of the code
