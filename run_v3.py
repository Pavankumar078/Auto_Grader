import os
import subprocess
import time
import sys
import shutil
from openpyxl import Workbook
import signal
import difflib


def compare_files_ignore_spelling_and_case(file1_path, file2_path):
    with open(file1_path, 'r') as file1, open(file2_path, 'r') as file2:
        content1 = file1.read().lower()
        content2 = file2.read().lower()

    # Tokenize the contents into words
    words1 = content1.split()
    words2 = content2.split()

    # Initialize a SequenceMatcher object with the two sequences
    matcher = difflib.SequenceMatcher(None, words1, words2)

    # Get the similarity ratio between the sequences
    similarity_ratio = matcher.ratio()
    return similarity_ratio

    




# Create a new workbook
workbook = Workbook()

# Select the active sheet
sheet = workbook.active



dir_path = sys.argv[1]  # directory path of source file
input_path = sys.argv[2]  # directory path of input file
output_path = sys.argv[3]  # directory path of output file


c_files = []
input_files = []
output_files = []
std = []

# storing the input files in input_files list
input_files = [f for f in os.listdir(input_path)]

# storing the input files in output_files list
output_files = [f for f in os.listdir(output_path)]


input_files.sort()  # sort input_list
print("input files:", input_files)
head = tuple()
head += ("Roll_Number",)

output_files.sort()  # sort input_list
print("output files:", output_files)

for o in output_files:
	col = o.split('.')
	head += (col[0],)
head += ('Total',)
head += ('Comments',)
std.append(head)

# Recursively search the directory and directory for C files
i = 0

for root, dirs, files in os.walk(dir_path):
	for file in files:
		if file.endswith('.c'):
			c_files.append(file)  # appending list with only c files

c_files.sort()
root = dir_path

for c_file in c_files:
	os.chdir(root)
	# changing the directory to source directory	
	print('\n')
	print('------------------------------------------------------------------------')
	prg = root+c_file
	cmd = 'gcc '+prg+' -lm'+' -o example'
	#print(cmd)   
	n = os.system(cmd)  # compiling the program
	test_case = []
	k = 0
	d = c_file.split(".")
	tp = tuple()
	tp += (d[0],)
	total = 0
	pn=d[0]
	remarks=''
	time_exceed=''
	timeout_seconds = 10
	#print("root: ",root)
	if n == 0:  # this block will executed only of compilation is error free
		j = 0

		d[0] = d[0]+'_output'
		for fn in os.listdir(root):
		  if fn.endswith('.txt'):
		    file_path = os.path.join(root, fn)
		    os.remove(file_path)
		
		if os.path.exists(d[0]) and os.path.isdir(d[0]):  # Remove the folder if exist already
				shutil.rmtree(d[0]) # creating the directory of student roll number to store his output files
		os.mkdir(d[0])
		segments=[]
		mismatch=[]
		infinite=[]
		segment=' '
		for input_file in input_files:			
			cmd = ['./example' ,input_path+input_file]			
			try:
				#process = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
				#stdout, stderr = process.communicate()
				process = subprocess.run(cmd, timeout=5)
				if process.returncode != 0:
					segments.append(input_file)
					tp += (0,)
					j = j+1
					continue		  
				else:
					
					file_path = root+'output.txt'
					
					if os.path.exists(file_path):
						new_name = "output_"+str(j+1)+".txt"
						os.rename("output.txt", new_name)
						shutil.move(new_name, d[0])
					
						# path of the student output files
						stud_file_path = root+d[0]+'/'+new_name
					
						# path of the expected output files
						exptd_file_path = output_path+output_files[j]
					
						# Comparing the student output file with expected output files
						ratio = compare_files_ignore_spelling_and_case(stud_file_path, exptd_file_path)
						marks=ratio*1
						tp += (marks,)
						total += marks
						j = j+1
						k = k+1
						if(marks!=1): mismatch.append(new_name)
						
						#if((total*6/len(output_files))!=6): remarks='Output files are partially matched'
						#else: remarks=''
						#print(tp)
						
					else:
						continue
			except subprocess.TimeoutExpired:				
				#tp=tuple()
				#tp += (pn,)
				#for o in output_files: tp += (0,)
				tp += (0,)
				j = j+1
				total+=0
				print(input_file)
				infinite.append(input_file)
				#remarks='Timeout: C program took too long to complete.'
				continue
				
		print('\n')
		
		if len(segments) !=0: segment='Segmentation fault (core dumped) or Aborted (core dumped) while passing '+str(segments)+' as command line argument'
		if len(mismatch)!=0: remarks='Output files '+str(mismatch)+' are partially matched'
		if len(infinite)!=0: time_exceed='Timeout: program enters into infinite loop while passing '+str(infinite)+' as command line argument'
		
		if(len(segments) !=0 and len(mismatch)!=0 and len(infinite)!=0): comments= segment+', '+time_exceed+' and '+remarks
		elif(len(segments) !=0  and len(infinite)!=0): comments= segment+' and '+time_exceed
		elif(len(segments) !=0  and len(mismatch)!=0): comments= segment+' and '+remarks
		elif(len(infinite) !=0  and len(mismatch)!=0): comments= time_exceed+' and '+remarks
		elif(len(segments) !=0): comments= segment
		elif(len(infinite)!=0): comments=time_exceed
		elif(len(mismatch)!=0):comments=remarks
		else: comments=''
		
		
		
		
		if(total>2): total=round(total*6/len(output_files),2)
		else: total=2
		
		
			
		tp+=(total,comments,)		
		std.append(tp)
		
		print(str(c_file)+" is executed suceessfully"+"\nOutput files are saved in "+d[0]+" directory")
		
		
	else:
		print("\n"+c_files[i]+" has compilation error")	#this block will executed only if compilation has error
		remarks='compilation error'
		for o in output_files: tp += (0,)
		total+=2
		tp += (total,remarks,)
		std.append(tp)
	i=i+1
	print('------------------------------------------------------------------------')
	

#saving to excel sheet
for s in std:
	#print(s)
	sheet.append(s)
workbook.save("Marks.xlsx")
#END
