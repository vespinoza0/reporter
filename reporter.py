import csv
import tkinter
from tkinter import filedialog
import re
import os
import email 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import config
import datetime

def connectAgain():
	server = smtplib.SMTP('smtp.gmail.com:587')
	server.ehlo()
	server.starttls()
	server.login(config.EMAIL_ADDRESS, config.PASSWORD)
		

	
def attachSend(dest,file,sentrow, mycol,emailSub):

	try:
		print("trying to send now to", dest)
		print("attachment is ",file)
		server = smtplib.SMTP('smtp.gmail.com:587')
		server.ehlo()
		server.starttls()
		server.login(config.EMAIL_ADDRESS, config.PASSWORD)
		#print("logged in!")
		msg = MIMEMultipart()
		msg['From'] = config.EMAIL_ADDRESS
		msg['To'] = dest
		msg['Subject'] = emailSub
		print(config.EMAIL_ADDRESS)
		#print("the email subject is",emailsub)
		
		mymsg= "Dear student,\n\nPlease find attached submission report for Hacker Rank assignment.\nWe understand that you now have access to the internal notes for this assignment, but please note that this may not contain the correct answer as we continue to debug each module. \n\nThank you,\nThe codehelp team "
		part1 = MIMEText(mymsg, 'plain')
		filename = file
		fp = open(filename, 'rb')
		attach = MIMEApplication(fp.read(), 'pdf')
		#print("attached a PDF")
		fp.close()
		attach.add_header('Content-Disposition', 'attachment', filename = filename)
		msg.attach(attach)
		msg.attach(part1)
		###############################################
		print("sendmail now!")
		server.sendmail(config.EMAIL_ADDRESS, dest, msg.as_string())
		#print("mail was sent!")
		sentMaster[sentrow][mycol] = 1
		print("success: Email sent to " ,dest,"!")
		server.quit()
		
	except:
		newSentNotSent(sentMaster)
		print("Error: Email was not sent to ", dest)
		exit()
		
		
def newSentNotSent(nc):
	now = format(datetime.datetime.now())
	nows = str(now)
	nows1 = nows.replace(":","_")
	newFN = nows1 +"_sentNotSent.csv"
	with open(newFN, "w") as output:
		writer = csv.writer(output, lineterminator='\n')
		writer.writerows(nc)
	print("##################################################################################################################")
	print("Successfully exported ",newFN)

	
def writeErrorLog(nomatch):
	now = format(datetime.datetime.now())
	nows = str(now)
	nows1 = nows.replace(":","_")

	noMatchName = nows1+ tail +'_unMatchedEmail.csv' 
	with open(noMatchName, "w") as output:
		writer = csv.writer(output, lineterminator='\n')
		writer.writerows(nomatch)
	print("Error Log file created for" , tail,"! Saved as ", noMatchName)

	
def getCol(hrtail):
	hrtail = hrtail.lower()
	print("folder name is ", hrtail)
	tail1 = hrtail.split("_")
	assignmentName = tail1[1:]
	print("assignment name is ", assignmentName)
	col = 1
	HRname = hrtail.lower()
	h = HRname.split('_')
	hh = h[0:len(h)-2]
	
	for i in range(0,len(Canheader)):
		thing = Canheader[i]
		thing = thing.lower()
		canName = thing.split()
		cName = canName[0:len(h)-1]
		if cName== assignmentName:
			return i
		if i == len(Canheader)-1:
			print("no column in canvas was found associated with this assignment!")
			col = int(input("Enter the column to edit in CANVAS, refer to CanvasColumn.xlsx file in TA drive: "))
			return col+3 

######## NEED TO LEARN HOW TO MAKE A MAIN FUNCTION EVENTUALLY ###########################

myTAlist = config.TAlist
root = tkinter.Tk()
root.withdraw()
Ca = filedialog.askopenfilename(title = "Select SentNotSent file",filetypes = (("CSV files","*.csv"),("all files","*.*")))  # get directory +filename.csv

with open(Ca) as csvDataFile:
	csvReader = csv.reader(csvDataFile)
	sentMaster = list(csv.reader(csvDataFile))

Canheader = sentMaster[0]

TUlistMaster = []

for row in range(0,len(sentMaster)):
	ting = sentMaster[row][2]
	TUlistMaster.append(ting)


root.directory = filedialog.askdirectory(title = "Select folder containing all PDFS for one assignment")
print("you have selected folder",root.directory)
arr = os.listdir(root.directory)
print(root.directory)
head, tail = os.path.split(root.directory)
print("the tail is ", tail)
mycol = getCol(tail)
print("my column is", mycol)
print(Canheader[mycol])
tdel = tail.split('_')
emailSub =' '.join(tdel)
print("email subject is ", emailSub)



regex_txt = r"^[a-zA-Z]{3}[0-9]{5}$" # what you're looking for
pattern = re.compile(regex_txt, re.IGNORECASE) # pattern which ignores case so TUGxxxxx is permissable
TUemail = []  #Dupe catcher
noTU = []
trr = ["email","pdfFileName"]
noTU.append(trr)
personalMail = []
emailsSent = 0 

#for i in range(0,len(filez)):
for i in range(0,len(arr)):
	fn = arr[i]
	fileDir = root.directory+"/"+fn
	pdfFile =fn
	thing = pdfFile.split("_")
	emailTemp = thing[5]
	emailTemp2 = thing[5:7]
	crap ='@'.join(emailTemp2)
	
	for j in range(0,len(thing)):
		if thing[j] in TUemail or thing[j] in myTAlist:
			continue
		elif pattern.search(thing[j]):
			TUemail.append(thing[j])
			temps = str(thing[j])
			temps = temps.lower()
			if temps not in TUlistMaster:
				print(temps,"is not in tuListMaster")
				continue
			sentrow = TUlistMaster.index(temps)
			#print(temps ," is in row", sentrow)
			check = sentMaster[sentrow][mycol]
			if check == '0':
				email = temps +"@temple.edu"
				#print("send it to ", email,"!!")
				attachSend(email,fileDir, sentrow, mycol,emailSub)  # try this!!!
				emailsSent = emailsSent+1
			
		elif j == len(thing)-1:
			if emailTemp not in TUemail and emailTemp not in myTAlist:
				print("---------------------------------------------------")
				print("try to find this email in config login IDs", crap)
				# if "temple" not in crap:
					# print("this is a personal email: ", crap)
					# personalMail.append(crap)
					# noTU_fn = [crap,fileDir]
					# noTU.append(noTU_fn)
					# continue
				if crap == config.loginID1:
					 ogTUemail = config.login1
					 a,b =ogTUemail.split('@')
					 print("found the actual email!", ogTUemail)
					 sentrow = TUlistMaster.index(a)
					 check = sentMaster[sentrow][mycol]
					 if check == '0':
						 print("try sending it!")
						 attachSend(ogTUemail,fileDir,sentrow,mycol,emailSub)  # 
						 emailsSent = emailsSent+1
				elif crap == config.loginID2:
					 ogTUemail = config.login2
					 a,b =ogTUemail.split('@')
					 print("found the actual email!", ogTUemail)
					 sentrow = TUlistMaster.index(a)
					 check = sentMaster[sentrow][mycol]
					 if check == '0':
						 print("try sending it!")
						 attachSend(ogTUemail,fileDir,sentrow,mycol,emailSub)  #
						 emailsSent = emailsSent+1
				
				elif crap == config.loginID3:
					 ogTUemail = config.login3
					 a,b =ogTUemail.split('@')
					 print("found the actual email!", ogTUemail)
					 sentrow = TUlistMaster.index(a)
					 check = sentMaster[sentrow][mycol]
					 if check == '0':
						 print("try sending it!")
						 attachSend(ogTUemail,fileDir,sentrow,mycol,emailSub)  #
						 emailsSent = emailsSent+1
				elif crap == config.loginID4:
					 ogTUemail = config.login4
					 a,b =ogTUemail.split('@')
					 print("found the actual email!", ogTUemail)
					 sentrow = TUlistMaster.index(a)
					 check = sentMaster[sentrow][mycol]
					 if check == '0':
						 print("try sending it!")
						 attachSend(ogTUemail,fileDir,sentrow,mycol,emailSub)  #
						 emailsSent = emailsSent+1
				elif crap == config.loginID5:
					 ogTUemail = config.login5
					 a,b =ogTUemail.split('@')
					 print("found the actual email!", ogTUemail)
					 sentrow = TUlistMaster.index(a)
					 check = sentMaster[sentrow][mycol]
					 if check == '0':
						 print("try sending it!")
						 attachSend(ogTUemail,fileDir,sentrow,mycol,emailSub)  #
						 emailsSent = emailsSent+1
				elif crap == config.loginID6:
					 ogTUemail = config.login6
					 a,b =ogTUemail.split('@')
					 print("found the actual email!", ogTUemail)
					 sentrow = TUlistMaster.index(a)
					 check = sentMaster[sentrow][mycol]
					 if check == '0':
						 print("try sending it!")
						 attachSend(ogTUemail,fileDir,sentrow,mycol,emailSub)  #
						 emailsSent = emailsSent+1
				else:
					print("idk what this is ", emailTemp)
					noTU_fn = [emailTemp,fn]
					noTU.append(noTU_fn)
					

###########################################################################################################################
#Log out email acount					 
# server.quit()
#########################################################################################################################################
print("emails sent: ",emailsSent)
newSentNotSent(sentMaster)
writeErrorLog(noTU)

# print("######################################################################")
# print("TUemail list is ", len(TUemail), " and ", len(TUemail[0]))
# print(TUemail[0])
# print(TUemail[1])
# print("######################################################################")
# print("NoTUemail list is ", len(noTU), " and ", len(noTU[0]))
# print(noTU)


		

		

	
	
