###############################
############################### CSV to GOOGLE SLIDE PROGRAM
############################### Version 2.1
############################### Consists of CustomTkinter Window and File Selector
############################### Uses the google API to load, copy, edit and save a google slides file using given CSV data
############################### def EditSlides is responsible for the main editing!
###############################
############################### NOTES!!!: The file selector only looks in Drive currently!
###############################

from __future__ import print_function
import os.path,csv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
from threading import Thread
from tkinter import *
import customtkinter

class CSVtoGSlideRipple:#USES GOOGLE API TO EDIT GOOGLE SLIDE FROM CSV DATA
    def __init__(self,CSVFile,templateId,outputFile,outputFolderId):
        self.CSVFile=CSVFile
        self.templateId=templateId
        self.outputFile=outputFile
        self.outputFolderId=outputFolderId
        if not os.path.isfile(self.CSVFile):#Check if csv exists
            printz("ERROR: CSV does not exist")
            return
        self.CREDS=CREDS
        self.TOKEN=TOKEN
        self.SCOPES=SCOPES
        self.connectAPI()
        self.editSlides()

    def connectAPI(self):#CONNECTS TO THE GOOGLE API AND CREATES A TOKEN
        creds = None
        # The file token.json stores the user's access and refresh tokens, created when user logs in
        if os.path.exists(self.TOKEN):
            creds = Credentials.from_authorized_user_file(self.TOKEN, self.SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.CREDS, self.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(self.TOKEN, 'w') as token:
                token.write(creds.to_json())
        try:#Try to connect to drive and slides API
            self.DRIVE= build('drive', 'v3', credentials=creds)
            self.SLIDES = build('slides', 'v1', credentials=creds)
        except:#If cannot connect then output the error
            printz("Google Token Expired! Login Again")


    def readCSV(self):#FORMATS THE CSV DATA WHEN WE KNOW THE TEMPLATE TABLE SIZES
        #Load the CSV file Here
        with open(self.CSVFile) as csv_file:
            csvReader = csv.reader(csv_file, delimiter=',')
            data=[]
            for row in csvReader:
                data.append(row)#Only take correct number of data columns
        csv_file.close()
        tableList=[]#Holds all CSV collected tables
        typeList=[]
        #FIND START COL!!!Using the date
        startCol=0
        for i in range(len(data[5])):
            if data[5][i].count("/")==2:
                startCol=i
                printz("Starting at column "+str(startCol))
        #Add each data row to an array
        for i in range(1,len(data)):
            if data[i][0]!="Skip":
                if "/" in data[i][startCol]:#Put dates into correct format
                    #TEMPORARY
                    data[i][startCol]=data[i][startCol].replace("22","16")
                    #FIX DAT!!
                    data[i][startCol]=datetime.strptime(data[i][startCol], '%m/%d/%Y').strftime('%#m/%#d/%y')
                tableList.append(data[i][startCol:])
                typeList.append(data[i][1])
            else:
                tableList.append("Skip")
                typeList.append(data[i][1])
        return tableList,typeList#Return when not enough data in the CSV

    def editSlides(self):#CREATES A NEW FILE AND EDITS THE GIVEN SLIDE
        #1. Load the CSV data 
        csv_data,type_data=self.readCSV()
        printz("1/4 Loaded CSV Correctly")
        #SET PROGRESS BAR
        progressbar.set(0.1)


        #2. Copy the template file 
        try:
            rsp = self.DRIVE.files().get(fileId=self.templateId).execute()
            DATA = {"name":self.outputFile,
                    "parents":[self.outputFolderId]}
            DECK_ID = self.DRIVE.files().copy(body=DATA,fileId=rsp["id"]).execute()["id"]#Save ID of new presentation
        except:
            printz("ERROR: Template Slide does not exist")
            return
        reqs=[]
        counter=0
        #SET PROGRESS BAR
        printz("2/4 Copied Template File")
        progressbar.set(0.3)

        #3. Iterate all slides and create update requests for each table in the CSV
        for slide in self.SLIDES.presentations().get(presentationId=DECK_ID,fields="slides").execute().get("slides",[]):

            #FOR EACH SLIDE!
            if counter<len(csv_data):#If there are still tables to update
                if csv_data[counter]=="Skip":
                    counter+=1
                else:

                    #3.1 Load all of the slides table data
                    tables=[]#Holds table dicts ID, rows, columns
                    table_data=[]#Holds all of table text data in arrays
                    #Check for table objects on the current slide
                    if "pageElements" in slide:
                        for obj in slide["pageElements"]:
                            if "table" in obj:

                                obj_id=obj["objectId"]#Gets object's ID
                                rows=obj["table"]["rows"]#Number of rows (exclude top row)
                                columns=obj["table"]["columns"]#Number of columns
                                tables.append({"ID":obj_id,"rows":rows,"columns":columns})#Adds to the table dict
                                #Get the each cell data from the JSON
                                current_table=[]
                                for x in range(int(rows)):
                                    current_row=[]
                                    if len(obj["table"]["tableRows"][x]["tableCells"])==columns:#Fixes odd headers!!
                                        for y in range(int(columns)):
                                            if "text" in obj["table"]["tableRows"][x]["tableCells"][y]:
                                                #Get the text in the current cell (if there is any)
                                                text=obj["table"]["tableRows"][x]["tableCells"][y]["text"]["textElements"][1]["textRun"]["content"]
                                                #Get the font of a valid cell (as default google font sometimes used)
                                                if y==0:
                                                    style=obj["table"]["tableRows"][x]["tableCells"][y]["text"]["textElements"][1]["textRun"]["style"]
                                                    textFont=style["weightedFontFamily"]
                                                    textSize=style["fontSize"]
                                                    textColor=style['foregroundColor']
                                                    textAlign=obj["table"]["tableRows"][x]["tableCells"][y]["tableCellProperties"]["contentAlignment"]
                                                #If valid text then strip the text of newlines
                                                if len(text.strip(" "))>1:
                                                    text=text.strip("\n")
                                                    current_row.append(text)
                                                else:
                                                    current_row.append("")#Else add empty string
                                            else:
                                                current_row.append("")#If no text in cell then also add empty string
                                    else:
                                        current_row=[""]*columns
                                    current_table.append(current_row)#Add current row to the table
                                table_data.append(current_table)#Add current table to the tables array

                    #Updates every table in the current slide
                    for i in range(len(tables)):
                        #FOR EACH TABLE IN SLIDE
                        normal = True
                        replace_count=0
                        if type_data[min(len(type_data)-1,counter)]=="Replace":
                            started_replace=True
                        else:
                            started_replace=False
                        while counter<len(type_data) and ((type_data[min(len(type_data),counter)]=="Replace" and started_replace) or normal==True) and replace_count<len(table_data[i]):
                            #IF NOT EMPTY ROW
                            if csv_data[counter][0]!="":

                                #3.2 Clean CSV data rows
                                #If an unused column in the row, then add this into the CSV row
                                #Also check if it's cash and then add dollar format if not already
                                for col in range(int(tables[i]["columns"])):
                                    empty=True
                                    for row in range(int(tables[i]["rows"])):
                                        if table_data[i][row][col]!="":
                                            empty=False
                                            #ADD $ signs and money format. REMOVE DECIMALS
                                            if table_data[i][row][col][0]=="$" and "$" not in csv_data[counter][col] and len(csv_data[counter][col])>1:#Put into dollar format
                                                csv_data[counter][col]=to_cash(csv_data[counter][col])
                                            if len(csv_data[counter][col])>3 and "$" in csv_data[counter][col]:
                                                csv_data[counter][col]=csv_data[counter][col].split(".")[0]

                                            #Fix empty points
                                            if csv_data[counter][col][-2:]==".0" or csv_data[counter][col][-3:]==".00":
                                                csv_data[counter][col]=csv_data[counter][col].split(".")[0]
                                    if empty==True:
                                        #Add to the CSV data array
                                        csv_data[counter].insert(col,"")
                                #If CSV row too short, add blank columns
                                for col in range(tables[i]["columns"]-len(csv_data[counter])):
                                    csv_data[counter].append("")



                                #3.3 Create the update requests for each cell
                                #Updates every cell in the current table
                                for x in range(int(tables[i]["rows"])):
                                    if table_data[i][x][0]!="":#If not a heading row
                                        for y in range(int(tables[i]["columns"])):

                                            #Insert text to cell as can't delete if empty
                                            reqs.append(insertRequest(tables[i]["ID"],x,y,"Temp"))
                                            #Delete any text in the cell
                                            reqs.append(deleteRequest(tables[i]["ID"],x,y))
                                            
                                            #Add new data into the current cell
                                            if x<tables[i]["rows"]-1:#Add data from row below
                                                reqs.append(insertRequest(tables[i]["ID"],x,y,table_data[i][x+1][y]))
                                                if table_data[i][x+1][y]!="":
                                                    #Fix the style if changed!
                                                    reqs=reqs+styleRequest(tables[i]["ID"],x,y,textFont,textSize,textColor,textAlign)
                                            
                                            else:#Add CSV data row!
                                                reqs.append(insertRequest(tables[i]["ID"],x,y,csv_data[counter][y]))
                                                #Fix the style if changed!
                                                if csv_data[counter][y]!="":
                                                    reqs=reqs+styleRequest(tables[i]["ID"],x,y,textFont,textSize,textColor,textAlign)
                                                

                                #ADD REPLACE LINE
                                print(table_data[i][1])
                                print(csv_data[counter])
                                table_data[i].append(csv_data[counter])
                                table_data[i].pop(1)

                            counter+=1#Append this, tracks updated tables
                            replace_count+=1
                            normal=False



                        #WINDOW RELATED LINE!!!!
                        progressbar.set(0.3+(counter/len(csv_data))*0.4)
                        ####

        printz("3/4 Sending API Request...")
        #4. Use batch update to request changing of text items
        self.SLIDES.presentations().batchUpdate(body={"requests":reqs},presentationId=DECK_ID,fields="").execute()
        progressbar.set(1)
        printz("4/4 New Slides Generated!")

#Puts value into dollar format
def to_cash(text):
    try:
        newText='{:20,.2f}'.format(float(text))
        newText="$"+str(newText)
        newText=newText[:-3]#Remove cents!!
        newText=newText.replace(" ","")
        return newText
    except:
        return text

#Creates an API insert request
def insertRequest(id,x,y,text):
    return {"insertText":{"objectId":id,"cellLocation":{"rowIndex":x,"columnIndex":y},"text":text}}

#Creates an API delete request
def deleteRequest(id,x,y):
    return {"deleteText":{"objectId":id,"cellLocation":{"rowIndex":x,"columnIndex":y},"textRange":{"startIndex":0,"type":"FROM_START_INDEX"}}}

#Creates an API update style request
def styleRequest(id,x,y,font,size,color,align):
    align = {"updateParagraphStyle":{"objectId":id,"cellLocation":{"rowIndex":x,"columnIndex":y},"style":{"alignment":"CENTER"},"fields":"alignment"}}
    style = {"updateTextStyle":{"objectId":id,"cellLocation":{"rowIndex":x,"columnIndex":y},"style":{"weightedFontFamily":font,"fontSize":size,"foregroundColor":color},"fields": 'foregroundColor,weightedFontFamily,fontSize'}}
    return [align,style]























#WINDOW CODE BELOW HERE ############################################


#####################################################################OPEN FILES FROM DRIVE
#Opens a window to select a google presentation file
#Class to manage each file in the custom explorer
class File:
    def __init__(self,id,name,mimeType):
        self.id=id
        self.name=name
        if mimeType=='application/vnd.google-apps.presentation':
            self.type="presentation"
        else:
            self.type="file"
        self.children=[]
        self.open=False

    def add(self,f):#Add a file into the structure
        if f["parents"][0]==self.id:
            self.children.append(File(f["id"],f["name"],f["mimeType"]))
            return True
        else:
            for c in self.children:
                if c.add(f):
                    return True
        return False

    def button(self,frameS,c,d,mode):#Frame,counter,depth,mode(template,output)
        tab="     "*d
        if self.type=="file":
            tab=tab+"📁 "
            if mode=="template":
                com=self.enter_template
            else:
                com=self.enter_output
        else:
            tab=tab+"🗎 "
            if mode=="template":
                com=self.set_file
            else:
                com=None
        button=customtkinter.CTkButton(master=frameS,text=str(tab)+str(self.name),fg_color="transparent",anchor="w",width=460,command=com)
        button.grid(row=c,column=0,sticky="w") 
        if self.open:
            c+=1
            for child in self.children:
                c=c+child.button(frameS,c,d+1,mode)
            return len(self.children)+1
        return 1

    def enter_output(self):#Enter the outputs name
        global output_folder_id,output_folder_name
        self.open = not self.open
        output_folder_id=self.id
        output_folder_name=self.name
        render_files("output")

    def enter_template(self):
        self.open = not self.open
        render_files("template")

    def set_file(self):#Set the selected file in the file explorer
        global template_filename,template_id
        template_filename=self.name
        template_id=self.id
        render_files("template")
    
#Formats the file list into a class
def format_files(files,rootId):
    root=File(rootId,"Drive","file")
    files.reverse()
    change=1
    while len(files)>0 and change!=0:#Stops loop if failing
        change=0
        for f in files:
            added=root.add(f)
            if added:
                files.remove(f)
                change+=1

    return root

#Loads the files from google drive
def get_drive_files():
    creds=None
    files=[]
    if os.path.exists(TOKEN):
        creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
        if creds.valid:
            try:#Try to connect to drive and slides API
                drive = build('drive', 'v3', credentials=creds)
                printz("Opening Google Drive...")
                driveid = drive.files().get(fileId='root').execute()['id']
                temp = drive.files().list(q="(mimeType='application/vnd.google-apps.presentation' or mimeType='application/vnd.google-apps.folder') and trashed = false and 'me' in owners",
                                        fields="files(id,name,parents,mimeType)",
                                        spaces="drive").execute()
                #GET THIS INTO A TREE STRUCTURE!!
                printz("Google Drive Open...")
                for f in temp.get("files",[]):
                    files.append(f)
                try:
                    files=format_files(files,driveid)
                except:
                    printz("Error formatting the files!")
                    return []
            except HttpError as err:#If cannot connect then output the error
                printz(err)
    return files

#Destroys the root frame of the explorer
def destroy_rootS():
    if "rootS" in globals():
        try:
            global rootS
            rootS.destroy()
        except:
            print("Error destroying window!")


def render_files(mode):
    for w in frameS.winfo_children():
        w.destroy()
    if mode=="template":
        text="File: "+template_filename
    else:
        text="Folder: "+output_folder_name
    file_text = customtkinter.CTkLabel(master=rootS, text=text, fg_color="transparent",width=200,anchor="w")
    file_text.place(x=20,y=266,anchor=NW)
    try:
        FILES.button(frameS,0,0,mode)
    except:
        printz("Error drawing the file buttons!")
    rootS.mainloop()


def create_explorer():
    global rootS,frameS,FILES
    try:
        if STARTED:
            return False
        try:
            FILES = get_drive_files()#LOAD THE GOOGLE DRIVE FILES!
        except:
            printz("Error getting the drive files!")
            return False
        if FILES==[]:
            printz("Google Token Expired! Login Again")
            return False
        #CREATE WINDOW
        destroy_rootS()
        rootS=customtkinter.CTk()
        rootS.geometry("500x300")
        rootS.resizable(False,False)
        frameS = customtkinter.CTkScrollableFrame(master=rootS, width=460, height=240)
        frameS.place(x=10,y=10,anchor=NW)
        return True
    except:
        printz("Error creating the explorer!")
        return False

def submit_template():
    template_text.configure(text=template_filename)
    rootS.destroy()

def submit_output(name):
    global output_filename
    output_filename=name
    output_text.configure(text=output_filename)
    rootS.destroy()

def select_template():
    try:
        if create_explorer():
            rootS.title("Google Drive: Select Template")
            submit_button=customtkinter.CTkButton(master=rootS,text="Select",width=80,command=submit_template)
            submit_button.place(x=410,y=266) 
            try:
                render_files("template")
            except:
                printz("Error rendering the drive files!")
    except:
        printz("Error creating window")

def select_output():
    try:
        if create_explorer():
            rootS.title("Google Drive: Select Output Folder")
            #HAVE AN ENTRY FOR THE NAME HERE!!!
            output_entry = customtkinter.CTkEntry(master=rootS, placeholder_text="Name",width=150)
            output_entry.place(x=250,y=266,anchor=NW)
            submit_button=customtkinter.CTkButton(master=rootS,text="Select",width=80,command=lambda:submit_output(output_entry.get()))
            submit_button.place(x=410,y=266) 
            try:
                render_files("output")
            except:
                printz("Error rendering the drive files!")
    except:
        printz("Error creating window")



##################################################################################################


#Main Functions

#Prints to the output window
def printz(string):
    textbox.insert(index="0.0",text=str(string)+"\n")
    root.update()

#Opens a window to select a file
def select_file():
    global filename
    if STARTED:
        return
    filetypes = (
        ('CSV files', '*.csv'),
        ('All files', '*.*')
    )
    filename = customtkinter.filedialog.askopenfilename(
        title='Open CSV file',
        initialdir='/',
        filetypes=filetypes)
    file_text.configure(text=filename.split("/")[-1])

#Allows the user to log in
def log_in():
    if STARTED:
        return
    creds=None
    if os.path.exists(TOKEN):
        creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDS, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN, 'w') as token:
            token.write(creds.to_json())
    check_credentials()

#Allows the user to log out
def log_out():
    if STARTED:
        return
    os.remove(TOKEN)
    check_credentials()

#Checks the if credentials and token exist
def check_credentials():
    creds,token=False,False
    if os.path.exists(TOKEN):
        token=True
        token_text.configure(text="✓ token.json")
    else:
        token_text.configure(text="❌ token.json")
    if os.path.exists(CREDS):
        creds=True
        cred_text.configure(text="✓ credentials.json")
    else:
        cred_text.configure(text="❌ credentials.json")

    if creds and token:
        login_button.configure(text="Log Out",command=log_out)
        return True
    elif creds:
        login_button.configure(text="Log In",command=log_in)
    else:
        login_button.configure(text="NO CREDENTIALS",command=None)

#Generates the slides if the inputs seem valid
def start():
    global STARTED
    CSVtoGSlideRipple(filename,template_id,output_filename,output_folder_id)#Data, template, new
    progressbar.set(0)
    STARTED=False

#Generates the new slides
def generate_slides():
    global STARTED
    if STARTED==False:
        if check_credentials and filename!="" and output_filename!="" and template_id!="" and output_folder_id!="":
            destroy_rootS()
            Thread(target=start).start()
            STARTED=True
        else:
            printz("Invalid Files Given! Check Again")
    else:
        printz("Program already running!")


#MAIN CODE HERE
if __name__ == '__main__':
    filename=""
    template_filename=""
    output_filename=""
    template_id=""
    output_folder_id=""
    output_folder_name=""
    STARTED=False
    TOKEN="token.json"
    CREDS="credentials.json"
    SCOPES = ['https://www.googleapis.com/auth/presentations','https://www.googleapis.com/auth/drive']

    
    #CREATING THE CUSTOM TKINTER WINDOW
    customtkinter.set_appearance_mode("dark")
    root=customtkinter.CTk()
    root.geometry("600x400")
    root.title("CSV to Google Slide Tables")
    root.resizable(False,False)

    frame = customtkinter.CTkFrame(master=root, width=350, height=350)
    frame.place(x=25,y=20,anchor=NW)
    frame2= customtkinter.CTkFrame(master=root, width=175, height=200)
    frame2.place(x=400,y=95,anchor=NW)
    
    #FRAME 2 HERE
    cred_text = customtkinter.CTkLabel(master=frame2, text="✓ Credential.json", fg_color="transparent")
    cred_text.place(relx=0.5,rely=0.2,anchor=CENTER)

    token_text = customtkinter.CTkLabel(master=frame2, text="❌ Token.json", fg_color="transparent")
    token_text.place(relx=0.5,rely=0.4,anchor=CENTER)

    login_button=customtkinter.CTkButton(master=frame2,text="Login")
    login_button.place(relx=0.5,rely=0.8,anchor=CENTER)
    check_credentials()
    #FRAME 1 HERE
    csv_text = customtkinter.CTkLabel(master=frame, text="CSV File:", fg_color="transparent")
    csv_text.place(x=20,y=30,anchor=NW)

    csv_button=customtkinter.CTkButton(master=frame,text="Open File",command=select_file)
    csv_button.place(x=25,y=60,anchor=NW)

    file_text = customtkinter.CTkLabel(master=frame, text="", fg_color="transparent")
    file_text.place(x=180,y=60,anchor=NW)

    template_text = customtkinter.CTkLabel(master=frame, text="Template File:", fg_color="transparent")
    template_text.place(x=20,y=90,anchor=NW)

    template_button=customtkinter.CTkButton(master=frame,text="Open File",command=select_template)
    template_button.place(x=25,y=120,anchor=NW)

    template_text = customtkinter.CTkLabel(master=frame, text="", fg_color="transparent")
    template_text.place(x=180,y=120,anchor=NW)

    output_text = customtkinter.CTkLabel(master=frame, text="Output File:", fg_color="transparent")
    output_text.place(x=20,y=150,anchor=NW)

    output_button=customtkinter.CTkButton(master=frame,text="Open File",command=select_output)
    output_button.place(x=25,y=180,anchor=NW)

    output_text = customtkinter.CTkLabel(master=frame, text="", fg_color="transparent")
    output_text.place(x=180,y=180,anchor=NW)

    button=customtkinter.CTkButton(master=frame,text="Generate",command=generate_slides)
    button.place(x=185,y=230,anchor=NW)

    textbox=customtkinter.CTkTextbox(master=frame,width=310,height=70)
    textbox.place(x=20,y=270,anchor=NW)

    progressbar = customtkinter.CTkProgressBar(master=root, orientation="horizontal",width=620,height=10)
    progressbar.place(relx=0.5,rely=0.99,anchor=CENTER)
    progressbar.set(0)
    printz("Program Loaded Correctly!")
    root.mainloop()
