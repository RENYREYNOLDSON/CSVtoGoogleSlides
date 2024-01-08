

###############################
############################### CSV to GOOGLE SLIDE WEB APP!
############################### Version 3.0
############################### Uses the google API to load, copy, edit and save a google slides file using given CSV data
############################### def EditSlides is responsible for the main editing!
###############################
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

import awsgi
from flask import Flask
from flask import render_template,request


app = Flask(__name__)

filename=""
template_filename=""
output_filename=""
template_id=""
output_folder_id=""
output_folder_name=""
STARTED=False
TOKEN="token.json"
CREDS="credentials.json"
SCOPES = ["https://www.googleapis.com/auth/drive.metadata.readonly",'https://www.googleapis.com/auth/presentations','https://www.googleapis.com/auth/drive']


class CSVtoGSlideRipple:#USES GOOGLE API TO EDIT GOOGLE SLIDE FROM CSV DATA
    def __init__(self,CSVdata,templateId,outputFile,token):
        self.CSVdata=CSVdata["data"]
        self.CSVdata.remove([""])
        print(self.CSVdata)
        self.templateId=templateId
        self.outputFile=outputFile
        self.TOKEN=token
        self.CREDS=CREDS
        #self.TOKEN=TOKEN
        self.SCOPES=SCOPES
        link="FAILED"
        self.connectAPI()
        link = self.editSlides()
        return link

    def connectAPI(self):#CONNECTS TO THE GOOGLE API AND CREATES A TOKEN
        creds = None
        # The file token.json stores the user's access and refresh tokens, created when user logs in
        """
        creds = Credentials.from_authorized_user_info({"token":self.TOKEN,
                                                       "client_secret":"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",
                                                       "client_id":"341236422016-fshrrsdi9un71mjrnh0rjhb0ee158k7o.apps.googleusercontent.com",
                                                       "token_uri":"https://oauth2.googleapis.com/token",
                                                       "refresh_token":self.TOKEN},
                                                        self.SCOPES)
        """
        creds = Credentials(token=self.TOKEN,
                            client_secret="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",
                            client_id="341236422016-fshrrsdi9un71mjrnh0rjhb0ee158k7o.apps.googleusercontent.com",
                            token_uri="https://oauth2.googleapis.com/token",
                            scopes=self.SCOPES)
        
        try:#Try to connect to drive and slides API
            self.DRIVE= build('drive', 'v3', credentials=creds)
            self.SLIDES = build('slides', 'v1', credentials=creds)
            print("CONNECTED")
        except:#If cannot connect then output the error
            print("Google Token Expired! Login Again")


    def readCSV(self):#FORMATS THE CSV DATA WHEN WE KNOW THE TEMPLATE TABLE SIZES
        #Load the CSV file Here
        data=self.CSVdata
        tableList=[]#Holds all CSV collected tables
        typeList=[]
        #FIND START COL!!!Using the date
        startCol=0
        for i in range(len(data[3])):
            if data[3][i].count("/")==2:
                startCol=i
                print("Starting at column "+str(startCol))
        #Add each data row to an array
        for i in range(1,len(data)):
            if data[i][0]!="Skip":
                print(str(i)+","+str(len(data)))
                print(data[i])
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
        print(type_data)
        print("1/4 Loaded CSV Correctly")
        #SET PROGRESS BAR


        #2. Copy the template file 

        rsp = self.DRIVE.files().get(fileId=self.templateId).execute()
        DATA = {"name":self.outputFile}
        DECK_ID = self.DRIVE.files().copy(body=DATA,fileId=rsp["id"]).execute()["id"]#Save ID of new presentation
        link = self.DRIVE.files().get(fileId=DECK_ID,fields="webViewLink").execute()["webViewLink"]
        print(link)


        reqs=[]
        counter=0
        #SET PROGRESS BAR
        print("2/4 Copied Template File")

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
                        if counter<len(csv_data):#If we still have tables to update

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
                                                
                        counter+=1#Append this, tracks updated tables
                        #WINDOW RELATED LINE!!!!
                        ####
        print("3/4 Sending API Request...")
        #4. Use batch update to request changing of text items
        self.SLIDES.presentations().batchUpdate(body={"requests":reqs},presentationId=DECK_ID,fields="").execute()
        print("4/4 New Slides Generated!")
        return link

#Puts value into dollar format
def to_cash(text):
    newText='{:20,.2f}'.format(float(text))
    newText="$"+str(newText)
    newText=newText[:-3]#Remove cents!!
    newText=newText.replace(" ","")
    return newText

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

#Generates the slides if the inputs seem valid
def start():
    global STARTED
    CSVtoGSlideRipple(filename,template_id,output_filename,output_folder_id)#Data, template, new
    STARTED=False




### REST API CODE HERE!!!

@app.route('/')

@app.route("/index")
def spec():
    user = {'username': 'Adam'}
    return render_template('index.html', title='Home', user=user)


@app.route('/generate_slides', methods=['POST'])
def generate_slides():
    # 1 Login Token
    # 2 CSV File
    # 4 Template File ID
    # 5 Parent File ID
    # 6 Output File
    json=request.get_json()
    token=json["accessToken"]
    fileID=json["fileID"]
    outputName=json["outputName"]
    CSVfile=json["CSVdata"]
    link = CSVtoGSlideRipple(CSVfile,fileID,outputName,token)

    return "SLIDES GENERATED "+link


#Lambda Event Handler
def lambda_handler(event, context):
    return awsgi.response(app, event, context, base64_content_types={"image/png"})


# Need to do:
# Make it look nice
# Give indication of how long left
# Do validation on generation
# Send json with link
