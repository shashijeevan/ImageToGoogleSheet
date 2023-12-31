# Copyright (C) 2023 Shashi Jeevan Midde Peddanna
#
# This file is released under the "GNU General Public License v3.0". 
# Please see the LICENSE file that should have been included as 
# part of this package.

import os
import datetime as dt
import sys
import json
import traceback
import logging
import numpy as np
import httplib2
from googleapiclient.discovery import build
from google.oauth2  import service_account
import google_auth_httplib2
from pprint import pprint
from PIL import Image
from ImageToJSON import RequestBuilder

class ImageToGoogleSheet:

    def __init__(self):
        self.imagepath = ""
        self.sheetName = ""
        self.sheetID = 0
        self.SQUAREWIDTH = 21
        self.SPREADSHEET_ID = os.environ['SPREADSHEETID']
        self.SERVICE_KEY = 'key.json'

    def __SetupProxy(self, creds, UseProxy):

        if(UseProxy):

            #Below code is for using proxy
            localproxy = httplib2.ProxyInfo(httplib2.socks.PROXY_TYPE_HTTP, 
                                            '127.0.0.1', 
                                            8888)
            
            http = httplib2.Http(proxy_info=localproxy, 
                                    disable_ssl_certificate_validation=True)

            authorized_http = google_auth_httplib2.AuthorizedHttp(creds, http=http)

            _service =  build('sheets', 'v4', http=authorized_http)
        else:
            #Not using proxy
            _service = build('sheets', 'v4', credentials=creds)

        return _service
    
    def __LoadImage(self):

        '''This private method loads the image and 
        creates an array of pixel RGB values'''
        
        print("---------------------")
        print("Loading Image " + self.imagepath)
        
        #Opening the image and checking image properties
        inputimage = Image.open(self.imagepath)

        self.width, self.height  = inputimage.size
        print('Image loaded. Width = ', self.width, ', Height = ', self.height)

        #Checking if image mode is RGB and convert to RGB if required
        print("Image mode = " + inputimage.mode)
        
        if inputimage.mode != "RGB":
            print("Image mode is not RGB. Converting to RGB")
            inputimage = inputimage.convert("RGB")

        #Loading image to the pixel data. This is an internal data structure of Pillow
        image_sequence = inputimage.getdata()

        #Sequence to ndarray for easy handling
        self.image_array = np.array(image_sequence)

        print("Array dimension = ", self.image_array.shape)

    def __AddNewSheet(self):

        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        creds = None

        creds = service_account.Credentials.from_service_account_file(self.SERVICE_KEY, 
                                                                      scopes=SCOPES)

        UseProxy = False

        service = self.__SetupProxy(creds, UseProxy)

        print("---------------------")
        print("Generating new Sheet name")

        #Simple random sheet name generator
        #Month WeekdayName interger part of timestamp
        now = dt.datetime.now()
        self.sheetName = f"{now:%b} {now:%a} {int(now.timestamp())}" 
        print(f"Sheet Name: {self.sheetName}")

        #Creating as new Sheet in the SpreadSheet. 
        #Using the Image dimensions as the Sheet dimensions.
        requests = []
 
        requests.append(
            {
                "addSheet": {
                    "properties": {
                        "title": self.sheetName,
                        "gridProperties": {
                            "rowCount": self.height,
                            "columnCount": self.width
                        }
                    }
                }
            }            
        )
        
        body = {
            'requests': requests
        }

        print("Adding new Sheet")

        addsheetresponse = service.spreadsheets().batchUpdate(
                                    spreadsheetId=self.SPREADSHEET_ID, 
                                    body=body).execute()

        print("Added new Sheet")

        self.sheetID = addsheetresponse['replies'][0]['addSheet']['properties']['sheetId']

        #Changing the column width of all the Columns to make it square
        requests = []

        requests.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': self.sheetID,
                    'dimension': "COLUMNS",
                    'startIndex': 0,
                    'endIndex': self.width,
                },
                'properties': {
                    'pixelSize': self.SQUAREWIDTH
                },
                'fields': 'pixelSize'
            }
        })

        body = {
            'requests': requests
        }

        print("Setting square cells")

        setwidthresponse = service.spreadsheets().batchUpdate(
                                    spreadsheetId=self.SPREADSHEET_ID, 
                                    body=body).execute()

        print("Setup of square cell is complete")

    def __SaveImageArrayToGSheet(self):

        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        creds = None

        creds = service_account.Credentials.from_service_account_file(self.SERVICE_KEY, 
                                                                        scopes=SCOPES)

        UseProxy = False

        service = self.__SetupProxy(creds, UseProxy)
        
        print("---------------------")
        
        #Generating the JSON using the Image data
        print("Building the request JSON from image pixel data")

        requestSON = RequestBuilder.ImageArrayToJSON(self.sheetID, self.image_array, 
                                                     self.width, self.height)

        print("JSON generation completed")

        logging.info(requestSON)

        #We need build Python dictionary from JSON string as API takes JSON object
        request = json.loads(requestSON)

        print("Submitting batch update request")
        
        response = service.spreadsheets().batchUpdate(spreadsheetId=self.SPREADSHEET_ID,
                                                      body=request).execute()

        print("Submitted batch update successfully")

    def SaveImageToGoogleSheet(self, imagepath: str):

        '''This public method takes in the images, adds a new sheet and 
        updates the Google Sheet with Image pixel data.'''

        self.imagepath = imagepath

        try:
            self.__LoadImage()
        except Exception as inst:
            print("Failed to load Image.")
            print(inst)
            return
        
        try:
            self.__AddNewSheet()
        except Exception as inst:
            print("Failed to Add new sheet.")
            print(inst)
            return
        
        try:
            self.__SaveImageArrayToGSheet()
        except Exception as inst:
            print("Failed to update Google sheet file.")
            print(inst)
            traceback.print_exc()
            return
        
if __name__ == "__main__":

    logging.basicConfig(filename='imagetogogglesheet.log', 
                        encoding='utf-8', level=logging.INFO)

    if(len(sys.argv) <= 1):
        print("Please pass the file name")
    else:
        imagefile = sys.argv[1]
        print(f"File name is {imagefile}")
        
        imageload = ImageToGoogleSheet()
        imageload.SaveImageToGoogleSheet(imagefile)
