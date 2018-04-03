import xlrd
from xlrd import open_workbook, cellname
import xlsxwriter

#Prepare file name and sheet

FileName = 'test'
SheetName = 'Air Canada Call Scoring'

SrcName = FileName + '.xlsx'
DestName = FileName + '-New' + '.xlsx'

SrcBook = xlrd.open_workbook(SrcName, 'r')
ScrSheet = SrcBook.sheet_by_index(0)

DestBook  = xlsxwriter.Workbook(DestName)
DestSheet = DestBook.add_worksheet(SheetName)



#Prepare Azure Text Analytics API Connection

import requests
subscription_key = "68ac716209334cf0a0a96cb41f8cb1b9"
assert subscription_key
headers   = {"Ocp-Apim-Subscription-Key": subscription_key}
text_analytics_base_url = "https://eastus2.api.cognitive.microsoft.com/text/analytics/v2.0/"
sentiment_api_url = text_analytics_base_url + "sentiment"
key_phrase_api_url = text_analytics_base_url + "keyPhrases"


#Copy Source sheet from source to new file

for row_index in range(ScrSheet.nrows):
        print(row_index/ScrSheet.nrows)
        for col_index in range(ScrSheet.ncols):
            DestSheet.write(row_index, col_index, ScrSheet.cell(row_index, col_index).value)
            
            #Prepare data and send data (column "text" only) to Azure for analysation(sentiment and key phrase) 
            if (row_index > 0) & (col_index == 5):
                
                AzureData1 = []
                AzureData2 = []
                
                AzureData1 = [{'id': 1,'language': 'en','text':ScrSheet.cell(row_index, col_index-1).value}]
                AzureData2 = [{'id': 1,'text':ScrSheet.cell(row_index, col_index-1).value}]
                
                documents_sentiments = {}
                documents_sentiments = {}
                
                documents_sentiments = {'documents': AzureData1}
                documents_key_phrases = {'documents': AzureData2}
                
                response_sentiments  = requests.post(sentiment_api_url, headers=headers, json=documents_sentiments)
                response_key_phrases  = requests.post(key_phrase_api_url, headers=headers, json=documents_key_phrases)
                
                sentiments = response_sentiments.json() 
                documents_key_phrases = response_key_phrases.json() 
                
                
                #Update result to next 2 columns
                #Update error message if document size is larger than 5120 characters (limitation on Azure).
                
                if(bool(sentiments['documents']) == True):
                    DestSheet.write(row_index, col_index, sentiments['documents'][0]['score'])
                    DestSheet.write(row_index, col_index+1, str(documents_key_phrases['documents'][0]['keyPhrases']))
                    
                else:
                    DestSheet.write(row_index, col_index, sentiments['errors'][0]['message'])
                    DestSheet.write(row_index, col_index+1, str(documents_key_phrases['errors'][0]['message']))

#Close and save file.

DestBook.close()