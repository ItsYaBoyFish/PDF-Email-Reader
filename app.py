# importing required modules 
import PyPDF2
import xlsxwriter

def writeToExcelSpreadsheet(list): 
  # Create workbook
  workbook = xlsxwriter.Workbook('emails.xlsx')
  
  # Add worksheet to the workbook.
  worksheet = workbook.add_worksheet()

  # # Test writing hello world in A1
  # worksheet.write('A1', 'Hello world')

  # create headers
  worksheet.write('A1', 'Emails')

  position = 2

  for item in list: 
    worksheet.write(f'A{position}', item)
    position += 1

  # Close workbook
  workbook.close()

emailDictionary = {
  "Emails": [],
  "Length": 0
}
# creating a pdf file object 
pdfFileObj = open('Gawda.pdf', 'rb') 
  
# creating a pdf reader object 
pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
  
## TODO:
# 1. Loop through the page range.
# 2. Print the results of each page to the console

savedWordsFromData = []
finishedEmailsList = []

for number in range(38, 112):
  print(f'+++++++ Page Number: {number} +++++++')
  # creating a page object 
  pageObj = pdfReader.getPage(number)

  # extracting text from page 
  data = pageObj.extractText().split()

  for index, word in enumerate(data):
    if '@' in word:
      positionOfSubstring = word.find('@') + 1
      wordLength = len(word)

      if(positionOfSubstring == wordLength):
        if(index + 1 < len(data)):
          # if the next word does not have a @ symbol in it.
          if(data[index + 1].find('@') == -1):
            word += data[index + 1]
            savedWordsFromData.append(word)
      else:
        savedWordsFromData.append(word)


# clean up the saved words from data (Trim off the extras after the .com, .net etc)
for word in savedWordsFromData:
  if word.find('.com') or word.find('.net') or word.find('.org'):
    if word.find('.com') != -1:
      if(word.endswith('.com')):
        finishedEmailsList.append(word)
      else:
        substring = word.find('.com') + 4
        finishedEmailsList.append(word[0:substring])
        

    if word.find('.net') != -1:
      if(word.endswith('.net')):
        finishedEmailsList.append(word)
      else:
        substring = word.find('.net') + 4
        finishedEmailsList.append(word[0:substring])

    if word.find('.org') != -1:
      if(word.endswith('.org')):
        finishedEmailsList.append(word)
      else:
        substring = word.find('.org') + 4
        finishedEmailsList.append(word[0:substring])




emailDictionary["Emails"] = finishedEmailsList

# closing the pdf file object 
pdfFileObj.close() 


# update length of the emails.
emailDictionary["Length"] = len(emailDictionary['Emails'])

# print(emails)
print(f"Total Emails: {emailDictionary['Length']}")

writeToExcelSpreadsheet(emailDictionary["Emails"])


