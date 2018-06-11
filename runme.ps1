#Explicitly import the module for testing
Import-Module 'SPIDocumentGenerator'

#Run each module function
New-SPInventoryDocument -InputFilePath 'C:\Temp\spinventory.xml' -OutputFilePath 'C:\Temp\spinventory.docx'