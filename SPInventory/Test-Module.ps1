
#This part must be run on the SharePoint farm
Import-Module 'SPInventory'
Export-SPIObject -Output 'C:\Temp\spinventory.xml'

#This part must be run on a PC with MS Word installed
Import-Module 'SPIDocumentGenerator'
New-SPInventoryDocument -InputFilePath 'C:\Temp\spinventory.xml' -OutputFilePath 'C:\Temp\spinventory.docx'