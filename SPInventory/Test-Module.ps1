#This part must be run on the SharePoint farm
Remove-Module SPInventory
Import-Module 'C:\Projects\SPInventory\SPInventory\SPInventory.psm1'
Export-SPIObject -Output 'C:\Projects\SPInventory\spinventory.xml'

#This part must be run on a PC with MS Word installed
#Import-Module 'SPIDocumentGenerator'
#New-SPInventoryDocument -InputFilePath 'C:\Projects\SPInventory\spinventory.xml' -OutputFilePath 'C:\Projects\SPInventory\spinventory.docx'