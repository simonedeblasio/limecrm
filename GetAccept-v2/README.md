#GetAccept - eSigning Version 2

CREATED BY: GetAccept
To use this app you need to have a GetAccept account, create one for free att www.getaccept.com 

# Close more deals faster
GetAccept is a third party tool helping sales people close more deals by taking control of the proposal and eSigning workflow. Now it is integrated with LIME Pro and you can send your document directly from LIME Pro through GetAccept. 

Features :
- Send all doucment from Lime Pro through GetAccept
- Document tracking, who has recived and opened the sent document
- Document analytics, when was it opened, how many times, what pages did they spend time on etc
- Commenting, discuss your proposal directly in the document
- Automatic reminders smooothly moving your deal forward
- eSigning, make it easy for you customers to say and skip the hazle with printers and scanners
- Automatic downloading of your signed documents **(NEW)**

# How does it work
You can add the GetAccept App on every object where there is a document tab present. 

##Suported file types 
There is a list in the VBA in GetAccept.CheckFileTypes where you can configure which file types that the integration should accept. Before adding a file type should you check if GetAccept can handle it. 

----------

# Installation
1. Copy the "GetAccept" folder to the apps folder in the Actionpad-folder.
2. Add a yes/no field named to "sent_with_ga" to the document table, set it as protected for editing in Lime Pro
3. Check if the History > type field is named "type" and if there is an option with they key "sentemail". this will be set in the VBA and if it doesn't exist it will not work (you can change it in the vba GetAccept.SetDocumentStatus)
4. Import the GetAccept.bas ("..\Install\VBA") to the VBA
5. Run the Install method in the GetAccept VBA module. You must have a localization table in the databas. It is  translated to English, Swedish, Norwegia, Danish and Finish. Check which fields you have in your localization table. Dependent on which fields you have you need to remove languages in AddOrCheckLocalize in the VBA (Example: If the lanugage Norwegian or Danish is missing in you localization table you should remove oRec.Value("no") = sNO and oRecs(1).Value("no") = sNO) and so on.
6. Import the html-tag below to the tables where you want the GetAccept App tho be shown. most commonly used from company.html or busniess.html. Th table must have a document table and you must be able to connect to a person tab either directly on the table or on a related table.

``` html
<div data-app="{app:'GetAccept-v2',config:{
	title_field: 'comment', 
	personSourceTab: '', 	
	personSourceField: 'company',
	businessValue:''  
	}}">
</div>
```
# Configuration:
- title_field: The document name field
- personSourceTab: If there is a realiton tab on the object where it shoud look for recipient persons directly, ex: if you place it in company.hmtl you should have persons
- personSourceField: If there is a realtion field on the object where it should look for persons connected to a sub table, ex: if you place it in busniess.html and you have a connection to the company where - persons are connected. 
- businessValue The name of the field containing the busniessvalue

You are now done. Each user will have their own login credentials which is used to start using the GetAccept integration.

---------

## Two ways integration (NEW)
__Requires the Lime CRM API and a api key.__
This feature allows GetAccept to automatically post back your signed documents to Lime CRM.
Log on to the GetAccept web application. Go to **Settings** --> **Integrations** --> and then add your api key and the server url to the integration page. 

Ex: Domain URL: https://[URL]/[DatabaseName]
		https://gaCRMDemo/getaccept%20CRM
		
API-key: 3FD114540187E43A9264743B7742528429511C042237ACF10034DEBEAADF770ECFBD8F966187491C7C62


### You need to have following fields in the Document table to be able to use the two way integration: 
[getacceptstatus] - type: Option field, 
	values: 
		name: Draft, key: draft
		name: Sent, key: sent
		name: Reviewd, key: reviewd
		name: Signed, key: signed
	
[sent_with_ga] - type: Yes/No field

[comment] - type: Text field

[document] - type: File field

[type] - type: Option field,
	values: 
		name: Agreement, key: agreement

	
# Important
Each user at a company using the GetAccept integration need to have a GetAccept account.
