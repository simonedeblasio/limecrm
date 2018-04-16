![alt text](https://www.getaccept.com/assets/img/GetAccept_Logo_Grey_Web.png "Logo Title Text 1")

# Close more deals faster
GetAccept is a third party tool helping sales people close more deals by taking control of the proposal and eSigning workflow. Now it is integrated with LIME CRM and you can send your document directly from LIME CRM through GetAccept. 

Features :
- Send all doucment from Lime CRM through GetAccept
- Document tracking, who has recived and opened the sent document
- Document analytics, when was it opened, how many times, what pages did they spend time on etc
- Commenting, discuss your proposal directly in the document
- Automatic reminders smooothly moving your deal forward
- eSigning, make it easy for you customers to say and skip the hazle with printers and scanners
- Automatic downloading of your signed documents **(NEW)**

----------

# Installation & Configuration

**Good to know:** The integration is built for the **"Lime Core database"**. If you need to install it in a different solutions you may need to do some manual configurations.

## How does it work
You can add the GetAccept App on every object where there is a document tab present. 

## Suported file types 
There is a list in the VBA in GetAccept.CheckFileTypes where you can configure which file types that the integration should accept. Before adding a file type should you check if GetAccept can handle it. 

## Setup
1. [Files] - Make sure to download both: 
	[GetAccept-v2](https://github.com/getaccept/limecrm/tree/master/GetAccept-v2)
	& 
	[GetAcceptEmail](https://github.com/getaccept/limecrm/tree/master/GetAcceptEmail)
2. [Files] - Copy the folders "GetAccept-v2" and "GetAcceptEmail" to the apps folder in the Actionpad-folder. (Don't forget to unblock files before unzipping and moving them)

3. [LISA] - Add a yes/no field named to "sent_with_ga" to the document table, set it as "Read only for LIME PRO" in LISA.
4. [LISA/VBA] - Check if the History > type field is named "type" and if there is an option with they key "sentemail". this will be set in the VBA and if it doesn't exist it will not work (you can change it in the vba GetAccept.SetDocumentStatus)

5. [VBA] - Import the GetAccept.bas ("..\Install\VBA") to the VBA
6. [VBA] - Run the Install method in the GetAccept VBA module. You must have a localization table in the databas. It is  translated to English, Swedish, Norwegia, Danish and Finish. Check which fields you have in your localization table. Dependent on which fields you have you need to remove languages in AddOrCheckLocalize in the VBA (Example: If the lanugage Norwegian or Danish is missing in you localization table you should remove oRec.Value("no") = sNO and oRecs(1).Value("no") = sNO) and so on.
7. [VBA] - Restart Lime CRM or run ThisApplication.setup to load the new translations

8. [Actionpad] - Import the html-tag below to the actionpad where you want the GetAccept App tho be shown. It's most commonly used from company.html or deal.html. Place the html-tag in the actionpad header. 
The table must have a document tab and you must be able to connect to a person tab either directly on the card or on a related table.

9. [LIME] - Publish the actionpad!

10. [Do this test](https://github.com/getaccept/limecrm/blob/master/GetAccept-v2/Install/test-of-workflow.md)

``` html
<div data-app="{app:'GetAccept-v2',config:{
	title_field: 'comment', 
	personSourceTab: '', 	
	personSourceField: 'company',
	businessValue: 'value'  
	}}">
</div>
```

## Configuration:
- title_field: The document name field
- personSourceTab: If there is a realiton tab on the object where it shoud look for recipient persons directly, ex: if you place it in company.hmtl you should have persons
- personSourceField: If there is a realtion field on the object where it should look for persons connected to a sub table, ex: if you place it in busniess.html and you have a connection to the company where - persons are connected. 
- businessValue The name of the field containing the busniessvalue

You are now done. Each user will have their own login credentials which is used to start using the GetAccept integration.

---------

# Two ways integration (NEW)
__Requires the Lime CRM API and a api key.__
This feature allows GetAccept to automatically post back a signed copy of your signed documents to Lime CRM. It will download the signed document with the signing certificate and store it back in the CRM. 

## How to set it up in LIME CRM
1. Create a API-user.
2. Give correct permissions to the user (should follow the LIME standard rules (Add/Read/Write))

## How to set it up in GetAccept.
1. Log on to the GetAccept web application. [app.getaccept.com](https://app.getaccept.com)
2. Go to **Settings** --> **Integrations** 
3. Add your api key and the server url to the integration page. 

**Ex: Domain URL:** https://[URL]/[DatabaseName]
		https://gaCRMDemo/getaccept%20CRM
		
**Ex: API-key:** 3FD114540187E43A9264743B7742528429511C042237ACF10034DEBEAADF770ECFBD8F966187491C7C62

## Document table setup
#### You need to have following fields in the Document table to be able to use the two way integration: 
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


# Troubleshoot
Some common errors: 

---------

CREATED BY: GetAccept
To use this app you need to have a GetAccept account, create one for free att www.getaccept.com 


