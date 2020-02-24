Issue: 
User is hitting 5k threshold limit error message while rendering document grid on entity record. Even though that record might have 10-100 of files/folders only.

Solution: 
Move the folders available in the main DL, for instace the "account" DL, to a new DL created inside the same subsite. The new DL can be called "account0". Then, the document location(s) correspondingly can be fixed to point to "account0" DL.
Each DL can have a max of 5k folders in the root. This number is configurable by the user in config.json file.

Config.json is the file where user can specify the configurations and the userID for sharepoint and CRM. The explanation of each item in config.json is as:
- crmUrl: The URL of the CRM organization of the user.
- crmUserName: The username of the user used to log inide CRM. The user must have permissions to read the entity in concern and update sharepoint document locations.
- sharepointUrl: The URL of sharepoint site connected with the above mentioned CRM. This URL must not have subsite.
- sharepointUserEmailID: The user email ID of the user to log inside sharepoint and have access to read, write and update data.
- entityLogicalName: The logical name of the entity in CRM.
- numberOfFoldersPerDL: The number of folders the user wants inside each DL that will be created. If not provided, it defaults to 3000 folders
- isThereAnySubsite: Determines if there is any subsite to where the DLs are in the sharepoint site. The value can be "yes" or "no"
- subsiteUrl: The URL for subsite. For example if the DLs are present in www.abc.com/def, "def" will be the value of this attribute.
- newDLPrefix: The prefix of the new DL to be created. In the abovementioned scenario, it is "account".
- newDLSuffixNumber: The suffix of the new DL to be created. This number will increase everytime a DL is created. In the abovementioned example, it has a value of 0
- descriptionOfNewDL: The description to be mentioned on the DL that is to be created in sharepoint.
- deleteFilesAfterCopying: Determines if the files are to be moved to a new destination or just copied. The value must be "yes" or "no".
- maxRecordsToBeMoved: The maximum number of records to be processed for movement. if not specified, the default value is infinite.

Place the config.json file wherever the .exe file is located. The three files containing logs, viz, working.json, failures.json and success.json will also be created at the same place. A sample config.json file can be found here: https://github.com/utkarshagoswami/ThrottlingSharepoint

Algorithm:
Do Sharepoint auth
Do CRM Auth
Fetch all sharepointdocumentlocations(SDLs)
For each SDL:
	Check if SDL has regardingobjectid
	Check if SDL points to a record with DL name same as entity logical name
	Create a DL with name provided by the user(if required)
	Write the details about entity and SDL in working.xml
	Copy Files to the new DL recursively
	Update the DL for that record in CRM
	Delete the files from the old DL if consented by the user
	Write the details about entity and SDL in success.xml or failures.xml
	Break the code on failure
	Increase the count of number of records updated

Caveats:
- The script does not move eccentric entities.
- If the script closes in between the processing, the entity being updated must be verified and fixed manually