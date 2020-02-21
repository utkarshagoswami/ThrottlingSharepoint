Do Sharepoint auth
Do CRM Auth
Fetch all sharepointdocumentlocations(SDLs)
For each SDL:
	Check if SDL has regardingobjectid
	Check if SDL points to a record with DL name same as entity logical name
	Create a DL with name provided by the user(if required)
	Copy Files to the new DL recursively
	Update the DL for that record in CRM
	Increase the count of number of records updated
