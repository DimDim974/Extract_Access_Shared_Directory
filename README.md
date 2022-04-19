# Extract_Access_Shared_Directory
Extract accesses from a shared directory

* prerequisite
Have the RSAT module installed on the machine
Have the Excel software installed on the machine

* Description
Extraction of accesses from a shared directory.
Extraction Will provide a file with users and groups.
If the access is managed by a group, an extraction of the group is provided.
For both users and groups, the type of access will be specified (read, modify, full control)
The output file will be in xlsx format.

If the directory has a very large number of directories and accesses under several levels, please modify the IdleTimeout variable, by default it is set to 2 hours. 

* Please use :  
Get-Item WSMan:\localhost\shell\IdleTimeout
 Set-Item WSMan:\localhost\shell\IdleTimeout 2147483647
2147483647 ==> Corresponds to 24 days.
