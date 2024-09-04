# Description
A PowerShell script designed to download files from an SFTP server or upload files to an SFTP server. 

All required parameters are read from a [.JSON file](Example.json). The input file can be modified as needed. 

A scheduled task can be created to run the script in an automated way.

# Best practices
- Read the [help section in the script](Move%20file%20over%20SFTP.ps1) for an explanation of the input file parameters.

# Requirements
- Download the required modules from the [repository](https://github.com/DarkLite1?tab=repositories)
  (See the `#Requires` section at the top of [the script](Move%20file%20over%20SFTP.ps1))
- The machine running the SFTP client needs to have the module [Posh-SSH](https://github.com/darkoperator/Posh-SSH) installed