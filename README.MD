# Description
Download files from an SFTP server or upload files to an SFTP server with this PowerShell script.

Simply create an input file based on the [example.JSON file](Example.json) and create a scheduled task to launch the script in an automated way.

# Best practices
- Read the [help section in the script](Move%20file%20over%20SFTP.ps1) for an explanation on the input file parameters.

# Requirements
- Download the required modules from the [repository](https://github.com/DarkLite1?tab=repositories)  (See the `#Requires` section at the top of [the script](Move%20file%20over%20SFTP.ps1))
- The machine running the SFTP client needs to have the module [Posh-SSH](https://github.com/darkoperator/Posh-SSH) installed