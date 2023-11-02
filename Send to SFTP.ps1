#Requires -Version 5.1
#Requires -Modules Posh-SSH

<#
.SYNOPSIS
    Upload files to an SFTP server.

.DESCRIPTION
    Send one or more files to an SFTP server.

.PARAMETER Path
    Full path to the files to upload or to the folder containing the files to 
    upload. Only files are uploaded subfolders are not.
    
.PARAMETER SftpComputerName
    The URL where the SFTP server can be reached.

.PARAMETER SftpPath
    Path on the SFTP server where the uploaded files will be saved.

.PARAMETER SftpUserName
    The user name used to authenticate to the SFTP server.

.PARAMETER SftpPassword
    The password used to authenticate to the SFTP server.

.PARAMETER OverwriteFileOnSftpServer
    Overwrite the file on the SFTP server in case it already exists.

.PARAMETER RemoveFileAfterUpload
    Remove the source file after a successful upload.

.PARAMETER ErrorWhenUploadPathIsNotFound
    Create an error in the returned object when the SFTP path is not found 
    on the SFTP server.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String[]]$Path,
    [Parameter(Mandatory)]
    [String]$SftpComputerName,
    [Parameter(Mandatory)]
    [String]$SftpPath,
    [Parameter(Mandatory)]
    [String]$SftpUserName,
    [Parameter(Mandatory)]
    [SecureString]$SftpPassword,
    [Boolean]$OverwriteFileOnSftpServer,
    [Boolean]$RemoveFileAfterUpload,
    [Boolean]$ErrorWhenUploadPathIsNotFound
)

try {
    #region Get files to upload
    $filesToUpload = @()

    foreach ($P in $Path) {
        try {
            Write-Verbose "Test path '$P'"
            $item = Get-Item -LiteralPath $P -ErrorAction 'Ignore'
      
            #region Test Path exists
            if (-not $item) {
                $M = "Path '$P' not found"
                if ($ErrorWhenUploadPathIsNotFound) {
                    throw $M
                }

                Write-Verbose $M
                Continue
            }
            #endregion

            #region Get files
            $filesToUpload += if ($item.PSIsContainer) {
                Write-Verbose "Get files in folder '$P'"
                Get-ChildItem -LiteralPath $item.FullName -File -ErrorAction 'Stop'
            }
            else {
                $item
            }
            #endregion
        }
        catch {
            [PSCustomObject]@{
                DateTime  = Get-Date
                LocalPath = $P
                SftpPath  = $SftpPath
                FileName  = $null
                Uploaded  = $false
                Action    = $null
                Error     = $_
            }
            Write-Warning $_
            $Error.RemoveAt(0)        
        }
    }
    #endregion
    
    if (-not $filesToUpload) {
        Write-Verbose 'No files to upload'
        Exit
    }

    #region Create SFTP credential
    try {
        Write-Verbose 'Create SFTP credential'
    
        $params = @{
            TypeName     = 'System.Management.Automation.PSCredential'
            ArgumentList = $SftpUserName, $SftpPassword
            ErrorAction  = 'Stop'
        }
        $sftpCredential = New-Object @params
    }
    catch {
        throw "Failed creating the SFTP credential with user name '$($SftpUserName)' and password '$SftpPassword': $_"
    }
    #endregion
        
    #region Open SFTP session
    try {
        Write-Verbose 'Start SFTP session'

        $params = @{
            ComputerName = $SftpComputerName
            Credential   = $sftpCredential
            AcceptKey    = $true
            ErrorAction  = 'Stop'
        }
        $sftpSession = New-SFTPSession @params
    }
    catch {
        throw "Failed creating an SFTP session to '$SftpComputerName': $_"
    }
    #endregion

    $sessionParams = @{
        SessionId   = $sftpSession.SessionID
        ErrorAction = 'Stop'
    }

    #region Test SFTP path exists
    Write-Verbose "Test SFTP path '$SftpPath' exists"
        
    if (-not (Test-SFTPPath @sessionParams -Path $SftpPath)) {
        throw "Path '$SftpPath' not found on SFTP server"
    }    
    #endregion

    foreach ($file in $filesToUpload) {
        try {
            Write-Verbose "Upload file '$($file.FullName)'"

            $uploadResult = [PSCustomObject]@{
                DateTime  = Get-Date
                LocalPath = $file.FullName | Split-Path -Parent
                SftpPath  = $SftpPath
                FileName  = $file.Name
                Uploaded  = $false
                Action    = @()
                Error     = $null
            }   
    
            #region Upload data to SFTP server
            $params = @{
                Path        = $file.FullName
                Destination = $SftpPath
            }
    
            if ($OverwriteFileOnSftpServer) {
                $params.Force = $true
            }
    
            Set-SFTPItem @sessionParams @params

            $uploadResult.Action += 'file uploaded'
            $uploadResult.Uploaded = $true
            #endregion
    
            #region Remove source file
            if ($RemoveFileAfterUpload) {
                $file | Remove-Item -Force -ErrorAction 'Stop'
    
                $uploadResult.Action += 'file removed'
            }
            #endregion
        }
        catch {
            $uploadResult.Error = $_
            Write-Warning $_
            $Error.RemoveAt(0)        
        }
        finally {
            $uploadResult
        }
    }
}
catch {
    [PSCustomObject]@{
        DateTime  = Get-Date
        LocalPath = $Path
        SftpPath  = $SftpPath
        FileName  = $null
        Uploaded  = $false
        Action    = $null
        Error     = "General error: $_"
    }
    Write-Warning $_
    $Error.RemoveAt(0)
}
finally {
    #region Close SFTP session
    if ($sessionParams.SessionID) {
        Write-Verbose 'Close SFTP session'

        $params = @{
            SessionId   = $sessionParams.SessionID
            ErrorAction = 'Ignore'
        }
        $null = Remove-SFTPSession @params
    }
    #endregion
}