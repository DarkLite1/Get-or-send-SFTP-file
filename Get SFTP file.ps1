#Requires -Version 5.1
#Requires -Modules Posh-SSH

<#
.SYNOPSIS
    Download files from an SFTP server.

.DESCRIPTION
    Download files from an SFTP server.

.PARAMETER Path
    Folder where the downloaded files will be saved.

.PARAMETER SftpPath
    Path on the SFTP server where the files are located.
    
.PARAMETER SftpComputerName
    The URL where the SFTP server can be reached.

.PARAMETER SftpUserName
    The user name used to authenticate to the SFTP server.

.PARAMETER SftpPassword
    The password used to authenticate to the SFTP server.

.PARAMETER OverwriteFile
    When a file that is being downloaded is already present with the same name
    it will be overwritten when OverwriteFile is TRUE.

.PARAMETER RemoveFileAfterDownload
    When the file is correctly downloaded, remove it from the SFTP server.

.PARAMETER ErrorWhenPathIsNotFound
    When ErrorWhenPathIsNotFound is TRUE and download folder does not exist, 
    an error is thrown. When ErrorWhenPathIsNotFound is FALSE and the download
    folder does not exist, the folder is created.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$Path,
    [Parameter(Mandatory)]
    [String]$SftpComputerName,
    [Parameter(Mandatory)]
    [String]$SftpPath,
    [Parameter(Mandatory)]
    [String]$SftpUserName,
    [Parameter(Mandatory)]
    [SecureString]$SftpPassword,
    [Boolean]$OverwriteFile,
    [Boolean]$RemoveFileAfterDownload,
    [Boolean]$ErrorWhenPathIsNotFound
)

try {
    Write-Verbose "Test download path '$Path'"
    $downloadPathItem = Get-Item -LiteralPath $Path -ErrorAction 'Ignore'
      
    #region Test Path exists
    if (-not $downloadPathItem) {
        $M = "Download folder '$Path' not found"
        if ($ErrorWhenPathIsNotFound) {
            throw $M
        }

        Write-Verbose $M

        try {
            $downloadPathItem = New-Item -Path $Path -ItemType Directory -ErrorAction 'Stop'
        }
        catch {
            throw "Failed creating download folder '$Path': $_"
        }
    }
    #endregion
    
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

    #region Get SFTP file list
    try {
        $M = "Get SFTP file list in path '{0}'" -f $SftpPath
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $sftpFiles = Get-SFTPChildItem @sessionParams
    }
    catch {
        throw "Failed retrieving the SFTP file list from '$SftpComputerName' in path '$SftpPath': $_"
    }
    #endregion

    foreach ($file in $sftpFiles) {
        try {
            $result = [PSCustomObject]@{
                DateTime   = Get-Date
                LocalPath  = $Path
                SftpPath   = $SftpPath
                FileName   = $file.Name
                Downloaded = $false
                Action     = @()
                Error      = $null
            }

            Write-Verbose "Download file '$($result.FileName)'"
    
            #region Download file from SFTP server
            try {
                $params = @{
                    Path        = $file.FullName
                    Destination = $downloadPathItem.FullName
                }
    
                if ($OverwriteFile) {
                    $params.Force = $true
                }
    
                Get-SFTPItem @sessionParams @params

                $result.Action += 'file downloaded'
                $result.Downloaded = $true
            }
            catch {
                $M = "Failed downloading file: $_"
                $Error.RemoveAt(0)
                throw $M
            }
            #endregion
    
            #region Remove file after download
            if ($RemoveFileAfterDownload) {
                try {
                    $M = "Remove file '{0}' from SFTP server" -f $file.FullName
                    Write-Verbose $M
    
                    Remove-SFTPItem @sessionParams -Path $file.FullName
        
                    $result.Action += 'file removed'    
                }
                catch {
                    $M = "Failed removing downloaded file: $_"
                    $Error.RemoveAt(0)
                    throw $M
                }
            }
            #endregion
        }
        catch {
            $result.Error = $_
            Write-Warning $_
            $Error.RemoveAt(0)        
        }
        finally {
            $result
        }
    }
}
catch {
    [PSCustomObject]@{
        DateTime   = Get-Date
        LocalPath  = $Path
        SftpPath   = $SftpPath
        FileName   = $null
        Downloaded = $false
        Action     = $null
        Error      = $_
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