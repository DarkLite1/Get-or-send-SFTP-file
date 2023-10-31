﻿#Requires -Version 5.1
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

.PARAMETER RemoveFileAfterDownload
    When the file is correctly downloaded, remove it from the SFTP server.
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
        $M = "Download path '$Path' not found"
        if ($ErrorWhenPathIsNotFound) {
            throw $M
        }

        Write-Verbose $M

        $downloadPathItem = New-Item -Path $Path -ItemType Directory -ErrorAction 'Stop'
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

        $sftpFiles = Get-SFTPChildItem @sftpSessionParams
    }
    catch {
        throw "Failed retrieving the SFTP file list from '$SftpComputerName' in path '$SftpPath': $_"
    }
    #endregion

    foreach ($file in $sftpFiles) {
        try {
            Write-Verbose "Download file '$file'"

            $downloadResult = [PSCustomObject]@{
                DateTime   = Get-Date
                Path       = $file.FullName
                Downloaded = $false
                Action     = @()
                Error      = $null
            }   
    
            #region Download file from SFTP server
            $params = @{
                Path        = $file.FullName
                Destination = $downloadPathItem.FullName
            }
    
            if ($OverwriteFile) {
                $params.Force = $true
            }
    
            Get-SFTPItem @sessionParams @params

            $downloadResult.Action += 'file downloaded'
            $downloadResult.Downloaded = $true
            #endregion
    
            #region Remove file after download
            if ($RemoveFileAfterDownload) {
                $M = "Remove file '{0}' from SFTP server" -f $file.FullName 
                Write-Verbose $M

                Remove-SFTPItem @sessionParams -Path $file.FullName
    
                $downloadResult.Action += 'file removed'
            }
            #endregion
        }
        catch {
            $downloadResult.Error = $_
            Write-Warning $_
            $Error.RemoveAt(0)        
        }
        finally {
            $downloadResult
        }
    }
}
catch {
    [PSCustomObject]@{
        DateTime   = Get-Date
        Path       = $Path
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