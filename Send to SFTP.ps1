#Requires -Version 5.1
#Requires -Modules Posh-SSH

<#
.SYNOPSIS
    Upload files to an SFTP server.

.DESCRIPTION
    Send one or more files to an SFTP server. After a successful upload the 
    source file is always removed.

    To avoid file locks:
    1. Rename the source file 'a.txt' to 'a.txt.UploadInProgress'
        > when a file can't be renamed it is locked
        > then we wait a few seconds for an unlock and try again
    2. Upload 'a.txt.UploadInProgress' to the SFTP server
    3. On the SFTP server rename 'a.txt.UploadInProgress' to 'a.txt'

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
    [Boolean]$ErrorWhenUploadPathIsNotFound,
    [Int]$RetryCountOnLockedFiles = 3,
    [Int]$RetryWaitSeconds = 3
)

try {
    $ErrorActionPreference = 'Stop'

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
                Get-ChildItem -LiteralPath $item.FullName -File
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
        }
        $sftpSession = New-SFTPSession @params
    }
    catch {
        throw "Failed creating an SFTP session to '$SftpComputerName': $_"
    }
    #endregion

    $sessionParams = @{
        SessionId = $sftpSession.SessionID
    }

    #region Test SFTP path exists
    Write-Verbose "Test SFTP path '$SftpPath' exists"
        
    if (-not (Test-SFTPPath @sessionParams -Path $SftpPath)) {
        throw "Path '$SftpPath' not found on SFTP server"
    }    
    #endregion

    foreach ($file in $filesToUpload) {
        try {
            Write-Verbose "File '$($file.FullName)'"

            $result = [PSCustomObject]@{
                DateTime  = Get-Date
                LocalPath = $file.FullName | Split-Path -Parent
                SftpPath  = $SftpPath
                FileName  = $file.Name
                Uploaded  = $false
                Action    = @()
                Error     = $null
            }

            $tempFile = @{
                UploadFileName = $file.Name -Replace "\$($file.Extension)", "$($file.Extension).UploadInProgress" 
            }
            $tempFile.UploadFilePath = Join-Path $result.LocalPath $tempFile.UploadFileName

            #region Rename source file to temp file
            $retryCount = 0
            $fileLocked = $true

            while (
                ($fileLocked) -and
                ($retryCount -lt $RetryCountOnLockedFiles)
            ) {
                try {
                    Write-Verbose 'Rename source file'
                    $file | Rename-Item -NewName $tempFile.UploadFileName
                    $fileLocked = $false
                    $result.Action += 'file renamed'
                }
                catch {
                    $retryCount++
                    Write-Warning "File locked, wait $RetryWaitSeconds seconds, attempt $retryCount/$RetryCountOnLockedFiles"
                    Start-Sleep -Seconds $RetryWaitSeconds
                }
            }

            if ($fileLocked) {
                throw "File in use by another process. Waited for $($RetryCountOnLockedFiles * $RetryWaitSeconds) seconds without success."
            }
            #endregion
    
            #region Upload temp file to SFTP server
            try {
                $params = @{
                    Path        = $tempFile.UploadFilePath
                    Destination = $SftpPath
                }
                
                if ($OverwriteFileOnSftpServer) {
                    Write-Verbose 'Overwrite file on SFTP server'
                    $params.Force = $true
                }
                
                Write-Verbose "Upload file '$($params.Path)'"
                Set-SFTPItem @sessionParams @params
    
                Write-Verbose 'File uploaded'
                $result.Action += 'file uploaded'
                $result.Uploaded = $true    
            }
            catch {
                throw "Failed to upload file '$($tempFile.UploadFilePath)': $_"
            }
            #endregion
    
            #region Remove file
            try {
                Write-Verbose 'Remove file'

                $tempFile.UploadFilePath | Remove-Item -Force
                $result.Action += 'file removed'    
            }
            catch {
                throw "Failed to remove file '$($tempFile.UploadFilePath)': $_"
            }
            #endregion
            
            #region Rename file on SFTP server
            try {
                $params = @{
                    Path    = $SftpPath + $tempFile.UploadFileName
                    NewName = $result.FileName
                }
                Rename-SFTPFile @sessionParams @params
    
                $result.Action += 'file renamed on SFTP server'    
            }
            catch {
                throw "Failed to rename the file '$($tempFile.UploadFileName)' to '$($result.FileName)' on the SFTP server: $_"
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