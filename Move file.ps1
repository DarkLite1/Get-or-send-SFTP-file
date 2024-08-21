#Requires -Version 7
#Requires -Modules Posh-SSH

<#
.SYNOPSIS
    Move files to or from an SFTP server.

.DESCRIPTION
    Move files to or from an SFTP server.

    To avoid file locks:
    1. Rename the source file on the SFTP server
       from 'a.txt' to 'a.txt.PartialFileExtension'
        > when a file can't be renamed it is locked
        > then we wait a few seconds for an unlock and try again
    2. Download 'a.txt.PartialFileExtension' from the SFTP server
    3. Remove the file on the SFTP server
    4. Rename the file from 'a.txt.PartialFileExtension' to 'a.txt'
       in the download folder

.PARAMETER Paths
    Lost of source and destination folders.

.PARAMETER SftpComputerName
    The URL where the SFTP server can be reached.

.PARAMETER SftpPath
    Path to th folder on the SFTP server.

.PARAMETER SftpUserName
    The user name used to authenticate to the SFTP server.

.PARAMETER SftpPassword
    The password used to authenticate to the SFTP server.

.PARAMETER SftpOpenSshKeyFile
    The password used to authenticate to the SFTP server. This is an
    SSH private key file in the OpenSSH format converted to an array of strings.

.PARAMETER FileExtensions
    Only the files with a matching file extension will be downloaded. If blank,
    all files will be downloaded.

.PARAMETER OverwriteFile
    When a file that is being downloaded is already present with the same name
    it will be overwritten when OverwriteFile is TRUE.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$SftpComputerName,
    [Parameter(Mandatory)]
    [String]$SftpUserName,
    [Parameter(Mandatory)]
    [HashTable[]]$Paths,
    [Parameter(Mandatory)]
    [Int]$MaxConcurrentJobs,
    [SecureString]$SftpPassword,
    [String[]]$SftpOpenSshKeyFile,
    [String[]]$FileExtensions,
    [Boolean]$OverwriteFile,
    [Int]$RetryCountOnLockedFiles = 3,
    [Int]$RetryWaitSeconds = 3,
    [hashtable]$PartialFileExtension = @{
        Upload   = '.UploadInProgress'
        Download = '.DownloadInProgress'
    }
)

try {
    #region Set defaults
    # workaround for https://github.com/PowerShell/PowerShell/issues/16894
    $ProgressPreference = 'SilentlyContinue'
    $ErrorActionPreference = 'Stop'
    #endregion

    Function Open-SftpSessionHM {
        <#
        .SYNOPSIS
            Open an SFTP session to the SFTP server
        #>

        try {
            #region Create credential
            Write-Verbose 'Create SFTP credential'

            $params = @{
                TypeName     = 'System.Management.Automation.PSCredential'
                ArgumentList = $SftpUserName, $SftpPassword
            }
            $sftpCredential = New-Object @params
            #endregion

            #region Open SFTP session
            Write-Verbose 'Open SFTP session'

            $params = @{
                ComputerName = $SftpComputerName
                Credential   = $sftpCredential
                AcceptKey    = $true
                Force        = $true
            }

            if ($SftpOpenSshKeyFile) {
                $params.KeyString = $SftpOpenSshKeyFile
            }

            New-SFTPSession @params
            #endregion
        }
        catch {
            $M = "Failed creating an SFTP session to '$SftpComputerName': $_"
            $Error.RemoveAt(0)
            throw $M
        }
    }

    $downloadPaths, $uploadPaths = $Paths.where(
        { $_.Source -like 'sftp*' }, 'Split'
    )

    if ($downloadPaths) {
        Write-Verbose "Paths.Source contains $($downloadPaths.Count) download folder(s)"

        $sftpSession = Open-SftpSessionHM
    }

    if ($uploadPaths) {
        Write-Verbose "Paths.Source contains $($uploadPaths.Count) upload folder(s)"

        $pathsWithFilesToUpload = @()

        foreach ($path in $uploadPaths) {
            Write-Verbose "Source folder '$($path.Source)'"

            #region Test source folder exists
            Write-Verbose 'Test if source folder exists'

            if (-not (
                    Test-Path -LiteralPath $path.Source -PathType 'Container')
            ) {
                [PSCustomObject]@{
                    Source      = $path.Source
                    Destination = $path.Destination
                    FileName    = $null
                    FileLength  = $null
                    DateTime    = Get-Date
                    Action      = @()
                    Error       = "Source folder '$($path.Source)' not found"
                }

                Continue
            }
            #endregion

            #region Test if there are files to upload
            Write-Verbose 'Test if there are files in the source folder'

            $filesToUpload = Get-ChildItem -LiteralPath $path.Source -File

            if ($FileExtensions) {
                $filesToUpload = $filesToUpload | Where-Object {
                    $FileExtensions -contains $_.Extension
                }
            }

            if ($filesToUpload) {
                Write-Verbose "Found $($filesToUpload.Count) file(s) to upload"
                $pathsWithFilesToUpload += $path
            }
            #endregion
        }

        if (-not $pathsWithFilesToUpload) {
            Write-Verbose 'No files in source folder'
            Write-Verbose 'Exit script'
            exit
        }

        if (-not $sftpSession) {
            $sftpSession = Open-SftpSessionHM
        }

        $scriptBlock = {
            try {
                $path = $_

                Write-Verbose "Source '$($path.Source)' Destination '$($path.Destination)'"

                #region Declare variables for code running in parallel
                if (-not $MaxConcurrentJobs) {
                    $ErrorActionPreference = $using:ErrorActionPreference
                    $ProgressPreference = $using:ProgressPreference
                    $sftpSession = $using:sftpSession
                    $FileExtensions = $using:FileExtensions
                    $PartialFileExtension = $using:PartialFileExtension
                    $RetryCountOnLockedFiles = $using:RetryCountOnLockedFiles
                    $RetryWaitSeconds = $using:RetryWaitSeconds
                    $OverwriteFile = $using:OverwriteFile
                }
                #endregion

                #region Get files to upload
                Write-Verbose 'Get files in source folder'

                $filesToUpload = Get-ChildItem -LiteralPath $path.Source -File

                if ($FileExtensions) {
                    Write-Verbose "Select files with extension '$FileExtensions'"

                    $filesToUpload = $filesToUpload | Where-Object {
                        $FileExtensions -contains $_.Extension
                    }
                }

                if (-not $filesToUpload) {
                    Write-Verbose 'No files to upload'
                    Continue
                }

                Write-Verbose "Found $($filesToUpload.Count) file(s) to upload"
                #endregion

                $sessionParams = @{
                    SessionId = $sftpSession.SessionID
                }

                $sftpPath = $path.Destination.TrimStart('sftp:')

                #region Test SFTP path exists
                Write-Verbose "Test if SFTP path '$sftpPath' exists"

                if (-not (Test-SFTPPath @sessionParams -Path $sftpPath)) {
                    throw "Path '$sftpPath' not found on the SFTP server"
                }
                #endregion

                #region Get all SFTP files
                try {
                    $sftpFiles = Get-SFTPChildItem @sessionParams -Path $SftpPath -File
                }
                catch {
                    $errorMessage = "Failed retrieving SFTP files: $_"
                    $Error.RemoveAt(0)
                    throw $errorMessage
                }
                #endregion

                #region Remove incomplete uploaded files from the SFTP server
                foreach (
                    $partialFile in
                    $sftpFiles.where(
                        { $_.Name -like "*$($PartialFileExtension.Upload)" }
                    )
                ) {
                    try {
                        $result = [PSCustomObject]@{
                            DateTime    = Get-Date
                            Source      = $path.Source
                            Destination = $path.Destination
                            FileName    = $partialFile.Name
                            FileLength  = $partialFile.Length
                            Action      = @()
                            Error       = $null
                        }

                        Write-Verbose "Remove incompletely uploaded file '$($partialFile.FullName)'"

                        Remove-SFTPItem @sessionParams -Path $partialFile.FullName

                        $result.Action = 'Removed incompletely uploaded file'
                    }
                    catch {
                        $result.Error = "Failed removing incomplete uploaded file: $_"
                        Write-Warning $result.Error
                        $Error.RemoveAt(0)
                    }
                    finally {
                        $result
                    }
                }
                #endregion

                foreach ($file in $filesToUpload) {
                    try {
                        Write-Verbose "File '$($file.FullName)'"

                        $result = [PSCustomObject]@{
                            DateTime    = Get-Date
                            Source      = $path.Source
                            Destination = $path.Destination
                            FileName    = $file.Name
                            FileLength  = $file.Length
                            Action      = @()
                            Error       = $null
                        }

                        $tempFile = @{
                            UploadFileName = $file.Name + $PartialFileExtension.Upload
                        }
                        $tempFile.UploadFilePath = Join-Path $result.Source $tempFile.UploadFileName

                        #region Duplicate file on SFTP server
                        if (
                            $sftpFile = $sftpFiles.where(
                                { $_.Name -eq $file.Name }, 'First'
                            )
                        ) {
                            Write-Verbose 'Duplicate file on SFTP server'

                            if ($OverwriteFile) {
                                $retryCount = 0
                                $fileLocked = $true

                                while (
                                    ($fileLocked) -and
                                    ($retryCount -lt $RetryCountOnLockedFiles)
                                ) {
                                    try {
                                        Write-Verbose 'Remove duplicate file on SFTP server'

                                        $removeParams = @{
                                            Path        = $sftpFile.FullName
                                            ErrorAction = 'Stop'
                                        }
                                        Remove-SFTPItem @sessionParams @removeParams

                                        $fileLocked = $false

                                        [PSCustomObject]@{
                                            DateTime    = $result.DateTime.AddSeconds(-1)
                                            Source      = $result.Source
                                            Destination = $result.Destination
                                            FileName    = $result.Name
                                            FileLength  = $result.Length
                                            Action      = @('Removed duplicate file from SFTP server')
                                            Error       = $null
                                        }
                                    }
                                    catch {
                                        $errorMessage = $_
                                        $Error.RemoveAt(0)
                                        $retryCount++
                                        Write-Warning "File locked, wait $RetryWaitSeconds seconds, attempt $retryCount/$RetryCountOnLockedFiles"
                                        Start-Sleep -Seconds $RetryWaitSeconds
                                    }
                                }

                                if ($fileLocked) {
                                    throw "Failed removing duplicate file from the SFTP server after multiple attempts within $($RetryCountOnLockedFiles * $RetryWaitSeconds) seconds (file in use): $errorMessage"
                                }
                            }
                            else {
                                throw 'Duplicate file on SFTP server, use Option.OverwriteFile if desired'
                            }
                        }
                        #endregion

                        #region Rename source file to temp file
                        $retryCount = 0
                        $fileLocked = $true

                        while (
                            ($fileLocked) -and
                            ($retryCount -lt $RetryCountOnLockedFiles)
                        ) {
                            try {
                                Write-Verbose "Rename source file to temp file '$($tempFile.UploadFileName)'"
                                $file |
                                Rename-Item -NewName $tempFile.UploadFileName
                                $fileLocked = $false
                            }
                            catch {
                                $errorMessage = $_
                                $Error.RemoveAt(0)
                                $retryCount++
                                Write-Warning "File locked, wait $RetryWaitSeconds seconds, attempt $retryCount/$RetryCountOnLockedFiles"
                                Start-Sleep -Seconds $RetryWaitSeconds
                            }
                        }

                        if ($fileLocked) {
                            throw "Failed renaming the source file after multiple attempts within $($RetryCountOnLockedFiles * $RetryWaitSeconds) seconds (file in use): $errorMessage"
                        }
                        #endregion

                        #region Upload temp file to SFTP server
                        try {
                            $params = @{
                                Path        = $tempFile.UploadFilePath
                                Destination = $SftpPath
                            }

                            Write-Verbose 'Upload temp file'
                            Set-SFTPItem @sessionParams @params
                        }
                        catch {
                            $errorMessage = "Failed to upload file '$($tempFile.UploadFilePath)': $_"
                            $Error.RemoveAt(0)
                            throw $errorMessage
                        }
                        #endregion

                        #region Rename file on SFTP server
                        try {
                            Write-Verbose "Rename temp file on SFTP server to '$($result.FileName)'"

                            $params = @{
                                Path    = $SftpPath + $tempFile.UploadFileName
                                NewName = $result.FileName
                            }
                            Rename-SFTPFile @sessionParams @params
                        }
                        catch {
                            $errorMessage = "Failed to rename the file on the SFTP server from '$($tempFile.UploadFileName)' to '$($result.FileName)': $_"
                            $Error.RemoveAt(0)
                            throw $errorMessage
                        }
                        #endregion

                        #region Remove local temp file
                        try {
                            Write-Verbose 'Remove local temp file'

                            $tempFile.UploadFilePath | Remove-Item -Force
                        }
                        catch {
                            $errorMessage = "Failed to remove the local temp file '$($tempFile.UploadFilePath)': $_"
                            $Error.RemoveAt(0)
                            throw $errorMessage
                        }
                        #endregion

                        $result.Action += 'File moved'
                    }
                    catch {
                        #region Rename temp file back to original file name
                        if (
                            Test-Path -LiteralPath $tempFile.UploadFilePath -PathType 'Leaf'
                        ) {
                            try {
                                Write-Warning 'Upload failed'
                                Write-Verbose "Rename temp file '$($tempFile.UploadFilePath)' back to its original name '$($file.Name)'"

                                $tempFile.UploadFilePath |
                                Rename-Item -NewName $file.Name
                            }
                            catch {
                                [PSCustomObject]@{
                                    DateTime    = Get-Date
                                    Source      = $result.Source
                                    Destination = $result.Destination
                                    FileName    = $tempFile.Name
                                    FileLength  = $result.Length
                                    Action      = @()
                                    Error       = "Failed to rename temp file '$($tempFile.UploadFilePath)' back to its original name '$($file.Name)': $_"
                                }

                                $Error.RemoveAt(0)
                            }
                        }
                        #endregion

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
                $M = "Failed upload: $_"
                Write-Warning $M

                [PSCustomObject]@{
                    DateTime    = Get-Date
                    Source      = $path.Source
                    Destination = $path.Destination
                    FileName    = $null
                    FileLength  = $null
                    Action      = @()
                    Error       = $M
                }
                $Error.RemoveAt(0)
            }
        }

        #region Run code serial or parallel
        $foreachParams = if ($MaxConcurrentJobs -eq 1) {
            @{
                Process = $scriptBlock
            }
        }
        else {
            @{
                Parallel      = $scriptBlock
                ThrottleLimit = $MaxConcurrentJobs
            }
        }

        Write-Verbose "Found $($uploadPaths.Count) source path(s) with files to upload"

        $pathsWithFilesToUpload | ForEach-Object @foreachParams

        Write-Verbose 'All upload jobs finished'
        #endregion
    }
}
catch {
    $M = $_
    Write-Warning $M
    $Error.RemoveAt(0)
    throw $M
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
