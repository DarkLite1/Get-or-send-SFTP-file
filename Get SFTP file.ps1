#Requires -Version 5.1
#Requires -Modules Posh-SSH

<#
.SYNOPSIS
    Download files from an SFTP server.

.DESCRIPTION
    Download one or more files from an SFTP server. Downloaded files will
    always be removed from the SFTP server. Simply  because renaming the file
    on the SFTP server, to avoid file locks, might otherwise interfere with
    the creation of new files with the same name on the SFTP server.

    To avoid file locks:
    1. Rename the source file on the SFTP server
       from 'a.txt' to 'a.txt.PartialFileExtension'
        > when a file can't be renamed it is locked
        > then we wait a few seconds for an unlock and try again
    2. Download 'a.txt.PartialFileExtension' from the SFTP server
    3. Remove the file on the SFTP server
    4. Rename the file from 'a.txt.PartialFileExtension' to 'a.txt'
       in the download folder

.PARAMETER Path
    Full path to the folder where the downloaded files will be saved.

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

.PARAMETER PartialFileExtension
    The name used for the file extension of the partial file that is being
    downloaded. The file that needs to be downloaded is first renamed
    by adding another file extension. This will make sure that errors like
    "file in use by another process" are avoided.

    After a rename the file is downloaded with the extension defined in
    "PartialFileExtension". After a successful download the file is then
    removed from the SFTP server and renamed to its original name in the
    download folder.

.PARAMETER FileExtensions
    Only the files with a matching file extension will be downloaded. If blank,
    all files will be downloaded.

.PARAMETER OverwriteFile
    When a file that is being downloaded is already present with the same name
    it will be overwritten when OverwriteFile is TRUE.

.PARAMETER RemoveFailedPartialFiles
    When the download process is interrupted, it is possible that files are not
    completely downloaded and that there are sill partial files present on the
    SFTP server or in the local folder.

    When RemoveFailedPartialFiles is TRUE these partial files will be removed
    before the script starts. When RemoveFailedPartialFiles is FALSE, manual
    intervention will be required to decide to still download the partial file
    found on the SFTP server, to rename the partial file on the local system,
    or to simply remove the partial file(s).
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
    [String]$PartialFileExtension,
    [SecureString]$SftpPassword,
    [String[]]$SftpOpenSshKeyFile,
    [String[]]$FileExtensions,
    [Boolean]$OverwriteFile,
    [Boolean]$RemoveFailedPartialFiles,
    [Int]$RetryCountOnLockedFiles = 3,
    [Int]$RetryWaitSeconds = 3
)

try {
    $ErrorActionPreference = 'Stop'

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

    #region Test download folder exists
    Write-Verbose "Test download folder '$Path' exists"

    if (-not (Test-Path -LiteralPath $Path -PathType 'Container')) {
        throw "Path '$Path' not found"
    }
    #endregion

    #region Get list of files on the SFTP server
    try {
        Write-Verbose "Get list of files from SFTP server path '$SftpPath'"

        $allFiles = Get-SFTPChildItem @sessionParams -Path $SftpPath -File
    }
    catch {
        [PSCustomObject]@{
            DateTime   = Get-Date
            LocalPath  = $P
            SftpPath   = $SftpPath
            FileName   = $null
            FileLength = $null
            Downloaded = $false
            Action     = $null
            Error      = $_
        }
        Write-Warning $_
        $Error.RemoveAt(0)
    }
    #endregion

    #region Remove partial files that failed downloading
    if ($RemoveFailedPartialFiles) {
        #region From the SFTP server
        foreach (
            $partialFile in
            $allFiles | Where-Object { $_.Name -like "*$PartialFileExtension" }
        ) {
            try {
                $result = [PSCustomObject]@{
                    DateTime   = Get-Date
                    LocalPath  = $null
                    SftpPath   = $SftpPath
                    FileName   = $partialFile.Name
                    FileLength = $partialFile.Length
                    Downloaded = $false
                    Action     = @()
                    Error      = $null
                }

                Write-Verbose "Remove failed downloaded partial file '$($partialFile.FullName)'"

                Remove-SFTPItem @sessionParams -Path $partialFile.FullName

                $result.Action = "removed failed downloaded partial file '$($partialFile.FullName)'"
            }
            catch {
                $result.Error = "Failed removing downloaded partial file: $_"
                Write-Warning $result.Error
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }
        #endregion

        #region From the download folder
        foreach (
            $partialFile in
            Get-ChildItem -LiteralPath $Path -File | Where-Object {
                $_.Name -like "*$PartialFileExtension"
            }
        ) {
            try {
                $result = [PSCustomObject]@{
                    DateTime   = Get-Date
                    LocalPath  = $null
                    SftpPath   = $SftpPath
                    FileName   = $partialFile.Name
                    FileLength = $partialFile.Length
                    Downloaded = $false
                    Action     = @()
                    Error      = $null
                }

                Write-Verbose "Remove failed downloaded partial file '$($partialFile.FullName)'"

                Remove-Item -LiteralPath $partialFile.FullName

                $result.Action = "removed failed downloaded partial file '$($partialFile.FullName)'"
            }
            catch {
                $result.Error = "Failed removing downloaded partial file: $_"
                Write-Warning $result.Error
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }
        #endregion
    }
    #endregion

    #region Only select the required files for download
    $filesToDownload = $allFiles | Where-Object {
        $_.Name -notmatch "$PartialFileExtension$"
    }


    if ($FileExtensions) {
        Write-Verbose "Only include files with extension '$FileExtensions'"
        $fileExtensionFilter = (
            $FileExtensions | ForEach-Object { "$_$" }
        ) -join '|'

        $filesToDownload = $filesToDownload | Where-Object {
            $_.Name -match $fileExtensionFilter
        }
    }
    #endregion

    if (-not $filesToDownload) {
        Write-Verbose 'No files to download'
        Exit
    }

    foreach ($file in $filesToDownload) {
        try {
            Write-Verbose "File '$($file.FullName)'"

            $result = [PSCustomObject]@{
                DateTime   = Get-Date
                LocalPath  = $Path
                SftpPath   = $SftpPath
                FileName   = $file.Name
                FileLength = $file.Length
                Downloaded = $false
                Action     = @()
                Error      = $null
            }

            $tempFile = @{
                DownloadFileName = $file.Name + $PartialFileExtension
                DownloadFilePath = $SftpPath + $file.Name + $PartialFileExtension
            }

            #region Rename source file to temp file on SFTP server
            $retryCount = 0
            $fileLocked = $true

            while (
                ($fileLocked) -and
                ($retryCount -lt $RetryCountOnLockedFiles)
            ) {
                try {
                    $params = @{
                        Path    = $file.FullName
                        NewName = $tempFile.DownloadFileName
                    }
                    Write-Verbose "rename file to '$($params.NewName)'"
                    Rename-SFTPFile @sessionParams @params

                    $fileLocked = $false
                    $result.Action += 'temp file created'
                }
                catch {
                    $retryCount++
                    Write-Warning "File locked, wait $RetryWaitSeconds seconds, attempt $retryCount/$RetryCountOnLockedFiles"
                    Start-Sleep -Seconds $RetryWaitSeconds
                }
            }

            if ($fileLocked) {
                throw "File in use on the SFTP server by another process. Waited for $($RetryCountOnLockedFiles * $RetryWaitSeconds) seconds without success."
            }
            #endregion

            #region download temp file from the SFTP server
            try {
                $params = @{
                    Path        = $tempFile.DownloadFilePath
                    Destination = $Path
                }

                if ($OverwriteFile) {
                    Write-Verbose 'Overwrite file on on the local file system'
                    $params.Force = $true
                }

                Write-Verbose 'download temp file'
                Get-SFTPItem @sessionParams @params

                $result.Action += 'temp file downloaded'
            }
            catch {
                throw "Failed to download file '$($tempFile.DownloadFilePath)': $_"
            }
            #endregion

            #region Remove file
            try {
                Write-Verbose 'Remove temp file'

                Remove-SFTPItem @sessionParams -Path $tempFile.DownloadFilePath
                $result.Action += 'temp file removed'
            }
            catch {
                throw "Failed to remove file '$($tempFile.DownloadFilePath)': $_"
            }
            #endregion

            #region Rename file
            try {
                $params = @{
                    LiteralPath = Join-Path $Path $tempFile.DownloadFileName
                    NewName     = $result.FileName
                }
                Write-Verbose 'Rename temp file'
                Rename-Item @params

                $result.Action += 'temp file renamed'
            }
            catch {
                throw "Failed to rename the file '$($params.LiteralPath)' to '$($result.FileName)': $_"
            }
            #endregion

            $result.Action += 'file successfully downloaded'
            $result.Downloaded = $true

            Write-Verbose 'file successfully downloaded'
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
        FileLength = $null
        Downloaded = $false
        Action     = $null
        Error      = "General error: $_"
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