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
    [Boolean]$RemoveFailedPartialFiles,
    [Int]$RetryCountOnLockedFiles = 3,
    [Int]$RetryWaitSeconds = 3,
    [hashtable]$PartialFileExtension = @{
        Upload   = 'UploadInProgress'
        Download = 'DownloadInProgress'
    }
)

#region Set defaults
$ErrorActionPreference = 'Stop'

# workaround for https://github.com/PowerShell/PowerShell/issues/16894
$ProgressPreference = 'SilentlyContinue'
#endregion

Function Open-SftpSessionHM {
    #region Create SFTP credential
    try {
        Write-Verbose 'Create SFTP credential'

        $params = @{
            TypeName     = 'System.Management.Automation.PSCredential'
            ArgumentList = $SftpUserName, $SftpPassword
        }
        $sftpCredential = New-Object @params

        Write-Verbose 'Start SFTP session'

        $params = @{
            ComputerName = $SftpComputerName
            Credential   = $sftpCredential
            AcceptKey    = $true
            Force        = $true
        }

        if ($SftpOpenSshKeyFile) {
            $params.KeyString = $SftpOpenSshKeyFile
        }

        $sftpSession = New-SFTPSession @params
    }
    catch {
        $errorMessage = "Failed creating an SFTP session to '$SftpComputerName': $_"
        $Error.RemoveAt(0)
        throw $errorMessage
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
}

#region Get files to upload
foreach (
    $path in
    $Paths.where({ $_.Destination -like 'sftp*' })
) {
    #region Test source folder exists
    if (-not (Test-Path -LiteralPath $path.Source -PathType 'Container')) {
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

    #region Get files to upload
    Write-Verbose "Get files in folder '$($path.Source)'"

    $filesToUpload = Get-ChildItem -LiteralPath $path.Source -File

    if ($FileExtensions) {
        Write-Verbose "Only include files with extension '$FileExtensions'"

        $filesToUpload = $filesToUpload.where(
            { $FileExtensions -contains $_.Extension }
        )
    }

    if (-not $filesToUpload) {
        Write-Verbose 'No files to upload'
        Continue
    }
    #endregion

    #region Remove partial files from the local folder
    if ($RemoveFailedPartialFiles) {
        foreach (
            $partialFile in
            $allFiles | Where-Object {
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
                    Uploaded   = $false
                    Action     = @()
                    Error      = $null
                }

                Write-Verbose "Remove failed uploaded partial file '$($partialFile.FullName)'"

                Remove-Item -LiteralPath $partialFile.FullName

                $result.Action = "removed failed uploaded partial file '$($partialFile.FullName)'"
            }
            catch {
                $result.Error = "Failed removing uploaded partial file: $_"
                Write-Warning $result.Error
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }
    }
    #endregion

    #region Only select the required files for upload
    try {
        $filesToUpload = $allFiles | Where-Object {
            $_.Name -notLike "*$PartialFileExtension"
        }

        if ($FileExtensions) {
            Write-Verbose "Only include files with extension '$FileExtensions'"
            $filesToUpload = $filesToUpload | Where-Object {
                $FileExtensions -contains $_.Extension
            }
        }
    }
    catch {
        $errorMessage = "Failed selecting the required files for upload: $_"
        $Error.RemoveAt(0)
        throw $errorMessage
    }
    #endregion
}
#endregion

$scriptBlock = {
    $ErrorActionPreference = 'Stop'

    $path = $_

    #region Declare variables for code running in parallel
    if (-not $MaxConcurrentJobs) {
        $ErrorActionPreference = $using:ErrorActionPreference
    }
    #endregion


    try {

    }
    catch {
        $action.Job.Results += [PSCustomObject]@{
            DateTime    = Get-Date
            Source      = $path.Source
            Destination = $path.Destination
            FileName    = $null
            FileLength  = $null
            Downloaded  = $false
            Uploaded    = $false
            Action      = $null
            Error       = "General error: $_"
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

Write-Verbose "Execute $($Paths.Count) SFTP jobs"

$Paths | ForEach-Object @foreachParams

Write-Verbose 'All tasks finished'
#endregion
