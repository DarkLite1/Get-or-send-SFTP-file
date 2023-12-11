#Requires -Version 5.1
#Requires -Modules Posh-SSH

<#
.SYNOPSIS
    Upload files to an SFTP server.

.DESCRIPTION
    Send one or more files to an SFTP server. After a successful upload the
    source file is always removed.

    To avoid file locks:
    1. Rename the source file 'a.txt' to 'a.txt.PartialFileExtension'
        > when a file can't be renamed it is locked
        > then we wait a few seconds for an unlock and try again
    2. Upload 'a.txt.PartialFileExtension' to the SFTP server
    3. On the SFTP server rename 'a.txt.PartialFileExtension' to 'a.txt'

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

.PARAMETER SftpOpenSshKeyFile
    The password used to authenticate to the SFTP server. This is an
    SSH private key file in the OpenSSH format converted to an array of strings.

.PARAMETER PartialFileExtension
    The name used for the file extension of the partial file that is being
    uploaded. The file that needs to be uploaded is first renamed
    by adding another file extension. This will make sure that errors like
    "file in use by another process" are avoided.

    After a rename the file is uploaded with the extension defined in
    "PartialFileExtension". After a successful upload the file is then renamed
    on the SFTP server to its original name with the correct file extension.

.PARAMETER FileExtensions
    Only the files with a matching file extension will be uploaded. If blank,
    all files will be uploaded.

.PARAMETER OverwriteFileOnSftpServer
    Overwrite the file on the SFTP server in case it already exists.

.PARAMETER ErrorWhenUploadPathIsNotFound
    Create an error in the returned object when the SFTP path is not found
    on the SFTP server.

.PARAMETER RemoveFailedPartialFiles
    When the upload process is interrupted, it is possible that files are not
    completely uploaded and that there are sill partial files present on the
    SFTP server or in the local folder.

    When RemoveFailedPartialFiles is TRUE these partial files will be removed
    before the script starts. When RemoveFailedPartialFiles is FALSE, manual
    intervention will be required to decide to still upload the partial file
    found in the local folder, to rename the partial file on the SFTP server,
    or to simply remove the partial file(s).
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
    [String]$PartialFileExtension,
    [SecureString]$SftpPassword,
    [String[]]$SftpOpenSshKeyFile,
    [Boolean]$OverwriteFileOnSftpServer,
    [Boolean]$ErrorWhenUploadPathIsNotFound,
    [Boolean]$RemoveFailedPartialFiles,
    [String[]]$FileExtensions,
    [Int]$RetryCountOnLockedFiles = 3,
    [Int]$RetryWaitSeconds = 3
)

try {
    $ErrorActionPreference = 'Stop'

    # workaround for https://github.com/PowerShell/PowerShell/issues/16894
    $ProgressPreference = 'SilentlyContinue'

    #region Get files to upload
    $allFiles = @()

    foreach ($P in $Path) {
        try {
            Write-Verbose "Test path '$P'"
            $item = Get-Item -LiteralPath $P -ErrorAction 'Ignore'

            #region Test local path exists
            if (-not $item) {
                $M = "Path '$P' not found"
                if ($ErrorWhenUploadPathIsNotFound) {
                    throw $M
                }

                Write-Verbose $M
                Continue
            }
            #endregion

            #region Get local files
            if ($item.PSIsContainer) {
                Write-Verbose "Get files in folder '$P'"

                $allFiles += Get-ChildItem -LiteralPath $item.FullName -File

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
            }
            else {
                $allFiles += $item

                #region Remove partial file
                $params = @{
                    LiteralPath = "$($item.FullName)$PartialFileExtension"
                    ErrorAction = 'Ignore'
                }
                if (
                    ($RemoveFailedPartialFiles) -and
                    ($partialFile = Get-Item @params)
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
                #endregion
            }
            #endregion
        }
        catch {
            [PSCustomObject]@{
                DateTime   = Get-Date
                LocalPath  = $P
                SftpPath   = $SftpPath
                FileName   = $null
                FileLength = $null
                Uploaded   = $false
                Action     = $null
                Error      = $_
            }
            Write-Warning $_
            $Error.RemoveAt(0)
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
        $errorMessage = "Failed creating the SFTP credential with user name '$($SftpUserName)' and password '$SftpPassword': $_"
        $Error.RemoveAt(0)
        throw $errorMessage
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

    #region Remove partial files that failed uploading from the SFTP server
    if ($RemoveFailedPartialFiles) {
        foreach (
            $partialFile in
            $sftpFiles | Where-Object { $_.Name -like "*$PartialFileExtension" }
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

                Write-Verbose "Remove partial file '$($partialFile.FullName)'"

                Remove-SFTPItem @sessionParams -Path $partialFile.FullName

                $result.Action = "removed partial file '$($partialFile.FullName)'"
            }
            catch {
                $result.Error = "Failed removing partial file: $_"
                Write-Warning $result.Error
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }
    }
    #endregion

    foreach ($file in $filesToUpload) {
        try {
            Write-Verbose "File '$($file.FullName)'"

            $result = [PSCustomObject]@{
                DateTime   = Get-Date
                LocalPath  = $file.FullName | Split-Path -Parent
                SftpPath   = $SftpPath
                FileName   = $file.Name
                FileLength = $file.Length
                Uploaded   = $false
                Action     = @()
                Error      = $null
            }

            $tempFile = @{
                UploadFileName = $file.Name + $PartialFileExtension
            }
            $tempFile.UploadFilePath = Join-Path $result.LocalPath $tempFile.UploadFileName

            #region Overwrite file on SFTP server
            if (
                $sftpFile = $sftpFiles | Where-Object {
                    $_.Name -eq $file.Name
                }
            ) {
                Write-Verbose 'Duplicate file on SFTP server'

                if ($OverwriteFileOnSftpServer) {
                    $retryCount = 0
                    $fileLocked = $true

                    while (
                        ($fileLocked) -and
                        ($retryCount -lt $RetryCountOnLockedFiles)
                    ) {
                        try {
                            $removeParams = @{
                                Path        = $sftpFile.FullName
                                ErrorAction = 'Stop'
                            }
                            Remove-SFTPItem @sessionParams @removeParams

                            $fileLocked = $false
                            $result.Action += 'removed duplicate file from SFTP server'
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
                    Write-Verbose 'Rename source file'
                    $file | Rename-Item -NewName $tempFile.UploadFileName
                    $fileLocked = $false
                    $result.Action += 'temp file created'
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

                Write-Verbose "Upload file '$($params.Path)'"
                Set-SFTPItem @sessionParams @params

                Write-Verbose 'File uploaded'
                $result.Action += 'temp file uploaded'
            }
            catch {
                $errorMessage = "Failed to upload file '$($tempFile.UploadFilePath)': $_"
                $Error.RemoveAt(0)
                throw $errorMessage
            }
            #endregion

            #region Remove local file
            try {
                Write-Verbose 'Remove file'

                $tempFile.UploadFilePath | Remove-Item -Force
                $result.Action += 'temp file removed'
            }
            catch {
                $errorMessage = "Failed to remove file '$($tempFile.UploadFilePath)': $_"
                $Error.RemoveAt(0)
                throw $errorMessage
            }
            #endregion

            #region Rename file on SFTP server
            try {
                $params = @{
                    Path    = $SftpPath + $tempFile.UploadFileName
                    NewName = $result.FileName
                }
                Rename-SFTPFile @sessionParams @params

                $result.Action += 'temp file renamed on SFTP server'
            }
            catch {
                $errorMessage = "Failed to rename the file '$($tempFile.UploadFileName)' to '$($result.FileName)' on the SFTP server: $_"
                $Error.RemoveAt(0)
                throw $errorMessage
            }
            #endregion

            $result.Action += 'file successfully uploaded'
            $result.Uploaded = $true
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
        Uploaded   = $false
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