#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog
#Requires -Modules Posh-SSH

<#
.SYNOPSIS
    Upload files to an SFTP server.

.DESCRIPTION
    Upload all files in a folder or a single file to an SFTP server.

.PARAMETER ImportFile
    A .JSON file that contains all the parameters used by the script.

.PARAMETER MailTo
    E-mail addresses of where to send the summary e-mail

.PARAMETER Upload.Type
    Defines which files need to be uploaded to the SFTP server.
    Valid values:
    - File         : a single file
    - Folder       : the complete content of a folder
    - FolderContent: all files in a folder, not recursive

.PARAMETER Upload.Path
    The path to the file or folder.

.PARAMETER Option.OverwriteExistingFile
    Overwrite a file when a file with the same name already exists on the 
    SFTP server.

.PARAMETER Option.RemoveFileAfterUpload
    Remove the source file when the upload to the SFTP servers was 
    successful.

.PARAMETER Option.ErrorWhenSourceFileIsNotFound
    Throw an error when there's no source file found or the source folder is 
    empty. Upon errors an e-mail is be sent to the admin.

.PARAMETER Sftp.Credential.UserName
    The user name used to authenticate to the SFTP server.

.PARAMETER Sftp.Credential.Password
    The password used to authenticate to the SFTP server.

.PARAMETER Sftp.ComputerName
    The URL where the SFTP server can be reached.

.PARAMETER Sftp.Path
    Path on the SFTP server where the uploaded files will be saved.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\Attentia\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Function Get-EnvironmentVariableValueHC {
            Param(
                [String]$Name
            )
        
            [Environment]::GetEnvironmentVariable($Name)
        }

        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        $Error.Clear()
        
        #region Create log folder
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
      
        $file = Get-Content $ImportFile -Raw -EA Stop -Encoding UTF8 | 
        ConvertFrom-Json
        #endregion
      
        #region Test .json file properties
        try {
            if (-not ($MailTo = $file.MailTo)) {
                throw "Property 'MailTo' not found"
            }
            if (-not ($Upload = $file.Upload)) {
                throw "Property 'Upload' not found"
            }

            foreach ($task in $Upload) {
                if (-not $task.Type) {
                    throw "Property 'Upload.Type' not found"
                }
                if ($task.Type -notMatch '^File$|^Folder$|^FolderContent$') {
                    throw "Property 'Upload.Type' must be 'File', 'Folder' or 'FolderContent'"
                }
                if (-not $task.Source) {
                    throw "Property 'Upload.Source' not found"
                }
                if (-not $task.Destination) {
                    throw "Property 'Upload.Destination' not found"
                }
                if (-not $task.Option) {
                    throw "Property 'Upload.Option' not found"
                }
                try {
                    [Boolean]::Parse($task.Option.OverwriteDestinationFile)
                }
                catch {
                    throw "Property 'Upload.Option.OverwriteDestinationFile' is not a boolean value"
                }
                try {
                    [Boolean]::Parse($task.Option.RemoveSourceAfterUpload)
                }
                catch {
                    throw "Property 'Upload.Option.RemoveSourceAfterUpload' is not a boolean value"
                }
                try {
                    [Boolean]::Parse($task.Option.ErrorWhen.SourceIsNotFound)
                }
                catch {
                    throw "Property 'Upload.Option.ErrorWhen.SourceIsNotFound' is not a boolean value"
                }
                if ($task.Type -match '^Folder$|^FolderContent$') {
                    try {
                        [Boolean]::Parse($task.Option.ErrorWhen.SourceFolderIsEmpty)
                    }
                    catch {
                        throw "Property 'Upload.Option.ErrorWhen.SourceFolderIsEmpty' is not a boolean value"
                    }   
                }
            }

            $Sftp = @{
                ComputerName = $file.Sftp.ComputerName
                Credential   = @{
                    UserName = Get-EnvironmentVariableValueHC -Name $file.Sftp.Credential.UserName
                    Password = Get-EnvironmentVariableValueHC -Name $file.Sftp.Credential.Password
                }
            }
            if (-not $sftp.ComputerName) {
                throw "Property 'Sftp.ComputerName' not found"
            }
            if (-not $sftp.Credential.UserName) {
                throw "Property 'Sftp.Credential.UserName' not found"
            }
            if (-not $sftp.Credential.Password) {
                throw "Property 'Sftp.Credential.Password' not found"
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion  

        #region Create SFTP credential
        try {
            $M = 'Create SFTP credential'
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $params = @{
                String      = $Sftp.Credential.Password 
                AsPlainText = $true
                Force       = $true
            }
            $secureStringPassword = ConvertTo-SecureString @params

            $params = @{
                TypeName     = 'System.Management.Automation.PSCredential'
                ArgumentList = $Sftp.Credential.UserName, $secureStringPassword
                ErrorAction  = 'Stop'
            }
            $sftpCredential = New-Object @params
        }
        catch {
            throw "Failed creating the SFTP credential with user name '$($Sftp.Credential.UserName)' and password '$($Sftp.Credential.Password)': $_"
        }
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        #region Open SFTP session
        try {
            $M = 'Start SFTP session'
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $params = @{
                ComputerName = $Sftp.ComputerName
                Credential   = $sftpCredential
                AcceptKey    = $true
                ErrorAction  = 'Stop'
            }
            $sftpSession = New-SFTPSession @params
        }
        catch {
            throw "Failed creating an SFTP session to '$($Sftp.ComputerName)': $_"
        }
        #endregion

        $results = @()

        $sessionParams = @{
            SessionId   = $sftpSession.SessionID
            ErrorAction = 'Stop'
        
        }
        foreach ($task in $Upload) {
            #region verbose
            $M = 'Upload task'
            $M += "Option OverwriteDestinationFile '{0}' RemoveSourceAfterUpload '{1}'" -f 
            $task.Option.OverwriteDestinationFile, 
            $task.Option.RemoveSourceAfterUpload
            $M += "Error when SourceIsNotFound '{0}' SourceFolderIsEmpty '{1}'" -f 
            $task.Option.ErrorWhen.SourceIsNotFound,
            $task.Option.ErrorWhen.SourceFolderIsEmpty
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            #endregion

            try {
                #region Test SFTP destination folder
                $M = "Test SFTP destination folder '{0}'" -f $task.Destination
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
                if (-not (
                        Test-SFTPPath @sessionParams -Path $task.Destination)
                ) {
                    throw "Upload destination folder '$($task.Destination)' not found on SFTP server"
                }    
                #endregion
            }
            catch {
                Write-Warning $_
                Continue
            }
            
            foreach ($source in $task.source) {
                try {
                    $result = [PSCustomObject]@{
                        Type        = $task.Type
                        Source      = $source
                        Destination = $task.Destination
                        UploadedOn  = $false
                        Info        = @()
                        Error       = $null
                    }

                    $M = "Type '{0}' Source '{1}' Destination '{2}'" -f 
                    $result.Type, $result.Source, $result.Destination
                    Write-Verbose $M
                    Write-EventLog @EventVerboseParams -Message $M
                
                    #region Test source exists
                    $testPathParams = @{
                        LiteralPath = $source
                        PathType    = 'Leaf'
                    }
                        
                    if ($task.Type -match '^Folder$|^FolderContent$') {
                        $testPathParams.PathType = 'Container'
                    }
                        
                    if (-not (Test-Path @testPathParams)) {
                        $M = 'Source path not found'

                        if ($task.Option.ErrorWhen.SourceIsNotFound) {
                            throw $M
                        }
                        $result.Info += $M
                        Continue
                    }
                    #endregion

                    #region Get source data
                    $sourceData = switch ($task.Type) {
                        'File' {
                            $result.Source
                            break
                        }
                        'Folder' {
                            $result.Source
                            break
                        }
                        'FolderContent' {
                            Get-ChildItem -LiteralPath $task.Source -File |
                            Select-Object -ExpandProperty 'FullName'
                            break
                        }
                        Default {
                            throw "Upload.Type '$_' not supported"
                        }
                    }
                    #endregion

                    #region Test if source folder is empty
                    $sourceFolderEmpty = $false

                    if (
                        ($task.Type -eq 'FolderContent') -and 
                        (-not $sourceData)
                    ) {
                        $sourceFolderEmpty = $true
                    }
                    if (
                        ($task.Type -eq 'Folder') -and
                        ((Get-ChildItem -LiteralPath $sourceData | 
                            Select-Object -First 1 | 
                            Measure-Object).Count -eq 0)
                    ) {
                        $sourceFolderEmpty = $true
                    }

                    if ($sourceFolderEmpty) {
                        $M = 'Source folder empty'
                        if ($task.Option.ErrorWhen.SourceFolderIsEmpty) {
                            throw $M
                        }
                        $result.Info += $M
                        Continue
                    }
                    #endregion

                    #region Upload data to SFTP server
                    $params = @{
                        Path        = $sourceData
                        Destination = $task.Destination
                    }

                    if ($task.Option.OverwriteDestinationFile) {
                        $params.Force = $true
                    }

                    Set-SFTPItem @sessionParams @params

                    $result.UploadedOn = Get-Date
                    #endregion

                    #region Remove source file or folder
                    if ($task.Option.RemoveSourceAfterUpload) {
                        $sourceData | Remove-Item -Force -EA Stop

                        switch ($task.Type) {
                            'File' { 
                                $result.Info += 'Removed source file'
                                break
                            }
                            'Folder' { 
                                $result.Info += 'Removed source folder'
                                break
                            }
                            'FolderContent' { 
                                $result.Info += 'Removed source folder content'
                                break
                            }
                            Default {
                                throw "Type '$_' not supported"
                            }
                        }
                    }
                    #endregion
                }
                catch {
                    $M = "Upload failed: $_"
                    Write-Warning $M
                    Write-EventLog @EventErrorParams -Message $M
                    $result.Error = $_
                    $Error.RemoveAt(0)
                }
                finally {
                    $results += $result
                }
            }                
        }
  
        #region Close SFTP session
        $M = 'Close SFTP session'
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
        Remove-SFTPSession -SessionId $sessionParams.SessionID -EA Ignore
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    try {
        $mailParams = @{}
       
        #region Create Excel worksheet Overview
        if ($results) {
            $excelParams = @{
                Path         = $logFile + ' - Log.xlsx'
                AutoSize     = $true
                FreezeTopRow = $true
            }

            $excelParams.WorksheetName = 'Overview'
            $excelParams.TableName = 'Overview'

            $M = "Export {0} rows to Excel sheet '{1}'" -f 
            $results.Count, $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
            $results | Select-Object *, @{
                Name       = 'Info'
                Expression = { $_.Info -join ', ' }
            } -ExcludeProperty 'Info' | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Send mail to user

        #region Counters
        $counter = @{
            Sources      = ($Upload.Source | Measure-Object).Count
            Uploaded     = $results.Where({ $_.UploadedOn }).Count
            UploadErrors = $results.Where({ $_.Error }).Count
            SystemErrors = (
                $Error.Exception.Message | Measure-Object
            ).Count
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'
        $mailParams.Subject = '{0}/{1} uploaded' -f 
        $counter.Uploaded, $counter.Sources

        if (
            $totalErrorCount = $counter.UploadErrors + $counter.SystemErrors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -ne 1) { 's' }
            )
        }
        #endregion
        
        #region Create html error list
        $systemErrorsHtmlList = if ($counter.SystemErrors) {
            "<p>Detected <b>{0} error{1}</b>:{2}</p>" -f $counter.SystemErrors, 
            $(
                if ($counter.SystemErrors -ne 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } | 
                ConvertTo-HtmlListHC
            )
        }
        #endregion

        #region Create html summary table
        $summaryHtmlTable = ''

        $i = 0

        foreach ($task in $Upload) {
            $i++

            $summaryHtmlTable += "
            <table>
                <tr>
                    <th colspan=`"2`">Task $i</th>
                </tr>
                <tr>
                    <td>Type</td>
                    <td>$($task.Type)</td>
                </tr>
                <tr>
                    <td>Source</td>
                    <td>$($task.Source -join '</br>')</td>
                </tr>
                <tr>
                    <td>Destination</td>
                    <td>$($task.Destination)</td>
                </tr>
                <tr>
                    <td>Options</td>
                    <td>Overwrite destination file: $($task.Option.OverwriteDestinationFile)</br>Remove source after upload: $($task.Option.RemoveSourceAfterUpload)</br>Error when source is not found: $($task.Option.ErrorWhen.SourceIsNotFound)</br>$(
                        if ($task.Type -ne 'File') {
                            "Error when source folder is empty: $($task.Option.ErrorWhen.SourceFolderIsEmpty)"
                        }
                    )</td>
                </tr>
            </table>
            " 
        }

        #endregion
                
        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                        $systemErrorsHtmlList
                        <p>Download files from an SFTP server.</p>
                        $summaryHtmlTable"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }
        
        if ($mailParams.Attachments) {
            $mailParams.Message += 
            "<p><i>* Check the attachment for details</i></p>"
        }
           
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion

    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}