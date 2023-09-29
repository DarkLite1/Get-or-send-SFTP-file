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

.PARAMETER SendMail.To
    E-mail addresses of where to send the summary e-mail.

.PARAMETER SendMail.When
    Indicate when an e-mail will be sent.

    Valid values:
    - Always              : Always sent an e-mail
    - Never               : Never sent an e-mail
    - OnlyOnError         : Only sent an e-mail when errors where detected
    - OnlyOnErrorOrUpload : Only sent an e-mail when errors where detected or
                            when items were uploaded

.PARAMETER ExportExcelFile.When
    Indicate when an Excel file will be created containing the log data.

    Valid values:
    - Always              : Always create an Excel log file
    - Never               : Never create an Excel log file
    - OnlyOnError         : Only create an Excel log file when 
                            errors where detected
    - OnlyOnErrorOrUpload : Only create an Excel log file when 
                            errors where detected or when items were uploaded

.PARAMETER Upload.Type
    Defines which files need to be uploaded to the SFTP server.
    Valid values:
    - File         : a single file
    - Folder       : the complete content of a folder
    - FolderContent: all files in a folder, not recursive

.PARAMETER Upload.Path
    The path to the file or folder.

.PARAMETER Option.OverwriteDestinationData
    Overwrite a data on the SFTP server when it already exists.

.PARAMETER Option.RemoveSourceAfterUpload
    Remove the source data when the upload was successful.

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
    [string]$SftpScriptPath = "$PSScriptRoot\Send to SFTP.ps1",
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Send file to SFTP server\$ScriptName",
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

        $null = Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        $Error.Clear()
        
        #region Test SFTP script path exits
        try {
            $params = @{
                Path        = $SftpScriptPath
                ErrorAction = 'Stop'
            }
            $SftpScriptPathItem = (Get-Item @params).FullName
        }
        catch {
            throw "SftpScriptPath '$SftpScriptPath' not found"
        }
        #endregion

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
            if (-not ($Tasks = $file.Tasks)) {
                throw "Property 'Tasks' not found"
            }

            if (-not ($MaxConcurrentJobs = $file.MaxConcurrentJobs)) {
                throw "Property 'MaxConcurrentJobs' not found"
            }

            #region Test integer value
            if (-not ($MaxConcurrentJobs -is [int])) {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$MaxConcurrentJobs' is not supported."
            }
            #endregion

            foreach ($task in $Tasks) {
                @('Task', 'Sftp', 'Upload', 'SendMail', 'ExportExcelFile') | 
                Where-Object { -not $task.$_ } |
                ForEach-Object {
                    throw "Property 'Tasks.$_' not found"
                }
                
                @('Name', 'ExecuteOnComputerName') | 
                Where-Object { -not $task.Task.$_ } |
                ForEach-Object {
                    throw "Property 'Tasks.Task.$_' not found"
                }

                @('ComputerName', 'Path', 'Credential') | 
                Where-Object { -not $task.Sftp.$_ } |
                ForEach-Object {
                    throw "Property 'Tasks.Sftp.$_' not found"
                }
                
                @('UserName', 'Password') | 
                Where-Object { -not $task.Sftp.Credential.$_ } |
                ForEach-Object {
                    throw "Property 'Tasks.Sftp.Credential.$_' not found"
                }
                
                @('Path', 'Option') | 
                Where-Object { -not $task.Upload.$_ } |
                ForEach-Object {
                    throw "Property 'Tasks.Upload.$_' not found"
                }
                
                @('To', 'When') | 
                Where-Object { -not $task.SendMail.$_ } |
                ForEach-Object {
                    throw "Property 'Tasks.SendMail.$_' not found"
                }
                
                @('When') | 
                Where-Object { -not $task.ExportExcelFile.$_ } |
                ForEach-Object {
                    throw "Property 'Tasks.ExportExcelFile.$_' not found"
                }

                #region Test boolean values
                foreach (
                    $boolean in 
                    @(
                        'OverwriteFileOnSftpServer', 
                        'RemoveFileAfterUpload'
                    )
                ) {
                    try {
                        $null = [Boolean]::Parse($task.Upload.Option.$boolean)
                    }
                    catch {
                        throw "Property 'Tasks.Upload.Option.$boolean' is not a boolean value"
                    }
                }

                foreach (
                    $boolean in 
                    @(
                        'UploadPathIsNotFound'
                    )
                ) {
                    try {
                        $null = [Boolean]::Parse($task.Upload.Option.ErrorWhen.$boolean)
                    }
                    catch {
                        throw "Property 'Tasks.Upload.Option.ErrorWhen.$boolean' is not a boolean value"
                    }
                }
                #endregion

                #region Test When is valid
                if ($task.SendMail.When -notMatch '^Always$|^Never$|^OnlyOnError$|^OnlyOnErrorOrUpload$') {
                    throw "Property 'Tasks.SendMail.When' with value '$($task.SendMail.When)' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrUpload'"
                }

                if ($task.ExportExcelFile.When -notMatch '^Always$|^Never$|^OnlyOnError$|^OnlyOnErrorOrUpload$') {
                    throw "Property 'Tasks.ExportExcelFile.When' with value '$($task.ExportExcelFile.When)' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrUpload'"
                }
                #endregion
            }

            #region Test unique Task.Name
            $Tasks.Task.Name | Group-Object | Where-Object {
                $_.Count -gt 1
            } | ForEach-Object {
                throw "Property 'Tasks.Task.Name' with value '$($_.Name)' is not unique. Each task name needs to be unique."
            }
            #endregion
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion

        try {
            foreach ($task in $Tasks) {
                #region Add environment variable SFTP Password
                $params = @{
                    String      = $null
                    AsPlainText = $true
                    Force       = $true
                    ErrorAction = 'Stop'
                }

                if (-not (
                        $params.String = Get-EnvironmentVariableValueHC -Name $task.Sftp.Credential.Password)
                ) {
                    throw "Environment variable '`$ENV:$($task.Sftp.Credential.Password)' in 'Sftp.Credential.Password' not found on computer $ENV:COMPUTERNAME"
                }
                
                $task.Sftp.Credential.Password = ConvertTo-SecureString @params
                #endregion

                #region Add environment variable SFTP UserName
                if (-not (
                        $params.String = Get-EnvironmentVariableValueHC -Name $task.Sftp.Credential.UserName)
                ) {
                    throw "Environment variable '`$ENV:$($task.Sftp.Credential.UserName)' in 'Sftp.Credential.UserName' not found on computer $ENV:COMPUTERNAME"
                }

                $task.Sftp.Credential.UserName = $params.String
                #endregion
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
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
        #region Start jobs to upload files        
        foreach ($task in $Tasks) {
            $task | Add-Member -NotePropertyMembers @{
                Job = @{
                    Object  = $null
                    Results = @()
                }
            }

            $invokeParams = @{
                FilePath     = $SftpScriptPathItem
                ArgumentList = $task.Upload.Path, $task.Sftp.ComputerName, 
                $task.Sftp.Path, $task.Sftp.UserName, $task.Sftp.Password, 
                $task.Option.OverwriteFileOnSftpServer, 
                $task.Option.RemoveFileAfterUpload,
                $task.Option.ErrorWhenUploadPathIsNotFound
            }
    
            $M = "Start job '{0}' on '{1}' with arguments: Sftp.ComputerName '{2}' Sftp.Path '{3}' Sftp.UserName '{4}' Option.OverwriteFileOnSftpServer '{5}' Option.RemoveFileAfterUpload '{6}' Option.ErrorWhenUploadPathIsNotFound '{7}' Upload.Path '{8}'" -f $task.Task.Name, $task.Task.ExecuteOnComputerName,
            $invokeParams.ArgumentList[1], $invokeParams.ArgumentList[2], 
            $invokeParams.ArgumentList[3], $invokeParams.ArgumentList[5],
            $invokeParams.ArgumentList[6], $invokeParams.ArgumentList[7],
            $($invokeParams.ArgumentList[0] -join "', '")
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
      
            $task.Job.Object = if (
                ($task.Task.ExecuteOnComputerName) -and
                ($task.Task.ExecuteOnComputerName -ne 'localhost') -and
                ($task.Task.ExecuteOnComputerName -ne $ENV:COMPUTERNAME) -and
                ($task.Task.ExecuteOnComputerName -ne "$ENV:COMPUTERNAME.$env:USERDNSDOMAIN")
            ) {
                $invokeParams.ComputerName = $task.Task.ExecuteOnComputerName
                $invokeParams.AsJob = $true
                Invoke-Command @invokeParams
            }
            else {
                Start-Job @invokeParams
            }
    
            $params = @{
                Name       = $Tasks.Task.Job.Object | Where-Object { $_ }
                MaxThreads = $MaxConcurrentJobs     
            }
            Wait-MaxRunningJobsHC @params
        }
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
        #region Counters
        $counter = @{
            Sources      = ($Upload.Source | Measure-Object).Count
            Uploaded     = ($results.UploadedItems | Measure-Object -Sum).Sum
            UploadErrors = $results.Where({ $_.Error }).Count
            SystemErrors = (
                $Error.Exception.Message | Measure-Object
            ).Count
        }
        #endregion

        $mailParams = @{}
       
        #region Create Excel worksheet Overview
        $createExcelFile = $false

        if (
            (
                $ExportExcelFile.When -eq 'Always'
            ) -or
            (   
                ($ExportExcelFile.When -eq 'OnlyOnError') -and 
                ($counter.UploadErrors -ne 0)
            ) -or
            (   
                ($ExportExcelFile.When -eq 'OnlyOnErrorOrUpload') -and 
                ($counter.UploadErrors -ne 0) -or ($counter.Uploaded -ne 0)
            )
        ) {
            $createExcelFile = $true
        }


        if ($createExcelFile -and $results) {
            $excelFileLogParams = @{
                LogFolder    = $logParams.LogFolder
                Format       = 'yyyy-MM-dd'
                Name         = "$ScriptName - Log.xlsx"
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }

            $excelParams = @{
                Path         = New-LogFileNameHC @excelFileLogParams
                Append       = $true
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

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'
        $mailParams.Subject = '{0} item{1} uploaded' -f 
        $counter.Uploaded, $(
            if ($counter.Uploaded -ne 1) {
                's'
            }
        )

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
                    <td>Overwrite destination file: $($task.Option.OverwriteDestinationData)</br>Remove source after upload: $($task.Option.RemoveSourceAfterUpload)</br>Error when source is not found: $($task.Option.ErrorWhen.SourceIsNotFound)</br>$(
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
            To        = $SendMail.To
            Bcc       = $ScriptAdmin
            Message   = "
                        $systemErrorsHtmlList
                        <p>Upload files to an SFTP server.</p>
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

        if (
            (
                $SendMail.When -eq 'Always'
            ) -or
            (   
                ($SendMail.When -eq 'OnlyOnError') -and 
                ($totalErrorCount -ne 0)
            ) -or
            (   
                ($SendMail.When -eq 'OnlyOnErrorOrUpload') -and 
                (($totalErrorCount -ne 0) -or ($counter.Uploaded -ne 0))
            )
        ) {
            Send-MailHC @mailParams
        }
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