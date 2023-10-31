#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog

<#
.SYNOPSIS
    Upload files to an SFTP server.

.DESCRIPTION
    Upload all files in a folder or a single file to an SFTP server. Send an 
    e-mail to the user when needed, but always send an e-mail to the admin on 
    errors.

    The computer that is running the SFTP code should have the module 'Posh-SSH'
    installed.

.PARAMETER ImportFile
    A .JSON file that contains all the parameters used by the script.

.PARAMETER Tasks
    Each task in Tasks represents a job that needs to be executed. Multiple
    SFTP jobs are supported.

.PARAMETER TaskName
    Name of the task. This name is used for naming the Excel log file and for
    naming the tasks in the e-mail sent to the user.

.PARAMETER Task.ExecuteOnComputerName
    Defines on which machine the SFTP commands are executed. Can be the 
    localhost or a remote computer name.

.PARAMETER Sftp.ComputerName
    The URL where the SFTP server can be reached.
    
.PARAMETER Sftp.Path
    Path on the SFTP server where the uploaded files will be saved.

.PARAMETER Sftp.Credential.UserName
    The user name used to authenticate to the SFTP server. This is an 
    environment variable on the client running the script.
    
.PARAMETER Sftp.Credential.Password
    The password used to authenticate to the SFTP server. This is an 
    environment variable on the client running the script.
    
.PARAMETER Upload.Path
    One ore more full paths to a file or folder. When Path is a folder, the files within that folder will be uploaded.

.PARAMETER Upload.Option.OverwriteFile
    Overwrite a file on the SFTP server when it already exists.

.PARAMETER Upload.Option.RemoveFileAfterwards
    Remove a file after it was successfully uploaded to the SFTP server.

.PARAMETER Upload.Option.ErrorWhen.PathIsNotFound
    Throw an error when the file to upload is not found. When Path is a folder
    this option is ignored, because a folder can be empty.

.PARAMETER SendMail.To
    E-mail addresses of users where to send the summary e-mail.

.PARAMETER SendMail.When
    Indicate when an e-mail will be sent to the user.

    Valid values:
    - Always              : Always sent an e-mail
    - Never               : Never sent an e-mail
    - OnlyOnError         : Only sent an e-mail when errors where detected
    - OnlyOnErrorOrAction : Only sent an e-mail when errors where detected or
                            when items were uploaded

.PARAMETER ExportExcelFile.When
    Indicate when an Excel file will be created containing the log data.

    Valid values:
    - Never               : Never create an Excel log file
    - OnlyOnError         : Only create an Excel log file when 
                            errors where detected
    - OnlyOnErrorOrAction : Only create an Excel log file when 
                            errors where detected or when items were uploaded
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [HashTable]$Path = @{
        UploadScript   = "$PSScriptRoot\Send to SFTP.ps1"
        DownloadScript = "$PSScriptRoot\Get SFTP file.ps1"
    },
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

        #region Test path exists
        $pathItem = @{}

        $Path.GetEnumerator().ForEach(
            {
                try {
                    $key = $_.Key
                    $value = $_.Value

                    $params = @{
                        Path        = $value
                        ErrorAction = 'Stop'
                    }
                    $PathItem[$key] = (Get-Item @params).FullName
                }
                catch {
                    throw "Path.$key '$value' not found"
                }
            }
        )
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
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
      
        $file = Get-Content $ImportFile -Raw -EA Stop -Encoding UTF8 | 
        ConvertFrom-Json
        #endregion
      
        #region Test .json file properties
        try {
            if (-not ($Tasks = $file.Tasks)) {
                throw "Property 'Tasks' not found"
            }

            if (-not $file.MaxConcurrentJobs) {
                throw "Property 'MaxConcurrentJobs' not found"
            }

            #region Test integer value
            try {
                [int]$MaxConcurrentJobs = $file.MaxConcurrentJobs
            }
            catch {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$($file.MaxConcurrentJobs)' is not supported."
            }
            #endregion

            foreach ($task in $Tasks) {
                @(
                    'TaskName', 'Sftp', 'Actions', 'SendMail', 'ExportExcelFile'
                ).where(
                    { -not $task.$_ }
                ).foreach(
                    { throw "Property 'Tasks.$_' not found" }
                )
                
                if (-not $task.TaskName) {
                    throw "Property 'Tasks.TaskName' not found" 
                }

                @('ComputerName', 'Credential').where(
                    { -not $task.Sftp.$_ }
                ).foreach(
                    { throw "Property 'Tasks.Sftp.$_' not found" }
                )
                                
                @('UserName', 'Password').Where(
                    { -not $task.Sftp.Credential.$_ }
                ).foreach(
                    { throw "Property 'Tasks.Sftp.Credential.$_' not found" }
                )

                foreach ($action in $task.Actions) {
                    @('Type', 'Parameter').Where(
                        { -not $action.$_ }
                    ).foreach(
                        { throw "Property 'Tasks.Actions.$_' not found" }
                    )

                    switch ($action.Type) {
                        'Download' {

                            break
                        }
                        'Upload' {
                            @(
                                'SftpPath', 'ComputerName', 
                                'Path', 'Option'
                            ).Where(
                                { -not $action.Parameter.$_ }
                            ).foreach(
                                { throw "Property 'Tasks.Actions.Parameter.$_' not found" }
                            )

                            #region Test boolean values
                            foreach (
                                $boolean in 
                                @(
                                    'OverwriteFile', 
                                    'RemoveFileAfterwards'
                                )
                            ) {
                                try {
                                    $null = [Boolean]::Parse($action.Parameter.Option.$boolean)
                                }
                                catch {
                                    throw "Property 'Tasks.Actions.Parameter.Option.$boolean' is not a boolean value"
                                }
                            }

                            foreach (
                                $boolean in 
                                @(
                                    'PathIsNotFound',
                                    'SftpPathIsNotFound'
                                )
                            ) {
                                try {
                                    $null = [Boolean]::Parse($action.Parameter.Option.ErrorWhen.$boolean)
                                }
                                catch {
                                    throw "Property 'Tasks.Actions.Parameter.Option.ErrorWhen.$boolean' is not a boolean value"
                                }
                            }
                            #endregion
                
                            break
                        }
                        Default {
                            throw "Tasks.Actions.Type '$_' not supported. Only the values 'Upload' or 'Download' are supported."
                        }
                    }
                }
                
                @('To', 'When').Where(
                    { -not $task.SendMail.$_ }
                ).foreach(
                    { throw "Property 'Tasks.SendMail.$_' not found" }
                )
                
                @('When').Where(
                    { -not $task.ExportExcelFile.$_ }
                ).foreach(
                    { throw "Property 'Tasks.ExportExcelFile.$_' not found" }
                )

                #region Test When is valid
                if ($file.SendMail.When -ne 'Never') {   
                    if ($task.SendMail.When -notMatch '^Always$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                        throw "Property 'Tasks.SendMail.When' with value '$($task.SendMail.When)' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
                    }

                    if ($task.ExportExcelFile.When -notMatch '^Always$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                        throw "Property 'Tasks.ExportExcelFile.When' with value '$($task.ExportExcelFile.When)' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
                    }
                }
                #endregion
            }

            #region Test unique TaskName
            $Tasks.TaskName | Group-Object | Where-Object {
                $_.Count -gt 1
            } | ForEach-Object {
                throw "Property 'Tasks.TaskName' with value '$($_.Name)' is not unique. Each task name needs to be unique."
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
        foreach ($task in $Tasks) {
            foreach ($action in $task.Actions) {
                #region Create job parameters
                switch ($action.Type) {
                    'Upload' {  
                        $invokeParams = @{
                            FilePath     = $PathItem.UploadScript
                            ArgumentList = $action.Parameter.Path, 
                            $task.Sftp.ComputerName, 
                            $action.Parameter.SftpPath, 
                            $task.Sftp.Credential.UserName, 
                            $task.Sftp.Credential.Password, 
                            $action.Parameter.Option.OverwriteFile, 
                            $action.Parameter.Option.RemoveFileAfterwards,
                            $action.Parameter.Option.ErrorWhen.PathIsNotFound
                        }
                
                        $M = "Start job '{0}' on '{1}' with arguments: Sftp.ComputerName '{2}' SftpPath '{3}' Sftp.UserName '{4}' Option.OverwriteFile '{5}' Option.RemoveFileAfterwards '{6}' Option.ErrorWhen.PathIsNotFound '{7}' Path '{8}'" -f 
                        $task.TaskName, 
                        $action.Parameter.ComputerName,
                        $invokeParams.ArgumentList[1], 
                        $invokeParams.ArgumentList[2], 
                        $invokeParams.ArgumentList[3], 
                        $invokeParams.ArgumentList[5],
                        $invokeParams.ArgumentList[6], 
                        $invokeParams.ArgumentList[7],
                        $($invokeParams.ArgumentList[0] -join "', '")
                        Write-Verbose $M; 
                        Write-EventLog @EventVerboseParams -Message $M

                        break
                    }
                    'Download' {  

                        break
                    }
                    Default {
                        throw "Tasks.Actions.Type '$_' not supported."
                    }
                }
                #endregion

                $action | Add-Member -NotePropertyMembers @{
                    Job = @{
                        Object  = $null
                        Results = @()
                    }
                }
          
                #region Start job
                $computerName = $action.Parameter.ComputerName 

                $action.Job.Object = if (
                    ($computerName) -and
                    ($computerName -ne 'localhost') -and
                    ($computerName -ne $ENV:COMPUTERNAME) -and
                    ($computerName -ne "$ENV:COMPUTERNAME.$env:USERDNSDOMAIN")
                ) {
                    $invokeParams.ComputerName = $computerName
                    $invokeParams.AsJob = $true
                    Invoke-Command @invokeParams
                }
                else {
                    $action.Parameter.ComputerName = $ENV:COMPUTERNAME
                    Start-Job @invokeParams
                }
                #endregion
        
                #region Wait for max running jobs
                $params = @{
                    Name       = $Tasks.Actions.Job.Object | Where-Object { $_ }
                    MaxThreads = $MaxConcurrentJobs     
                }
                Wait-MaxRunningJobsHC @params
                #endregion
            }
        }

        #region Wait for all jobs to finish
        Write-Verbose 'Wait for all jobs to finish'
        
        $null = $Tasks.Actions.Job.Object | Wait-Job
        #endregion

        #region Get job results
        foreach ($task in $Tasks) {
            foreach ($action in $task.Actions) {
                $action.Job.Results += Receive-Job -Job $action.Job.Object
                
                $M = "Task '{0}' type '{1}' {2} job result{3}" -f 
                $task.TaskName,
                $action.Type,
                $action.Job.Results.Count,
                $(if ($action.Job.Results.Count -ne 1) { 's' })
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            }
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
        $countSystemErrors = (
            $Error.Exception.Message | Measure-Object
        ).Count
        
        #region Create error html lists
        $systemErrorsHtmlList = if ($countSystemErrors) {
            "<p>Detected <b>{0} system error{1}</b>:{2}</p>" -f $countSystemErrors, 
            $(
                if ($countSystemErrors -ne 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } | 
                ConvertTo-HtmlListHC
            )
        }
        #endregion

        foreach ($task in $Tasks) {
            $mailParams = @{}

            #region Counter
            $counter = @{
                Total  = @{
                    Errors          = $countSystemErrors
                    UploadedFiles   = 0
                    DownloadedFiles = 0
                }
                Action = @{
                    Errors          = 0
                    UploadedFiles   = 0
                    DownloadedFiles = 0
                }
            }
            #endregion

            foreach ($action in $task.Actions) {
                #region Update counters
                $counter.Action.UploadedFiles = $action.Job.Results.Where(
                    { $_.Uploaded }).Count
                $counter.Action.DownloadedFiles = $action.Job.Results.Where(
                    { $_.Downloaded }).Count
                $counter.Action.Errors = $action.Job.Results.Where(
                    { $_.Error }).Count

                $counter.Total.Errors += $counter.Action.Errors
                $counter.Total.UploadedFiles += $counter.Action.UploadedFiles
                $counter.Total.DownloadedFiles += $counter.Action.DownloadedFiles
                #endregion

                #region Create HTML table
                $htmlTableTasks += "
            <table>
            <tr>
                <th colspan=`"2`">$($action.Type)</th>
            </tr>
            <tr>
                <td>Upload path</td>
                <td>$($task.Upload.Path -join '<br>')</td>
            </tr>
            <tr>
                <td>Options</td>
                <td>
                    Overwrite file on SFTP server: $($task.Upload.Option.OverwriteFile)<br>
                    Remove file after upload: $($task.Upload.Option.RemoveFileAfterwards)<br>
                    Error when upload path is not found: $($task.Upload.Option.ErrorWhen.PathIsNotFound)<br>
                </td>
            </tr>

            <tr>
                <td>Details</td>
                <td>
                    <a href=`"$($inputFile.ExcelFile.OutputFolder)`">Output folder</a>
                </td>
            </tr>
            <tr>
                <td>$($counter.RowsInExcel)</td>
                <td>Files to download</td>
            </tr>
            <tr>
                <td>$($counter.DownloadedFiles)</td>
                <td>Files successfully downloaded</td>
            </tr>
            $(
                if ($counter.Errors.InExcelFile) {
                    "<tr>
                        <td style=``"background-color: red``">$($counter.Errors.InExcelFile)</td>
                        <td style=``"background-color: red``">Error{0} in the Excel file</td>
                    </tr>" -f $(if ($counter.Errors.InExcelFile -ne 1) {'s'})
                }
            )
            $(
                if ($counter.Errors.DownloadingFiles) {
                    "<tr>
                        <td style=``"background-color: red``">$($counter.Errors.DownloadingFiles)</td>
                        <td style=``"background-color: red``">File{0} failed to download</td>
                    </tr>" -f $(if ($counter.Errors.DownloadingFiles -ne 1) {'s'})
                }
            )
            $(
                if ($counter.Errors.Other) {
                    "<tr>
                        <td style=``"background-color: red``">$($counter.Errors.Other)</td>
                        <td style=``"background-color: red``">Error{0} found:<br>{1}</td>
                    </tr>" -f $(
                        if ($counter.Errors.Other -ne 1) {'s'}
                    ),
                    (
                        '- ' + $($inputFile.Error -join '<br> - ')
                    )
                }
            )
            $(
                if($inputFile.Tasks) {
                    "<tr>
                        <th colspan=``"2``">Downloads per folder</th>
                    </tr>"
                }
            )
            $(
                foreach (
                    $task in 
                    (
                        $inputFile.Tasks | 
                        Sort-Object {$_.DownloadFolder.Name}
                    )
                ) {
                    $errorCount = $task.Job.Result.Where(
                        {$_.Error}).Count

                    $template = if ($errorCount) {
                        "<tr>
                        <td style=``"background-color: red``">{0}/{1}</td>
                        <td style=``"background-color: red``">{2}{3}</td>
                        </tr>"     
                    } else {
                        "<tr>
                            <td>{0}/{1}</td>
                            <td>{2}{3}</td>
                        </tr>" 
                    }

                    $template -f 
                    $(
                        $task.Job.Result.Where({$_.DownloadedOn}).Count
                    ),
                    $(
                        ($task.ItemsToDownload | Measure-Object).Count
                    ),
                    $(
                        $task.DownloadFolder.Name
                    ),
                    $(
                        if ($errorCount) {
                            ' ({0} error{1})' -f 
                            $errorCount, $(if ($errorCount -ne 1) {'s'})
                        }
                    )
                }
            )
        </table>
        "
                #endregion
            }

            $htmlTableTasks = $htmlTableTasks -join '<br>'
       
            #region Create Excel worksheet Overview
            $createExcelFile = $false

            if (
                (   
                    ($task.ExportExcelFile.When -eq 'OnlyOnError') -and 
                    ($counter.UploadErrors -ne 0)
                ) -or
                (   
                    ($task.ExportExcelFile.When -eq 'OnlyOnErrorOrAction') -and 
                    (
                        ($counter.UploadErrors -ne 0) -or 
                        ($counter.Uploaded -ne 0)
                    )
                )
            ) {
                $createExcelFile = $true
            }


            if ($createExcelFile) {
                $excelFileLogParams = @{
                    LogFolder    = $logParams.LogFolder
                    Format       = 'yyyy-MM-dd'
                    Name         = "$ScriptName - $($task.TaskName) - Log.xlsx"
                    Date         = 'ScriptStartTime'
                    NoFormatting = $true
                }

                $excelParams = @{
                    Path         = New-LogFileNameHC @excelFileLogParams
                    Append       = $true
                    AutoSize     = $true
                    FreezeTopRow = $true
                    Verbose      = $false
                }

                $excelParams.WorksheetName = 'Overview'
                $excelParams.TableName = 'Overview'

                $M = "Export {0} rows to Excel sheet '{1}'" -f 
                $task.Job.Results.Count, $excelParams.WorksheetName
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
                $task.Job.Results | Select-Object @{
                    Name       = 'ComputerName'
                    Expression = { $_.PSComputerName }
                },
                DateTime, 
                @{
                    Name       = 'Path'
                    Expression = { $_.Path -join ', ' }
                }, 
                @{
                    Name       = 'Action'
                    Expression = { $_.Action -join ', ' }
                }, Error | Export-Excel @excelParams

                $mailParams.Attachments = $excelParams.Path
            }
            #endregion

            #region Mail subject and priority
            $mailParams.Priority = 'Normal'
            $mailParams.Subject = '{0} item{1} uploaded' -f 
            $counter.Uploaded, $(if ($counter.Uploaded -ne 1) { 's' })

            if ($counter.TotalErrors) {
                $mailParams.Priority = 'High'
                $mailParams.Subject += ",{0} error{1}" -f 
                $counter.TotalErrors,
                $(if ($counter.TotalErrors -ne 1) { 's' })
            }
            #endregion

            #region Check to send mail to user
            $sendMailToUser = $false

            if (
                (
                    ($task.SendMail.When -eq 'Always')
                ) -or
                (   
                    ($task.SendMail.When -eq 'OnlyOnError') -and 
                    ($counter.TotalErrors)
                ) -or
                (   
                    ($task.SendMail.When -eq 'OnlyOnErrorOrAction') -and 
                    (
                        ($counter.Uploaded) -or ($counter.TotalErrors)
                    )
                )
            ) {
                $sendMailToUser = $true
            }
            #endregion

            #region Create html summary table
            $summaryHtmlTable = "
            <table>
                <tr>
                    <th colspan=`"2`">$($task.TaskName)</th>
                </tr>
                <tr>
                    <td>SFTP Server</td>
                    <td>$($task.Sftp.ComputerName)</td>
                </tr>
                <tr>
                    <td>SFTP User name</td>
                    <td>$($task.Sftp.Credential.UserName)</td>
                </tr>
            </table>"
            #endregion
                
            #region Send mail
            $mailParams += @{
                To        = $task.SendMail.To
                Message   = "
                        $systemErrorsHtmlList
                        <p>Uploaded <b>{0} file{1}</b> to the SFTP server below{2}.</p>
                        $summaryHtmlTable" -f 
                $counter.Uploaded, 
                $(
                    if ($counter.Uploaded -ne 1) { 's' }
                ),
                $(
                    if ($counter.TotalErrors) {
                        ' and found <b>{0} error{1}</b>' -f 
                        $counter.TotalErrors,
                        $(if ($counter.TotalErrors -ne 1) { 's' })
                    }
                )
                
                LogFolder = $LogParams.LogFolder
                Header    = $ScriptName
                Save      = $LogFile + ' - Mail.html'
            }
        
            if ($mailParams.Attachments) {
                $mailParams.Message += 
                "<p><i>* Check the attachment for details</i></p>"
            }
        
            Get-ScriptRuntimeHC -Stop

            if ($sendMailToUser) {
                Write-Verbose 'Send e-mail to the user'

                if ($counter.TotalErrors) {
                    $mailParams.Bcc = $ScriptAdmin
                }
                Send-MailHC @mailParams
            }
            else {
                Write-Verbose 'Send no e-mail to the user'

                if ($counter.TotalErrors) {
                    Write-Verbose 'Send e-mail to admin only with errors'
                    
                    $mailParams.To = $ScriptAdmin
                    Send-MailHC @mailParams
                }
            }
            #endregion
        }
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