#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog

<#
.SYNOPSIS
    Download files from an SFTP server or upload files to an SFTP server.

.DESCRIPTION
    Read an input file that contains all the required parameters for this 
    script. Then based on the action, a file is uploaded or downloaded.
    
    Based on the input file parameters a summary e-mail is sent to the user or 
    not. In any case, when there is an error, there's always an e-mail sent to 
    the admin.

    The computer that is running the SFTP code should have the module 'Posh-SSH'
    installed.

.PARAMETER ImportFile
    A .JSON file that contains all the parameters used by the script.

.PARAMETER Tasks
    Each task is a collection of upload or download actions. The exported Excel
    file is created based on the TaskName and is unique to a single task. 

    If different Excel log files or e-mails are required, use separate tasks 
    with a unique TaskName.

.PARAMETER Tasks.TaskName
    Name of the task. This name is used for naming the Excel log file and to 
    identify a task in the e-mail sent to the user.

.PARAMETER Tasks.Sftp.ComputerName
    The URL where the SFTP server can be reached.

.PARAMETER Tasks.Sftp.Credential.UserName
    The user name used to authenticate to the SFTP server. This is an 
    environment variable on the client running the script.
    
.PARAMETER Tasks.Sftp.Credential.Password
    The password used to authenticate to the SFTP server. This is an 
    environment variable on the client running the script.
    
.PARAMETER Tasks.Actions
    Each action represents a job that will either upload of download files
    to or from an SFTP server.

.PARAMETER Tasks.Actions.Type
    Indicate wether to upload or download files.

    Valid values:
    - Upload   : Upload files to the SFTP server
    - Download : Download files from the SFTP server

.PARAMETER Tasks.Actions.Parameter
    Parameters specific to the upload or download action.

.PARAMETER Tasks.Actions.Parameter.ComputerName
    The client where the SFTP code will be executed. This machine needs to 
    have the module 'Posh-SSH' installed.

.PARAMETER Tasks.Actions.Parameter.Path
    Type 'Upload':
    - One ore more full paths to a file or folder. When Path is a folder, 
      the files within that folder will be uploaded.

    Type 'Download':
    - A single folder on the machine defined in Tasks.Actions.Parameter.ComputerName where the files will be downloaded to.

.PARAMETER Tasks.Actions.Parameter.Option.OverwriteFile
    Overwrite a file on the SFTP server when it already exists.

.PARAMETER Tasks.Actions.Parameter.Option.RemoveFileAfterwards
    Remove a file after it was successfully uploaded to the SFTP server.

.PARAMETER Tasks.Actions.Parameter.Option.ErrorWhen.PathIsNotFound
    Throw an error when the file to upload is not found. When Path is a folder
    this option is ignored, because a folder can be empty.

.PARAMETER Tasks.SendMail.To
    E-mail addresses of users where to send the summary e-mail.

.PARAMETER Tasks.SendMail.When
    Indicate when an e-mail will be sent to the user.

    Valid values:
    - Always              : Always sent an e-mail
    - Never               : Never sent an e-mail
    - OnlyOnError         : Only sent an e-mail when errors where detected
    - OnlyOnErrorOrAction : Only sent an e-mail when errors where detected or
                            when items were uploaded

.PARAMETER Tasks.ExportExcelFile.When
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

                if (-not $task.Actions) {
                    throw 'Tasks.Actions is missing'
                }

                foreach ($action in $task.Actions) {
                    @('Type', 'Parameter').Where(
                        { -not $action.$_ }
                    ).foreach(
                        { throw "Property 'Tasks.Actions.$_' not found" }
                    )

                    switch ($action.Type) {
                        'Download' {
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
                
                        $M = "Start SFTP upload job '{0}' on '{1}' with arguments: Sftp.ComputerName '{2}' SftpPath '{3}' Sftp.UserName '{4}' Option.OverwriteFile '{5}' Option.RemoveFileAfterwards '{6}' Option.ErrorWhen.PathIsNotFound '{7}' Path '{8}'" -f 
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
                        $invokeParams = @{
                            FilePath     = $PathItem.DownloadScript
                            ArgumentList = $action.Parameter.Path, 
                            $task.Sftp.ComputerName, 
                            $action.Parameter.SftpPath, 
                            $task.Sftp.Credential.UserName, 
                            $task.Sftp.Credential.Password, 
                            $action.Parameter.Option.OverwriteFile, 
                            $action.Parameter.Option.RemoveFileAfterwards,
                            $action.Parameter.Option.ErrorWhen.PathIsNotFound
                        }
                
                        $M = "Start SFTP download job '{0}' on '{1}' with arguments: Sftp.ComputerName '{2}' SftpPath '{3}' Sftp.UserName '{4}' Option.OverwriteFile '{5}' Option.RemoveFileAfterwards '{6}' Option.ErrorWhen.PathIsNotFound '{7}' Path '{8}'" -f 
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
                    Actions         = 0
                }
                Action = @{
                    Errors          = 0
                    UploadedFiles   = 0
                    DownloadedFiles = 0
                }
            }
            #endregion

            $htmlTableActions = @()
            $exportToExcel = @()

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
                $counter.Total.Actions += $counter.Action.UploadedFiles
                $counter.Total.Actions += $counter.Action.DownloadedFiles
                #endregion

                #region Create HTML table
                $htmlTableActions += "
                <table>
                    <tr>
                        <th colspan=`"2`">
                            $(
                                if ($action.Type -eq 'Upload') {
                                    'UPLOAD FILES TO THE SFTP SERVER'
                                }
                            )
                            $(
                                if ($action.Type -eq 'Download') {
                                    'DOWNLOAD FILES FROM THE SFTP SERVER'
                                }
                            )
                        </th>
                    </tr>
                    <tr>
                        <td>SFTP path</td>
                        <td>$($action.Parameter.SftpPath)</td>
                    </tr>
                    <tr>
                        <td>Computer name</td>
                        <td>$($action.Parameter.ComputerName)</td>
                    </tr>
                    <tr>
                        <td>Path</td>
                        <td>$($action.Parameter.Path -join '<br>')</td>
                    </tr>
                    $(
                        if ($counter.Action.Errors) {
                            "<tr>
                                <td style=``"background-color: red``">Errors</td>
                                <td style=``"background-color: red``">$($counter.Action.Errors)</td>
                            </tr>"
                        }
                    )
                    $(
                        if ($action.Type -eq 'Upload') {
                            "<tr>
                                <td>Files uploaded</td>
                                <td>$($counter.Action.UploadedFiles)</td>
                            </tr>"
                        }
                    )
                    $(
                        if ($action.Type -eq 'Download') {
                            "<tr>
                                <td>Files downloaded</td>
                                <td>$($counter.Action.DownloadedFiles)</td>
                            </tr>"
                        }
                    )
                </table>
                "
                #endregion

                #region Create Excel objects

                $exportToExcel += $action.Job.Results | Select-Object DateTime, 
                @{
                    Name       = 'Type'
                    Expression = { $action.Type }
                }, @{
                    Name       = 'ComputerName'
                    Expression = { $_.PSComputerName }
                },
                @{
                    Name       = 'Source'
                    Expression = { 
                        if ($action.Type -eq 'Upload') {
                            $_.LocalPath -join ', ' 
                        }
                        else {
                            $_.SftpPath -join ', '
                        }
                    }
                }, 
                @{
                    Name       = 'Destination'
                    Expression = { 
                        if ($action.Type -eq 'Upload') {
                            $_.SftpPath -join ', '
                        }
                        else {
                            $_.LocalPath -join ', ' 
                        }
                    }
                }, 
                'FileName',
                @{
                    Name       = 'Action'
                    Expression = { $_.Action -join ', ' }
                }, 
                Error
                #endregion
            }

            $htmlTableActions = $htmlTableActions -join '<br>'
       
            #region Create Excel worksheet Overview
            $createExcelFile = $false

            if (
                (   
                    ($task.ExportExcelFile.When -eq 'OnlyOnError') -and 
                    ($counter.Total.Errors)
                ) -or
                (   
                    ($task.ExportExcelFile.When -eq 'OnlyOnErrorOrAction') -and 
                    (
                        ($counter.Total.Errors) -or ($counter.Total.Actions)
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
                    Path          = New-LogFileNameHC @excelFileLogParams
                    Append        = $true
                    AutoSize      = $true
                    FreezeTopRow  = $true
                    WorksheetName = 'Overview'
                    TableName     = 'Overview'
                    Verbose       = $false
                }

                $M = "Export {0} rows to Excel sheet '{1}'" -f 
                $exportToExcel, $excelParams.WorksheetName
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
                $exportToExcel | Export-Excel @excelParams

                $mailParams.Attachments = $excelParams.Path
            }
            #endregion

            #region Mail subject and priority
            $mailParams.Priority = 'Normal'
            $mailParams.Subject = @() 
            
            if ($task.Actions.Type -contains 'Upload') {
                $mailParams.Subject += "$($counter.Total.UploadedFiles) uploaded"
            }
            if ($task.Actions.Type -contains 'Download') {
                $mailParams.Subject += "$($counter.Total.DownloadedFiles) downloaded"
            }
            
            if ($counter.Total.Errors) {
                $mailParams.Priority = 'High'
                $mailParams.Subject += "{0} error{1}" -f 
                $counter.Total.Errors,
                $(if ($counter.Total.Errors -ne 1) { 's' })
            }

            $mailParams.Subject = $mailParams.Subject -join ', '
            #endregion

            #region Check to send mail to user
            $sendMailToUser = $false

            if (
                (
                    ($task.SendMail.When -eq 'Always')
                ) -or
                (   
                    ($task.SendMail.When -eq 'OnlyOnError') -and 
                    ($counter.Total.Errors)
                ) -or
                (   
                    ($task.SendMail.When -eq 'OnlyOnErrorOrAction') -and 
                    (
                        ($counter.Total.Actions) -or ($counter.Total.Errors)
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
                $(
                    if ($task.Actions.Type -contains 'Upload') {
                        "<tr>
                            <td>Total files uploaded</td>
                            <td>$($counter.Total.UploadedFiles)</td>
                        </tr>"
                    }
                )
                $(
                    if ($task.Actions.Type -contains 'Download') {
                        "<tr>
                            <td>Total files downloaded</td>
                            <td>$($counter.Total.DownloadedFiles)</td>
                        </tr>"
                    }
                )
                $(
                    if ($counter.Total.Errors) {
                        "<tr>
                            <td style=``"background-color: red``">Total errors</td>
                            <td style=``"background-color: red``">$($counter.Total.Errors)</td>
                        </tr>"
                    }
                )
            </table>"
            #endregion
                
            #region Send mail
            $mailParams += @{
                To        = $task.SendMail.To
                Message   = "
                        $systemErrorsHtmlList
                        <p>Summary of all SFTP actions.</p>
                        $summaryHtmlTable
                        <p>Action details.</p>
                        $htmlTableActions"
                
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

                if ($counter.Total.Errors) {
                    $mailParams.Bcc = $ScriptAdmin
                }
                Send-MailHC @mailParams
            }
            else {
                Write-Verbose 'Send no e-mail to the user'

                if ($counter.Total.Errors) {
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