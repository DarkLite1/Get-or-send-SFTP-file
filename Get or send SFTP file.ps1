#Requires -Version 7
#Requires -Modules ImportExcel
#Requires -Modules Toolbox.EventLog, Toolbox.HTML

<#
.SYNOPSIS
    Download files from an SFTP server or upload files to an SFTP server.

.DESCRIPTION
    Read an input file that contains all the parameters for this script. The
    Action.Type defines if it's an upload or a download. When SendMail.When is
    used, a summary e-mail is sent.

    The computer that is running the SFTP code should have the module 'Posh-SSH'
    installed.

    Actions run in parallel when MaxConcurrentJobs is more than 1. Tasks
    run sequentially.

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

.PARAMETER Tasks.Sftp.Credential.PasswordKeyFile
    The password used to authenticate to the SFTP server. This is an
    SSH private key file in the OpenSSH format.

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

.PARAMETER Tasks.Actions.Parameter.PartialFileExtension
    When type is 'Upload' the file that needs to be uploaded is first renamed
    by adding another file extension. This will make sure that errors like
    "file in use by another process" are avoided.

    After a rename the file is uploaded with the extension defined in
    "PartialFileExtension". After a successful upload the file is then renamed
    on the SFTP server to its original name with the correct file extension.

.PARAMETER Tasks.Actions.Parameter.Option.OverwriteFile
    Overwrite a file on the SFTP server when it already exists.

.PARAMETER Tasks.Actions.Parameter.Option.RemoveFileAfterwards
    Remove a the file after a successful download from the SFTP server when type
    is 'Download'. For type 'Upload' the source file is always removed after a
    successful upload.

.PARAMETER Tasks.Actions.Parameter.Option.ErrorWhen.PathIsNotFound
    Throw an error when the file to upload is not found. When Path is a folder
    this option is ignored, because a folder can be empty.

.PARAMETER Tasks.Actions.Parameter.Option.ErrorWhen.RemoveFailedPartialFiles
    When the upload process is interrupted, it is possible that files are not
    completely uploaded and that there are sill partial files present on the
    SFTP server or in the local folder.

    When RemoveFailedPartialFiles is TRUE these partial files will be removed
    before the script starts. When RemoveFailedPartialFiles is FALSE, manual
    intervention will be required to decide to still upload the partial file
    found in the local folder, to rename the partial file on the SFTP server,
    or to simply remove the partial file(s).

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

.PARAMETER PSSessionConfiguration
    The version of PowerShell on the remote endpoint as returned by
    Get-PSSessionConfiguration.
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
    [String]$PSSessionConfiguration = 'PowerShell.7',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Get or send SFTP file\$ScriptName",
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
        Write-Verbose "Import .json file '$ImportFile'"
        # Write-EventLog @EventVerboseParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        Write-Verbose 'Test .json file properties'

        try {
            @(
                'MaxConcurrentJobs', 'SendMail', 'ExportExcelFile', 'Tasks'
            ).where(
                { -not $file.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            #region Test SendMail
            @('To', 'When').Where(
                { -not $file.SendMail.$_ }
            ).foreach(
                { throw "Property 'SendMail.$_' not found" }
            )

            if ($file.SendMail.When -notMatch '^Never$|^Always$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                throw "Property 'SendMail.When' with value '$($file.SendMail.When)' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
            }
            #endregion

            #region Test ExportExcelFile
            @('When').Where(
                { -not $file.ExportExcelFile.$_ }
            ).foreach(
                { throw "Property 'ExportExcelFile.$_' not found" }
            )

            if ($file.ExportExcelFile.When -notMatch '^Never$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                throw "Property 'ExportExcelFile.When' with value '$($file.ExportExcelFile.When)' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
            }
            #endregion

            #region Test integer value
            try {
                [int]$MaxConcurrentJobs = $file.MaxConcurrentJobs
            }
            catch {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$($file.MaxConcurrentJobs)' is not supported."
            }
            #endregion

            $Tasks = $file.Tasks

            foreach ($task in $Tasks) {
                @(
                    'TaskName', 'Sftp', 'Actions'
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

                @('UserName').Where(
                    { -not $task.Sftp.Credential.$_ }
                ).foreach(
                    { throw "Property 'Tasks.Sftp.Credential.$_' not found" }
                )

                if (
                    $task.Sftp.Credential.Password -and
                    $task.Sftp.Credential.PasswordKeyFile
                ) {
                    throw "Property 'Tasks.Sftp.Credential.Password' and 'Tasks.Sftp.Credential.PasswordKeyFile' cannot be used at the same time"
                }

                if (
                    (-not $task.Sftp.Credential.Password) -and
                    (-not $task.Sftp.Credential.PasswordKeyFile)
                ) {
                    throw "Property 'Tasks.Sftp.Credential.Password' or 'Tasks.Sftp.Credential.PasswordKeyFile' not found"
                }

                if (-not $task.Actions) {
                    throw 'Tasks.Actions is missing'
                }

                foreach ($action in $task.Actions) {
                    @('Type', 'Parameter').Where(
                        { -not $action.$_ }
                    ).foreach(
                        { throw "Property 'Tasks.Actions.$_' not found" }
                    )

                    @(
                        'SftpPath', 'Option', 'PartialFileExtension'
                    ).Where(
                        { -not $action.Parameter.$_ }
                    ).foreach(
                        {
                            throw "Property 'Tasks.Actions.Parameter.$_' not found"
                        }
                    )

                    #region Test partial file extensions
                    @(
                        'PartialFileExtension'
                    ).Where(
                        { $action.Parameter.$_ -notLike '.*' }
                    ).foreach(
                        { throw "Property 'Tasks.Actions.Parameter.$_' needs to start with a dot. For example: '.txt', '.xml', ..." }
                    )
                    #endregion

                    #region Test file extensions
                    $action.Parameter.FileExtensions.Where(
                        { $_ -and ($_ -notLike '.*') }
                    ).foreach(
                        { throw "Property 'Tasks.Actions.Parameter.FileExtensions' needs to start with a dot. For example: '.txt', '.xml', ..." }
                    )
                    #endregion

                    #region Test boolean values
                    foreach (
                        $boolean in
                        @(
                            'OverwriteFile',
                            'RemoveFailedPartialFiles'
                        )
                    ) {
                        try {
                            $null = [Boolean]::Parse($action.Parameter.Option.$boolean)
                        }
                        catch {
                            throw "Property 'Tasks.Actions.Parameter.Option.$boolean' is not a boolean value"
                        }
                    }
                    #endregion

                    switch ($action.Type) {
                        'Download' {
                            @(
                                'Path'
                            ).Where(
                                { -not $action.Parameter.$_ }
                            ).foreach(
                                { throw "Property 'Tasks.Actions.Parameter.$_' not found" }
                            )

                            break
                        }
                        'Upload' {
                            @(
                                'Paths'
                            ).Where(
                                { -not $action.Parameter.$_ }
                            ).foreach(
                                { throw "Property 'Tasks.Actions.Parameter.$_' not found" }
                            )

                            #region Test boolean values
                            foreach (
                                $boolean in
                                @(
                                    'PathIsNotFound'
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

        #region Convert .json file
        Write-Verbose 'Convert .json file'

        try {
            foreach ($task in $Tasks) {
                #region Set ComputerName if there is none
                foreach ($action in $task.Actions) {
                    if (
                        (-not $action.Parameter.ComputerName) -or
                        ($action.Parameter.ComputerName -eq 'localhost') -or
                        ($action.Parameter.ComputerName -eq "$ENV:COMPUTERNAME.$env:USERDNSDOMAIN")
                    ) {
                        $action.Parameter.ComputerName = $env:COMPUTERNAME
                    }
                }
                #endregion

                #region Set secure string as password
                if ($task.Sftp.Credential.PasswordKeyFile) {
                    try {
                        $PasswordKeyFileStrings = Get-Content -LiteralPath $task.Sftp.Credential.PasswordKeyFile -ErrorAction Stop

                        if (-not $PasswordKeyFileStrings) {
                            throw 'File empty'
                        }

                        $task.Sftp.Credential.PasswordKeyFile = $PasswordKeyFileStrings
                    }
                    catch {
                        throw "Failed converting the task.Sftp.Credential.PasswordKeyFile '$($task.Sftp.Credential.PasswordKeyFile)' to an array: $_"
                    }

                    # Avoid password popup
                    $task.Sftp.Credential.Password = New-Object System.Security.SecureString
                }
                else {
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
                }
                #endregion

                #region Add environment variable SFTP UserName
                if (-not (
                        $params.String = Get-EnvironmentVariableValueHC -Name $task.Sftp.Credential.UserName)
                ) {
                    throw "Environment variable '`$ENV:$($task.Sftp.Credential.UserName)' in 'Sftp.Credential.UserName' not found on computer $ENV:COMPUTERNAME"
                }

                $task.Sftp.Credential.UserName = $params.String
                #endregion

                #region Convert SFTP path to '/path/'
                foreach ($action in $task.Actions) {
                    $action.Parameter.SftpPath = $action.Parameter.SftpPath -replace '\\', '/'
                    $action.Parameter.SftpPath = $action.Parameter.SftpPath.TrimEnd('/') + '/'
                    $action.Parameter.SftpPath = '/' + $action.Parameter.SftpPath.TrimStart('/')
                }
                #endregion

                #region Add properties
                $task.Actions | ForEach-Object {
                    $_ | Add-Member -NotePropertyMembers @{
                        Job = @{
                            Results = @()
                        }
                    }
                }
                #endregion
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
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
        $scriptBlock = {
            try {
                $action = $_

                #region Declare variables for code running in parallel
                if (-not $MaxConcurrentJobs) {
                    $task = $using:task
                    $PathItem = $using:PathItem
                    $PSSessionConfiguration = $using:PSSessionConfiguration
                    $EventVerboseParams = $using:EventVerboseParams
                }
                #endregion

                #region Create job parameters
                switch ($action.Type) {
                    'Upload' {
                        $invokeParams = @{
                            FilePath     = $PathItem.UploadScript
                            ArgumentList = $action.Parameter.Paths,
                            $task.Sftp.ComputerName,
                            $action.Parameter.SftpPath,
                            $task.Sftp.Credential.UserName,
                            $action.Parameter.PartialFileExtension,
                            $task.Sftp.Credential.Password,
                            $task.Sftp.Credential.PasswordKeyFile,
                            $action.Parameter.Option.OverwriteFile,
                            $action.Parameter.Option.ErrorWhen.PathIsNotFound,
                            $action.Parameter.Option.RemoveFailedPartialFiles,
                            $action.Parameter.FileExtensions
                        }

                        $M = "Start SFTP upload job '{0}' on '{1}' script '{10}' with arguments: Sftp.ComputerName '{2}' SftpPath '{3}' Sftp.UserName '{4}' PartialFileExtension '{5}' Option.OverwriteFile '{6}' Option.ErrorWhen.PathIsNotFound '{7}' RemoveFailedPartialFiles '{8}' Paths '{9}' FileExtensions '{11}'" -f
                        $task.TaskName,
                        $action.Parameter.ComputerName,
                        $invokeParams.ArgumentList[1],
                        $invokeParams.ArgumentList[2],
                        $invokeParams.ArgumentList[3],
                        $invokeParams.ArgumentList[4],
                        $invokeParams.ArgumentList[7],
                        $invokeParams.ArgumentList[8],
                        $invokeParams.ArgumentList[9],
                        $($invokeParams.ArgumentList[0] -join "', '"),
                        $invokeParams.FilePath,
                        $($invokeParams.ArgumentList[10] -join "', '")

                        Write-Verbose $M
                        # Write-EventLog @EventVerboseParams -Message $M

                        break
                    }
                    'Download' {
                        $invokeParams = @{
                            FilePath     = $PathItem.DownloadScript
                            ArgumentList = $action.Parameter.Path,
                            $task.Sftp.ComputerName,
                            $action.Parameter.SftpPath,
                            $task.Sftp.Credential.UserName,
                            $action.Parameter.PartialFileExtension,
                            $task.Sftp.Credential.Password,
                            $task.Sftp.Credential.PasswordKeyFile,
                            $action.Parameter.FileExtensions,
                            $action.Parameter.Option.OverwriteFile,
                            $action.Parameter.Option.RemoveFailedPartialFiles
                        }

                        $M = "Start SFTP download job '{0}' on '{1}' script '{2}' with arguments: Sftp.ComputerName '{3}' SftpPath '{4}' Sftp.UserName '{5}' PartialFileExtension '{6}' FileExtensions '{7}' Option.OverwriteFile '{8}' Option.RemoveFailedPartialFiles '{9}' Path '{10}'" -f
                        $task.TaskName,
                        $action.Parameter.ComputerName,
                        $invokeParams.FilePath,
                        $invokeParams.ArgumentList[1],
                        $invokeParams.ArgumentList[2],
                        $invokeParams.ArgumentList[3],
                        $invokeParams.ArgumentList[4],
                        $($invokeParams.ArgumentList[7] -join ', '),
                        $invokeParams.ArgumentList[8],
                        $invokeParams.ArgumentList[9],
                        $invokeParams.ArgumentList[0]
                        Write-Verbose $M
                        # Write-EventLog @EventVerboseParams -Message $M

                        break
                    }
                    Default {
                        throw "Tasks.Actions.Type '$_' not supported."
                    }
                }
                #endregion

                #region Start job
                $computerName = $action.Parameter.ComputerName

                $action.Job.Results += if (
                    $computerName -eq $ENV:COMPUTERNAME
                ) {
                    $params = $invokeParams.ArgumentList
                    & $invokeParams.FilePath @params
                }
                else {
                    $invokeParams += @{
                        ConfigurationName = $PSSessionConfiguration
                        ComputerName      = $computerName
                        ErrorAction       = 'Stop'
                    }
                    Invoke-Command @invokeParams
                }
                #endregion

                #region Get job results
                if ($action.Job.Results.Count -ne 0) {
                    $M = "Task '{0}' Type '{1}' SftpPath '{2}' ComputerName '{3}' Path '{4}': {5} job result{6}" -f
                    $task.TaskName, $action.Type, $action.Parameter.SftpPath,
                    $action.Parameter.ComputerName,
                    $(
                        if ($action.Parameter.Path) { $action.Parameter.Path }
                        else { $action.Parameter.Paths -join ', ' }
                    ),
                    $action.Job.Results.Count,
                    $(if ($action.Job.Results.Count -ne 1) { 's' })
                    Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                }
                #endregion
            }
            catch {
                $action.Job.Results += [PSCustomObject]@{
                    DateTime   = Get-Date
                    LocalPath  = $invokeParams.ArgumentList[0]
                    SftpPath   = $invokeParams.ArgumentList[2]
                    FileName   = $null
                    FileLength = $null
                    Downloaded = $false
                    Uploaded   = $false
                    Action     = $null
                    Error      = "General error: $_"
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

        foreach ($task in $Tasks) {
            Write-Verbose "Execute task '$($task.TaskName)' with $($task.Actions.Count) actions"

            $task.Actions | ForEach-Object @foreachParams
        }

        Write-Verbose 'All tasks finished'
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
        #region Counter
        $counter = @{
            Total = @{
                Errors          = $countSystemErrors
                UploadedFiles   = 0
                DownloadedFiles = 0
            }
        }
        #endregion

        #region Create system error html lists
        $countSystemErrors = (
            $Error.Exception.Message | Measure-Object
        ).Count

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

        $mailParams = @{}
        $htmlTable = @()
        $exportToExcel = @()

        $htmlTable += '<table>'

        foreach ($task in $Tasks) {
            $htmlTable += "
                <tr>
                    <th style=`"text-align: center; background-color: lightgray;`" colspan=`"4`">$($task.TaskName) - SFTP Server: $($task.SFTP.ComputerName)</th>
                </tr>
                <tr>
                    <th>ComputerName</th>
                    <th>Source</th>
                    <th>Destination</th>
                    <th>Result</th>
                </tr>"

            foreach ($action in $task.Actions) {
                #region Counter
                $counter.Action = @{
                    Errors          = 0
                    UploadedFiles   = 0
                    DownloadedFiles = 0
                }
                #endregion

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

                #region Create HTML table row
                $htmlTable += "
                        $(
                            if ($counter.Action.Errors) {
                                '<tr style="background-color: red">'
                            }
                            else {
                                '<tr>'
                            }
                        )
                        <td>
                            $($action.Parameter.ComputerName)
                        </td>
                        <td>
                            $(
                                if ($action.Type -eq 'Upload') {
                                    $action.Parameter.Paths -join ', '
                                }
                                else {
                                    $action.Parameter.SftpPath
                                }
                            )
                        </td>
                        <td>
                            $(
                                if ($action.Type -eq 'Download') {
                                    $action.Parameter.SftpPath
                                }
                                else {
                                    $action.Parameter.Path
                                }
                            )
                        </td>
                        <td>
                            $(
                                $result = if ($action.Type -eq 'Download') {
                                    "$($counter.Action.DownloadedFiles) downloaded"
                                }
                                else {
                                    "$($counter.Action.UploadedFiles) uploaded"
                                }

                                if ($counter.Action.Errors) {
                                    $result += ", {0} error{1}" -f
                                    $(
                                        $counter.Action.Errors
                                    ),
                                    $(
                                        if($counter.Action.Errors -ne 1) {'s'}
                                    )
                                }

                                $result
                            )
                        </td>
                    </tr>"
                #endregion

                #region Create Excel objects
                $exportToExcel += $action.Job.Results | Select-Object DateTime,
                @{
                    Name       = 'TaskName'
                    Expression = { $task.TaskName }
                },
                @{
                    Name       = 'Type'
                    Expression = { $action.Type }
                },
                @{
                    Name       = 'SFTP Server'
                    Expression = { $task.SFTP.ComputerName }
                },
                @{
                    Name       = 'ComputerName'
                    Expression = { $action.Parameter.ComputerName }
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
                    Name       = 'FileSize'
                    Expression = { $_.FileLength / 1KB }
                },
                @{
                    Name       = 'Successful'
                    Expression = {
                        if ($action.Type -eq 'Upload') {
                            $_.Uploaded
                        }
                        else {
                            $_.Downloaded
                        }
                    }
                },
                @{
                    Name       = 'Action'
                    Expression = { $_.Action -join ', ' }
                },
                Error
                #endregion
            }
        }

        $htmlTable += '</table>'
        $htmlTable = $htmlTable -join '<br>'

        #region Create Excel worksheet Overview
        $createExcelFile = $false

        if (
            (
                ($file.ExportExcelFile.When -eq 'OnlyOnError') -and
                ($counter.Total.Errors)
            ) -or
            (
                ($file.ExportExcelFile.When -eq 'OnlyOnErrorOrAction') -and
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
                AutoNameRange = $true
                Append        = $true
                AutoSize      = $true
                FreezeTopRow  = $true
                WorksheetName = 'Overview'
                TableName     = 'Overview'
                Verbose       = $false
            }

            $M = "Export {0} rows to Excel sheet '{1}'" -f
            $exportToExcel.Count, $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $exportToExcel | Export-Excel @excelParams -CellStyleSB {
                Param (
                    $WorkSheet,
                    $TotalRows,
                    $LastColumn
                )

                @($WorkSheet.Names['FileSize'].Style).ForEach(
                    { $_.NumberFormat.Format = '0.00\ \K\B' }
                )
            }

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
                ($file.SendMail.When -eq 'Always')
            ) -or
            (
                ($file.SendMail.When -eq 'OnlyOnError') -and
                ($counter.Total.Errors)
            ) -or
            (
                ($file.SendMail.When -eq 'OnlyOnErrorOrAction') -and
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
            To             = $file.SendMail.To
            Message        = "
                                $systemErrorsHtmlList
                                <p>Summary of all SFTP actions.</p>
                                $summaryHtmlTable
                                <p>Action details.</p>
                                $htmlTable"

            LogFolder      = $LogParams.LogFolder
            Header         = $ScriptName
            EventLogSource = $ScriptName
            Save           = $LogFile + ' - Mail.html'
            ErrorAction    = 'Stop'
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