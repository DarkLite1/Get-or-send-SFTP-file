#Requires -Version 7
#Requires -Modules ImportExcel
#Requires -Modules Toolbox.EventLog, Toolbox.HTML

<#
.SYNOPSIS
    Download files from an SFTP server or upload files to an SFTP server.

.DESCRIPTION
    Read an input file that contains all the parameters for this script.

    The computer that is running the SFTP code should have the module 'Posh-SSH'
    installed.

    Actions run in parallel when MaxConcurrentJobs is more than 1. Tasks
    run sequentially.

.PARAMETER ImportFile
    A .JSON file that contains all the parameters used by the script.

.PARAMETER Tasks
    Each task is a collection of upload or download actions. Tasks run one after
    the other, meaning they run sequentially.

    Actions run in parallel based on the "MaxConcurrentJobs" argument.

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
    to or from an SFTP server. The action in a task are run in parallel when
    "MaxConcurrentJobs" is more than 1.

.PARAMETER Tasks.Actions.ComputerName
    The client where the SFTP code will be executed. This machine needs to
    have the module 'Posh-SSH' installed.

.PARAMETER Tasks.Actions.Paths
    Combination of Source and Destination where either can be an SFTP path,
    indicated with "sftp:\the path" and a local or SMB path.

.PARAMETER Tasks.Option.OverwriteFile
    Overwrite files on the SFTP server when they already exist.

.PARAMETER Tasks.Option.FileExtensions
    If blank, all files are treated. If used, only those files meeting the
    extension filter will be moved.

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
    [HashTable]$ScriptPath = @{
        MoveFile = "$PSScriptRoot\Move file.ps1"
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
        $scriptPathItem = @{}

        $ScriptPath.GetEnumerator().ForEach(
            {
                try {
                    $key = $_.Key
                    $value = $_.Value

                    $params = @{
                        Path        = $value
                        ErrorAction = 'Stop'
                    }
                    $scriptPathItem[$key] = (Get-Item @params).FullName
                }
                catch {
                    throw "ScriptPath.$key '$value' not found"
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
                    'TaskName', 'Sftp', 'Actions', 'Option'
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

                #region Test boolean values
                foreach (
                    $boolean in
                    @(
                        'OverwriteFile'
                    )
                ) {
                    try {
                        $null = [Boolean]::Parse($task.Option.$boolean)
                    }
                    catch {
                        throw "Property 'Tasks.Option.$boolean' is not a boolean value"
                    }
                }
                #endregion

                #region Test file extensions
                $task.Option.FileExtensions.Where(
                    { $_ -and ($_ -notLike '.*') }
                ).foreach(
                    { throw "Property 'Tasks.Option.FileExtensions' needs to start with a dot. For example: '.txt', '.xml', ..." }
                )
                #endregion

                if (-not $task.Actions) {
                    throw 'Tasks.Actions is missing'
                }

                #region Test unique ComputerName
                $task.Actions | Group-Object -Property {
                    $_.ComputerName
                } |
                Where-Object { $_.Count -ge 2 } | ForEach-Object {
                    throw "Duplicate 'Tasks.Actions.ComputerName' found: $($_.Name)"
                }
                #endregion

                foreach ($action in $task.Actions) {
                    if ($action.PSObject.Properties.Name -notContains 'ComputerName') {
                        throw "Property 'Tasks.Actions.ComputerName' not found"
                    }

                    @('Paths').Where(
                        { -not $action.$_ }
                    ).foreach(
                        { throw "Property 'Tasks.Actions.$_' not found" }
                    )

                    foreach ($path in $action.Paths) {
                        @(
                            'Source', 'Destination'
                        ).Where(
                            { -not $path.$_ }
                        ).foreach(
                            {
                                throw "Property 'Tasks.Actions.Paths.$_' not found"
                            }
                        )

                        if (
                            (
                                ($path.Source -like '*/*') -and
                                ($path.Destination -like '*/*')
                            ) -or
                            (
                                ($path.Source -like '*\*') -and
                                ($path.Destination -like '*\*')
                            ) -or
                            (
                                ($path.Source -like 'sftp*') -and
                                ($path.Destination -like 'sftp*')
                            ) -or
                            (
                                -not (
                                    ($path.Source -like 'sftp:/*') -or
                                    ($path.Destination -like 'sftp:/*')
                                )
                            )
                        ) {
                            throw "Property 'Tasks.Actions.Paths.Source' and 'Tasks.Actions.Paths.Destination' needs to have one SFTP path ('sftp:/....') and one folder path (c:\... or \\server$\...). Incorrect values: Source '$($path.Source)' Destination '$($path.Destination)'"
                        }
                    }

                    #region Test unique Source Destination
                    $action.Paths | Group-Object -Property 'Source' |
                    Where-Object { $_.Count -ge 2 } | ForEach-Object {
                        throw "Duplicate 'Tasks.Actions.Paths.Source' found: '$($_.Name)'. Use separate Tasks to run them sequentially instead of in Actions, which is ran in parallel"
                    }
                    #endregion

                    #region Test unique Source Destination
                    $action.Paths | Group-Object -Property 'Destination' |
                    Where-Object { $_.Count -ge 2 } | ForEach-Object {
                        throw "Duplicate 'Tasks.Actions.Paths.Destination' found: '$($_.Name)'. Use separate Tasks to run them sequentially instead of in Actions, which is ran in parallel"
                    }
                    #endregion
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
                #region Set SFTP UserName
                if (-not (
                        $params.String = Get-EnvironmentVariableValueHC -Name $task.Sftp.Credential.UserName)
                ) {
                    throw "Environment variable '`$ENV:$($task.Sftp.Credential.UserName)' in 'Sftp.Credential.UserName' not found on computer $ENV:COMPUTERNAME"
                }

                $task.Sftp.Credential.UserName = $params.String
                #endregion

                #region Set SFTP password
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

                foreach ($action in $task.Actions) {
                    #region Set ComputerName
                    if (
                        (-not $action.ComputerName) -or
                        ($action.ComputerName -eq 'localhost') -or
                        ($action.ComputerName -eq "$ENV:COMPUTERNAME.$env:USERDNSDOMAIN")
                    ) {
                        $action.ComputerName = $env:COMPUTERNAME
                    }
                    #endregion

                    #region Convert Paths
                    foreach ($path in $action.Paths) {
                        if ($path.Source -like 'sftp*') {
                            $path.Source = $path.Source.TrimEnd('/') + '/'
                        }
                        if ($path.Destination -like 'sftp*') {
                            $path.Destination = $path.Destination.TrimEnd('/') + '/'
                        }
                    }
                    #endregion

                    #region Add properties
                    $action | Add-Member -NotePropertyMembers @{
                        Job = @{
                            Results = @()
                            Error   = $null
                        }
                    }
                    #endregion
                }
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
                    $MaxConcurrentJobs = $using:MaxConcurrentJobs
                    $scriptPathItem = $using:scriptPathItem
                    $PSSessionConfiguration = $using:PSSessionConfiguration
                    $EventVerboseParams = $using:EventVerboseParams
                }
                #endregion

                #region Create job parameters
                $invokeParams = @{
                    FilePath     = $scriptPathItem.MoveFile
                    ArgumentList = $task.Sftp.ComputerName,
                    $task.Sftp.Credential.UserName,
                    $action.Paths,
                    $MaxConcurrentJobs,
                    $task.Sftp.Credential.Password,
                    $task.Sftp.Credential.PasswordKeyFile,
                    $task.Option.FileExtensions,
                    $task.Option.OverwriteFile
                }

                $M = "Start SFTP job '{6}' on '{7}' with: Sftp.ComputerName '{0}' Sftp.UserName '{1}' Paths {2} MaxConcurrentJobs '{3}' FileExtensions '{4}' OverwriteFile '{5}'" -f
                $invokeParams.ArgumentList[0],
                $invokeParams.ArgumentList[1],
                $(
                    $invokeParams.ArgumentList[2].foreach(
                        { "Source '$($_.Source)' Destination '$($_.Destination)'" }
                    ) -join ', '
                ),
                $invokeParams.ArgumentList[3],
                $($invokeParams.ArgumentList[6] -join ', '),
                $invokeParams.ArgumentList[7],
                $task.TaskName,
                $action.ComputerName

                Write-Verbose $M
                # Write-EventLog @EventVerboseParams -Message $M
                #endregion

                #region Start job
                $computerName = $action.ComputerName

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
                    $M = "Job result '{6}' on '{7}' with: Sftp.ComputerName '{0}' Sftp.UserName '{1}' Paths {2} MaxConcurrentJobs '{3}' FileExtensions '{4}' OverwriteFile '{5}': {8} result{9}" -f
                    $invokeParams.ArgumentList[0],
                    $invokeParams.ArgumentList[1],
                    $(
                        $invokeParams.ArgumentList[2].foreach(
                            { "Source '$($_.Source)' Destination '$($_.Destination)'" }
                        ) -join ', '
                    ),
                    $invokeParams.ArgumentList[3],
                    $($invokeParams.ArgumentList[6] -join ', '),
                    $invokeParams.ArgumentList[7],
                    $task.TaskName,
                    $action.ComputerName,
                    $action.Job.Results.Count,
                    $(if ($action.Job.Results.Count -ne 1) { 's' })
                    Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                }
                #endregion
            }
            catch {
                $action.Job.Error = $_
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
                MovedFiles = 0
                Errors     = $countSystemErrors
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
            #region Create HTML table header
            $htmlTable += "
                <tr style=`"background-color: lightgrey;`">
                    <th style=`"text-align: center;`" colspan=`"2`">$($task.TaskName)</th>
                    <th>sftp:/$($task.SFTP.ComputerName)</th>
                </tr>
                <tr>
                    <th>Source</th>
                    <th>Destination</th>
                    <th>Result</th>
                </tr>"
            #endregion

            foreach ($action in $task.Actions) {
                #region Counter
                $counter.Action = @{
                    MovedFiles = 0
                    Errors     = 0
                }

                $counter.Action.MovedFiles = $action.Job.Results.Where(
                    { -not $_.Error }).Count
                $counter.Action.Errors = $action.Job.Results.Where(
                    { $_.Error }).Count
                $counter.Action.Errors += $action.Job.Error.Count

                $counter.Total.Errors += $counter.Action.Errors
                $counter.Total.MovedFiles += $counter.Action.MovedFiles
                #endregion

                #region Log errors
                if ($counter.Action.Errors) {
                    $action.Job.Results.Where(
                        { $_.Error }
                    ).foreach(
                        {
                            $M = "Error for TaskName '$($task.TaskName)' Sftp.ComputerName '$($task.Sftp.ComputerName)' ComputerName '$($action.ComputerName) Source '$($_.Source)' Destination '$($_.Destination)' FileName '$($_.FileName)': $($_.Error)"
                            Write-Warning $M
                            Write-EventLog @EventErrorParams -Message $M
                        }
                    )
                }
                #endregion

                #region Create HTML Error row
                if ($action.Job.Error) {
                    $htmlTable += "
                    <tr style=`"background-color: #f78474;`">
                        <td colspan=`"3`">ERROR: $($action.Job.Error)</td>
                    </tr>"

                    $M = "Error for TaskName '$($task.TaskName)' Sftp.ComputerName '$($task.Sftp.ComputerName)' ComputerName '$($action.ComputerName) {0}: $($action.Job.Error)" -f
                    $(
                        $action.Paths.ForEach(
                            {
                                "Source '$($_.Source)' Destination '$($_.Destination)'"
                            }
                        )
                    )
                    Write-Warning $M
                    Write-EventLog @EventErrorParams -Message $M
                }
                #endregion

                foreach ($path in $action.Paths) {
                    #region Counter
                    $counter.Path = @{
                        MovedFiles = 0
                        Errors     = 0
                    }

                    $counter.Path.Errors += $action.Job.Results.Where(
                        {
                        ($_.Error) -and
                        ($_.Source -eq $path.Source) -and
                        ($_.Destination -eq $path.Destination)
                        }).Count

                    $counter.Path.MovedFiles += $action.Job.Results.Where(
                        {
                        (-not $_.Error) -and
                        ($_.Source -eq $path.Source) -and
                        ($_.Destination -eq $path.Destination)
                        }).Count
                    #endregion

                    #region Create HTML table row
                    $htmlTable += "
                        $(
                            if (
                                $counter.Action.Errors -or
                                $counter.Path.Errors
                            ) {
                                '<tr style="background-color: #f78474">'
                            }
                            else {
                                '<tr>'
                            }
                        )
                        <td>
                            $($path.Source)
                        </td>
                        <td>
                            $($path.Destination)
                        </td>
                        <td>
                            $(
                                $result = "$($counter.Path.MovedFiles) moved"

                                if ($counter.Path.Errors) {
                                    $result += ", {0} error{1}" -f
                                    $(
                                        $counter.Path.Errors
                                    ),
                                    $(
                                        if($counter.Path.Errors -ne 1) {'s'}
                                    )
                                }

                                $result
                            )
                        </td>
                    </tr>"
                    #endregion
                }

                #region Create HTML Action summary row
                $htmlTable += "
                <tr>
                    <th colspan=`"2`"></th>
                    <th>$($counter.Action.MovedFiles) moved on $($action.ComputerName)</th>
                </tr>"
                #endregion

                #region Create Excel objects
                $exportToExcel += $action.Job.Results |
                Select-Object 'DateTime',
                @{
                    Name       = 'TaskName'
                    Expression = { $task.TaskName }
                },
                @{
                    Name       = 'SftpServer'
                    Expression = { $task.Sftp.ComputerName }
                },
                @{
                    Name       = 'ComputerName'
                    Expression = { $action.ComputerName }
                },
                'Source',
                'Destination',
                'FileName',
                @{
                    Name       = 'FileSize'
                    Expression = { $_.FileLength / 1KB }
                },
                @{
                    Name       = 'Action'
                    Expression = { $_.Action -join ', ' }
                },
                'Error'
                #endregion
            }
        }

        $htmlTable += '</table>'

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
                    ($counter.Total.Errors) -or ($counter.Total.MovedFiles)
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

        if ($counter.Total.MovedFiles) {
            $mailParams.Subject += "$($counter.Total.MovedFiles) moved"
        }
        if ($counter.Total.Errors) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += "{0} error{1}" -f
            $counter.Total.Errors,
            $(if ($counter.Total.Errors -ne 1) { 's' })
        }

        $mailParams.Subject = $mailParams.Subject -join ', '

        if (-not $mailParams.Subject) {
            $mailParams.Subject = 'Nothing done'
        }
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
                ($counter.Total.MovedFiles) -or ($counter.Total.Errors)
                )
            )
        ) {
            $sendMailToUser = $true
        }
        #endregion

        #region Send mail
        $mailParams += @{
            To             = $file.SendMail.To
            Message        = "
                            $systemErrorsHtmlList
                            <p>Find a summary of all SFTP actions below:</p>
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