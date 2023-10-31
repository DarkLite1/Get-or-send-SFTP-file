#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $realCmdLet = @{
        StartJob      = Get-Command Start-Job
        InvokeCommand = Get-Command Invoke-Command
    }

    $testInputFile = @{
        MaxConcurrentJobs = 1
        Tasks             = @(
            @{
                TaskName        = 'App x'
                Sftp            = @{
                    ComputerName = 'PC1'
                    Credential   = @{
                        UserName = 'envVarBob'
                        Password = 'envVarPasswordBob'
                    }
                }
                Actions         = @(
                    @{
                        Type      = 'Upload'
                        Parameter = @{
                            SftpPath     = '/SFTP/folder/'
                            ComputerName = 'localhost'
                            Path         = @(
                                (New-Item 'TestDrive:\a.txt').FullName
                                (New-Item 'TestDrive:\b.txt').FullName
                            )
                            Option       = @{
                                OverwriteFile        = $false
                                RemoveFileAfterwards = $false
                                ErrorWhen            = @{
                                    PathIsNotFound     = $true
                                    SftpPathIsNotFound = $false
                                }
                            }
                        }
                    }
                    @{
                        Type      = 'Download'
                        Parameter = @{
                            SftpPath     = '/SFTP/folder/'
                            ComputerName = 'localhost'
                            Path         = (New-Item 'TestDrive:\d' -ItemType Directory).FullName
                            Option       = @{
                                OverwriteFile        = $false
                                RemoveFileAfterwards = $false
                                ErrorWhen            = @{
                                    PathIsNotFound     = $true
                                    SftpPathIsNotFound = $false
                                }
                            }
                        }
                    }
                )
                SendMail        = @{
                    To   = 'bob@contoso.com'
                    When = 'Always'
                }
                ExportExcelFile = @{
                    When = 'OnlyOnErrorOrAction'
                }
            }
        )
    }

    $testData = @(
        [PSCustomObject]@{
            Path     = $testInputFile.Tasks[0].Actions[0].Parameter.Path[0]
            DateTime = Get-Date
            Uploaded = $true
            Action   = @('file uploaded', 'file removed')
            Error    = $null
        }     
        [PSCustomObject]@{
            Path     = $testInputFile.Tasks[0].Actions[0].Parameter.Path[1]
            DateTime = Get-Date
            Uploaded = $true
            Action   = @('file uploaded', 'file removed')
            Error    = $null
        }
        [PSCustomObject]@{
            Path       = 'sftp file.txt'
            DateTime   = Get-Date
            Downloaded = $true
            Action     = @('file downloaded', 'file removed')
            Error      = $null
        }
    )

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        Path        = @{
            UploadScript   = (New-Item 'TestDrive:/u.ps1' -ItemType 'File').FullName
            DownloadScript = (New-Item 'TestDrive:/d.ps1' -ItemType 'File').FullName
        }
        LogFolder   = (New-Item 'TestDrive:/log' -ItemType Directory).FullName
        ScriptAdmin = 'admin@contoso.com'
    }

    Function Get-EnvironmentVariableValueHC {
        Param(
            [String]$Name
        )
    }
    
    Mock Get-EnvironmentVariableValueHC {
        'bobUserName'
    } -ParameterFilter {
        $Name -eq $testInputFile.Tasks[0].Sftp.Credential.UserName
    }
    Mock Get-EnvironmentVariableValueHC {
        'bobPassword'
    } -ParameterFilter {
        $Name -eq $testInputFile.Tasks[0].Sftp.Credential.Password
    }
    Mock ConvertTo-SecureString {
        'bobPasswordEncrypted'
    } -ParameterFilter {
        $String -eq 'bobPassword'
    }
    Mock Start-Job {
        & $realCmdLet.InvokeCommand -Scriptblock { 
            $using:testData[0]
            $using:testData[1]
        } -AsJob -ComputerName $env:COMPUTERNAME
    } -ParameterFilter {
        $FilePath -eq $testParams.Path.UploadScript
    }
    Mock Start-Job {
        & $realCmdLet.InvokeCommand -Scriptblock { 
            $using:testData[2]
        } -AsJob -ComputerName $env:COMPUTERNAME
    } -ParameterFilter {
        $FilePath -eq $testParams.Path.DownloadScript
    }
    Mock Invoke-Command {
        & $realCmdLet.InvokeCommand -Scriptblock { 
            
        } -AsJob -ComputerName $env:COMPUTERNAME
    }
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and 
            ($Priority -eq 'High') -and 
            ($Subject -eq 'FAILURE')
        }    
    }
    It 'the log folder cannot be created' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the file is not found' {
        It 'Path.UploadScript' {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.Path.UploadScript = 'c:\upDoesNotExist.ps1'
            
            $testInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*Path.UploadScript 'c:\upDoesNotExist.ps1' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'Path.DownloadScript' {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.Path.DownloadScript = 'c:\downDoesNotExist.ps1'
            
            $testInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*Path.DownloadScript 'c:\downDoesNotExist.ps1' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.ImportFile = 'nonExisting.json'
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'MaxConcurrentJobs', 'Tasks'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'MaxConcurrentJobs not a number' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.MaxConcurrentJobs = 'wrong'
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'MaxConcurrentJobs' needs to be a number, the value 'wrong' is not supported.*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.<_> not found' -ForEach @(
                'TaskName', 'Sftp', 'Actions', 'SendMail', 'ExportExcelFile'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.TaskName not found' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].TaskName = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.TaskName' not found*")
                }
            }
            It 'Tasks.Sftp.<_> not found' -ForEach @(
                'ComputerName', 'Credential'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Sftp.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Sftp.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.Sftp.Credential.<_> not found' -ForEach @(
                'UserName', 'Password'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Sftp.Credential.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Sftp.Credential.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.Actions.<_> not found' -ForEach @(
                'Type', 'Parameter'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Actions[0].$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Actions.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context "Tasks.Actions.Type is 'Download'" {
                It 'Tasks.Actions.Parameter.<_> not found' -ForEach @(
                    'SftpPath', 'ComputerName', 'Path', 'Option'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Actions[1].Parameter.$_ = $null
    
                    $testNewInputFile | ConvertTo-Json -Depth 7 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Actions.Parameter.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Tasks.Actions.Parameter.Option.<_> not a boolean' -ForEach @(
                    'OverwriteFile', 
                    'RemoveFileAfterwards'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Actions[1].Parameter.Option.$_ = $null
    
                    $testNewInputFile | ConvertTo-Json -Depth 7 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Actions.Parameter.Option.$_' is not a boolean value*")
                    }
                }
                It 'Tasks.Actions.Parameter.Option.ErrorWhen.<_> not a boolean' -ForEach @(
                    'PathIsNotFound',
                    'SftpPathIsNotFound'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Actions[1].Parameter.Option.ErrorWhen.$_ = $null
    
                    $testNewInputFile | ConvertTo-Json -Depth 7 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Actions.Parameter.Option.ErrorWhen.$_' is not a boolean value*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context "Tasks.Actions.Type is 'Upload'" {
                It 'Tasks.Actions.Parameter.<_> not found' -ForEach @(
                    'SftpPath', 'ComputerName', 'Path', 'Option'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Actions[0].Parameter.$_ = $null
    
                    $testNewInputFile | ConvertTo-Json -Depth 7 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Actions.Parameter.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Tasks.Actions.Parameter.Option.<_> not a boolean' -ForEach @(
                    'OverwriteFile', 
                    'RemoveFileAfterwards'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Actions[0].Parameter.Option.$_ = $null
    
                    $testNewInputFile | ConvertTo-Json -Depth 7 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Actions.Parameter.Option.$_' is not a boolean value*")
                    }
                }
                It 'Tasks.Actions.Parameter.Option.ErrorWhen.<_> not a boolean' -ForEach @(
                    'PathIsNotFound',
                    'SftpPathIsNotFound'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Actions[0].Parameter.Option.ErrorWhen.$_ = $null
    
                    $testNewInputFile | ConvertTo-Json -Depth 7 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Actions.Parameter.Option.ErrorWhen.$_' is not a boolean value*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            It 'Tasks.SendMail.<_> not found' -ForEach @(
                'To', 'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].SendMail.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.SendMail.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.ExportExcelFile.<_> not found' -ForEach @(
                'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].ExportExcelFile.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.ExportExcelFile.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.ExportExcelFile.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].ExportExcelFile.When = 'wrong'
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.ExportExcelFile.When' with value 'wrong' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.SendMail.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].SendMail.When = 'wrong'
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.SendMail.When' with value 'wrong' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.Name is not unique' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks = @(
                    $testInputFile.Tasks[0]
                    $testInputFile.Tasks[0]
                )
                $testNewInputFile.Tasks[0].TaskName = 'Name1'
                $testNewInputFile.Tasks[1].TaskName = 'Name1'
    
                $testNewInputFile | ConvertTo-Json -Depth 7 | 
                Out-File @testOutParams
                    
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.TaskName' with value 'Name1' is not unique*")
                }
            }
        }
        It 'the SFTP password is not found in the environment variables' {
            Mock Get-EnvironmentVariableValueHC {
                'user'
            } -ParameterFilter {
                $Name -eq $testInputFile.Tasks[0].Sftp.Credential.UserName
            }
            Mock Get-EnvironmentVariableValueHC -ParameterFilter {
                $Name -eq $testInputFile.Tasks[0].Sftp.Credential.Password
            }

            $testInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams

            .$testScript @testParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*Environment variable '`$ENV:$($testInputFile.Tasks[0].Sftp.Credential.Password)' in 'Sftp.Credential.Password' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'the SFTP user name is not found in the environment variables' {
            Mock Get-EnvironmentVariableValueHC {
                'pass'
            } -ParameterFilter {
                $Name -eq $testInputFile.Tasks[0].Sftp.Credential.Password
            }
            Mock Get-EnvironmentVariableValueHC -ParameterFilter {
                $Name -eq $testInputFile.Tasks[0].Sftp.Credential.UserName
            }

            $testInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams

            .$testScript @testParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*Environment variable '`$ENV:$($testInputFile.Tasks[0].Sftp.Credential.UserName)' in 'Sftp.Credential.UserName' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
    }
}
Describe 'execute the SFTP script' {
    BeforeAll {
        $testJobArguments = {
            ($FilePath -eq $testParams.Path.UploadScript) -and
            ($ArgumentList[0][0] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Path[0]) -and
            ($ArgumentList[0][1] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Path[1]) -and
            ($ArgumentList[1] -eq $testInputFile.Tasks[0].Sftp.ComputerName) -and
            ($ArgumentList[2] -eq $testInputFile.Tasks[0].Actions[0].Parameter.SftpPath) -and
            ($ArgumentList[3] -eq 'bobUserName') -and
            ($ArgumentList[4] -eq 'bobPasswordEncrypted') -and
            ($ArgumentList[5] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Option.OverwriteFile) -and
            ($ArgumentList[6] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Option.RemoveFileAfterwards) -and
            ($ArgumentList[7] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Option.ErrorWhen.PathIsNotFound)
        }
    }
    It 'with Invoke-Command when Tasks.Actions.Parameter.ComputerName is not the localhost' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Actions[0].Parameter.ComputerName = 'PC1'

        $testNewInputFile | ConvertTo-Json -Depth 7 | 
        Out-File @testOutParams
            
        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter $testJobArguments
    }
    It 'with Start-Job when Tasks.Actions.Parameter.ComputerName is the localhost' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Actions[0].Parameter.ComputerName = 'localhost'

        $testNewInputFile | ConvertTo-Json -Depth 7 | 
        Out-File @testOutParams
            
        .$testScript @testParams

        Should -Invoke Start-Job -Times 1 -Exactly -ParameterFilter $testJobArguments
    }
}
Describe 'when the SFTP script runs successfully' {
    BeforeAll {
        $testInputFile | ConvertTo-Json -Depth 7 | 
        Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'create an Excel file' {
        BeforeAll {
            $testExportedExcelRows = $testData | 
            Select-Object Path, DateTime, @{
                Name       = 'Action'
                Expression = { $_.Action -join ', ' }
            }, Error

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter "* - $($testInputFile.Tasks[0].TaskName) - Log.xlsx"

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'in the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Path -eq $testRow.Path
                }
                $actualRow.DateTime.ToString('yyyyMMdd') | 
                Should -Be $testRow.DateTime.ToString('yyyyMMdd')
                $actualRow.Action | Should -Be $testRow.Action
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Type | Should -Match 'Upload|Download'
            }
        }
    }
    Context 'send an e-mail' {
        It 'with attachment to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.Tasks[0].SendMail.To) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq '2 uploaded, 1 downloaded') -and
            ($Attachments -like '*- Log.xlsx') -and
            ($Message -like "*table*$($testInputFile.Tasks[0].TaskName)*SFTP Server*$($testInputFile.Tasks[0].Sftp.ComputerName)*SFTP User name*bobUserName*Total files uploaded*2*UPLOAD FILES TO THE SFTP SERVER*SFTP path*$($testInputFile.Tasks[0].Actions[0].SftpPath)*$($testInputFile.Tasks[0].Actions[0].ComputerName)*Path*$($testInputFile.Tasks[0].Actions[0].Parameter.Path[0])*$($testInputFile.Tasks[0].Actions[0].Parameter.Path[1])*DOWNLOAD FILES FROM THE SFTP SERVER*")
            }
        }
    }
}
Describe 'ExportExcelFile.When' {
    Context 'create no Excel file' {
        It "'Never'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'Never'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
        It "'OnlyOnError' and no errors are found" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnError'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are no errors and no actions" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                   
                } -AsJob -ComputerName $env:COMPUTERNAME
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnErrorOrAction'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
    }
    Context 'create an Excel file' {
        It "'OnlyOnError' and there are errors" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path     = 'a'
                        DateTime = Get-Date
                        Action   = @()
                        Error    = 'oops'
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnError'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are actions but no errors" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path     = 'a'
                        DateTime = Get-Date
                        Uploaded = $true
                        Action   = @('upload')
                        Error    = $null
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnErrorOrAction'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are errors but no actions" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path     = 'a'
                        Uploaded = $false
                        DateTime = Get-Date
                        Action   = @()
                        Error    = 'oops'
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnErrorOrAction'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
    }
}
Describe 'SendMail.When' {
    BeforeAll {
        $testParamFilter = @{
            ParameterFilter = { $To -eq $testNewInputFile.Tasks[0].SendMail.To }
        }
    }
    Context 'send no e-mail to the user' {
        It "'Never'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'Never'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Not -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnError' and no errors are found" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnError'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Not -Invoke Send-MailHC
        }
        It "'OnlyOnErrorOrAction' and there are no errors and no actions" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                   
                } -AsJob -ComputerName $env:COMPUTERNAME
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnErrorOrAction'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Not -Invoke Send-MailHC
        }
    }
    Context 'send an e-mail to the user' {
        It "'OnlyOnError' and there are errors" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path     = 'a'
                        DateTime = Get-Date
                        Action   = @()
                        Error    = 'oops'
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnError'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are actions but no errors" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path     = 'a'
                        DateTime = Get-Date
                        Uploaded = $true
                        Action   = @('upload')
                        Error    = $null
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnErrorOrAction'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are errors but no actions" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path     = 'a'
                        DateTime = Get-Date
                        Action   = @()
                        Error    = 'oops'
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnErrorOrAction'
    
            $testNewInputFile | ConvertTo-Json -Depth 7 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Invoke Send-MailHC @testParamFilter
        }
    }
}