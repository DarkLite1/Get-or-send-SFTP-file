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
                Task            = @{
                    Name                  = 'App x'
                    ExecuteOnComputerName = 'localhost'
                }
                Sftp            = @{
                    ComputerName = 'PC1'
                    Path         = '/SFTP/folder/'
                    Credential   = @{
                        UserName = 'envVarBob'
                        Password = 'envVarPasswordBob'
                    }
                }
                Upload          = @{
                    Path   = @(
                        (New-Item 'TestDrive:\a.txt').FullName
                        (New-Item 'TestDrive:\b.txt').FullName
                    )
                    Option = @{
                        OverwriteFileOnSftpServer = $false
                        RemoveFileAfterUpload     = $false
                        ErrorWhen                 = @{
                            UploadPathIsNotFound = $true
                        }
                    }
                }
                SendMail        = @{
                    To   = 'bob@contoso.com'
                    When = 'Always'
                }
                ExportExcelFile = @{
                    When = 'OnlyOnErrorOrUpload'
                }
            }
        )
    }

    $testData = @(
        [PSCustomObject]@{
            Path       = $testInputFile.Tasks[0].Upload.Path[0]
            UploadedOn = Get-Date
            Action     = @('file uploaded', 'file removed')
            Error      = $null
        }     
        [PSCustomObject]@{
            Path       = $testInputFile.Tasks[0].Upload.Path[1]
            UploadedOn = Get-Date
            Action     = @('file uploaded', 'file removed')
            Error      = $null
        }
    )

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName     = 'Test (Brecht)'
        ImportFile     = $testOutParams.FilePath
        SftpScriptPath = (New-Item 'TestDrive:/s.ps1' -ItemType 'File').FullName
        LogFolder      = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin    = 'admin@conotoso.com'
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
            $using:testData
        } -AsJob -ComputerName $env:COMPUTERNAME
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
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
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
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
                'Task', 'Sftp', 'Upload', 'SendMail', 'ExportExcelFile'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
            It 'Tasks.Task.<_> not found' -ForEach @(
                'Name', 'ExecuteOnComputerName'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Task.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Task.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.Sftp.<_> not found' -ForEach @(
                'ComputerName', 'Path', 'Credential'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Sftp.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
            It 'Tasks.Upload.<_> not found' -ForEach @(
                'Path', 'Option'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Upload.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Upload.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.Upload.Option.<_> not a boolean' -ForEach @(
                'OverwriteFileOnSftpServer', 
                'RemoveFileAfterUpload'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Upload.Option.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Upload.Option.$_' is not a boolean value*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.Upload.Option.ErrorWhen.<_> not a boolean' -ForEach @(
                'UploadPathIsNotFound'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Upload.Option.ErrorWhen.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Upload.Option.ErrorWhen.$_' is not a boolean value*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.SendMail.<_> not found' -ForEach @(
                'To', 'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].SendMail.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.ExportExcelFile.When' with value 'wrong' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrUpload'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.SendMail.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].SendMail.When = 'wrong'
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.SendMail.When' with value 'wrong' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrUpload'*")
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
                $testNewInputFile.Tasks[0].Task.Name = 'Name1'
                $testNewInputFile.Tasks[1].Task.Name = 'Name1'
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Tasks.Task.Name' with value 'Name1' is not unique*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
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

            $testInputFile | ConvertTo-Json -Depth 5 | 
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

            $testInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams

            .$testScript @testParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*Environment variable '`$ENV:$($testInputFile.Tasks[0].Sftp.Credential.UserName)' in 'Sftp.Credential.UserName' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'the SFTP script is not found' {
            $testNewParams = $testParams.Clone()
            $testNewParams.SftpScriptPath = 'c:\doesNotExist.ps1'
            
            $testInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams

            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*SftpScriptPath 'c:\doesNotExist.ps1' not found*")
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
            ($FilePath -eq $testParams.SftpScriptPath) -and
            ($ArgumentList[0][0] -eq $testInputFile.Tasks[0].Upload.Path[0]) -and
            ($ArgumentList[0][1] -eq $testInputFile.Tasks[0].Upload.Path[1]) -and
            ($ArgumentList[1] -eq $testInputFile.Tasks[0].Sftp.ComputerName) -and
            ($ArgumentList[2] -eq $testInputFile.Tasks[0].Sftp.Path) -and
            ($ArgumentList[3] -eq 'bobUserName') -and
            ($ArgumentList[4] -eq 'bobPasswordEncrypted') -and
            ($ArgumentList[5] -eq $testInputFile.Tasks[0].Upload.Option.OverwriteFileOnSftpServer) -and
            ($ArgumentList[6] -eq $testInputFile.Tasks[0].Upload.Option.RemoveFileAfterUpload) -and
            ($ArgumentList[7] -eq $testInputFile.Tasks[0].Upload.Option.ErrorWhen.UploadPathIsNotFound)
        }
    }
    It 'with Invoke-Command when ExecuteOnComputerName is not the localhost' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Task.ExecuteOnComputerName = 'PC1'

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams
            
        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter $testJobArguments
    }
    It 'with Start-Job when ExecuteOnComputerName is the localhost' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Task.ExecuteOnComputerName = 'localhost'

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams
            
        .$testScript @testParams

        Should -Invoke Start-Job -Times 1 -Exactly -ParameterFilter $testJobArguments
    }
}
Describe 'export the results of the SFTP script' {
    BeforeAll {
        $testInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'to an Excel file' {
        BeforeAll {
            $testExportedExcelRows = $testData | 
            Select-Object Path, UploadedOn, @{
                Name       = 'Action'
                Expression = { $_.Action -join ', ' }
            }, Error

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter "* - $($testInputFile.Tasks[0].Task.Name) - Log.xlsx"

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
                $actualRow.UploadedOn.ToString('yyyyMMdd') | 
                Should -Be $testRow.UploadedOn.ToString('yyyyMMdd')
                $actualRow.Action | Should -Be $testRow.Action
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
}
Describe 'ExportExcelFile.When' {
    Context 'create no Excel file' {
        It "'Never'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'Never'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
        It "'OnlyOnError' and no errors are found" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnError'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrUpload' and there are no errors and no uploads" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                   
                } -AsJob -ComputerName $env:COMPUTERNAME
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnErrorOrUpload'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
                        Path       = 'a'
                        UploadedOn = Get-Date
                        Action     = @()
                        Error      = 'oops'
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnError'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrUpload' and there are uploads but no errors" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path       = 'a'
                        UploadedOn = Get-Date
                        Action     = @('upload')
                        Error      = $null
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnErrorOrUpload'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrUpload' and there are errors but no uploads" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path       = 'a'
                        UploadedOn = Get-Date
                        Action     = @()
                        Error      = 'oops'
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].ExportExcelFile.When = 'OnlyOnErrorOrUpload'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
    }
}
Describe 'SendMail.When' {
    Context 'send an e-mail' {
        It 'with attachment to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.SendMail.To) -and
            ($Bcc -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq '2 items uploaded') -and
            ($Attachments -like '*- Log.xlsx') -and
            ($Message -like "*table*Type*File*Source*Destination*")
            }
        }
    }
    Context 'send no e-mail' {
        It "'Never'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'Never'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Not -Invoke Send-MailHC
        }
        It "'OnlyOnError' and no errors are found" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnError'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Not -Invoke Send-MailHC
        }
        It "'OnlyOnErrorOrUpload' and there are no errors and no uploads" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                   
                } -AsJob -ComputerName $env:COMPUTERNAME
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnErrorOrUpload'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Not -Invoke Send-MailHC
        }
    }
    Context 'send an e-mail' {
        It "'OnlyOnError' and there are errors" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path       = 'a'
                        UploadedOn = Get-Date
                        Action     = @()
                        Error      = 'oops'
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnError'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Invoke Send-MailHC
        }
        It "'OnlyOnErrorOrUpload' and there are uploads but no errors" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path       = 'a'
                        UploadedOn = Get-Date
                        Action     = @('upload')
                        Error      = $null
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnErrorOrUpload'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Invoke Send-MailHC
        }
        It "'OnlyOnErrorOrUpload' and there are errors but no uploads" {
            Mock Start-Job {
                & $realCmdLet.InvokeCommand -Scriptblock { 
                    [PSCustomObject]@{
                        Path       = 'a'
                        UploadedOn = Get-Date
                        Action     = @()
                        Error      = 'oops'
                    }     
                } -AsJob -ComputerName $env:COMPUTERNAME
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyOnErrorOrUpload'
    
            $testNewInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams
    
            .$testScript @testParams
    
            Should -Invoke Send-MailHC
        }
    } -Tag test
}