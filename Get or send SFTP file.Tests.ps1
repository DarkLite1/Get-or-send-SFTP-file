#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $testInputFile = @{
        MaxConcurrentJobs = 1
        Tasks             = @(
            @{
                TaskName = 'App x'
                Sftp     = @{
                    ComputerName = 'sftp.server.com'
                    Credential   = @{
                        UserName        = 'envVarBob'
                        Password        = 'envVarPasswordBob'
                        PasswordKeyFile = $null
                    }
                }
                Actions  = @(
                    @{
                        Type      = 'Upload'
                        Parameter = @{
                            SftpPath             = '/SFTP/folder/'
                            ComputerName         = 'PC1'
                            Path                 = (New-Item 'TestDrive:\a' -ItemType Directory).FullName
                            FileExtension        = $null
                            PartialFileExtension = '.UploadInProgress'
                            Option               = @{
                                OverwriteFile            = $false
                                RemoveFailedPartialFiles = $false
                                ErrorWhen                = @{
                                    PathIsNotFound = $true
                                }
                            }
                        }
                    }
                    @{
                        Type      = 'Download'
                        Parameter = @{
                            SftpPath             = '/SFTP/folder/'
                            ComputerName         = 'PC2'
                            Path                 = (New-Item 'TestDrive:\d' -ItemType Directory).FullName
                            FileExtensions       = @('.txt')
                            PartialFileExtension = '.DownloadInProgress'
                            Option               = @{
                                RemoveFailedPartialFiles = $false
                                OverwriteFile            = $false
                            }
                        }
                    }
                )
            }
        )
        SendMail          = @{
            To   = 'bob@contoso.com'
            When = 'Always'
        }
        ExportExcelFile   = @{
            When = 'OnlyOnErrorOrAction'
        }
    }

    $testData = @(
        [PSCustomObject]@{
            LocalPath  = $testInputFile.Tasks[0].Actions[0].Parameter.Path
            SftpPath   = $testInputFile.Tasks[0].Actions[0].Parameter.SftpPath
            FileName   = $testInputFile.Tasks[0].Actions[0].Parameter.Path | Split-Path -Leaf
            FileLength = 5KB
            DateTime   = Get-Date
            Uploaded   = $true
            Action     = @('file uploaded', 'file removed')
            Error      = $null
        }
        [PSCustomObject]@{
            LocalPath  = $testInputFile.Tasks[0].Actions[1].Parameter.Path
            SftpPath   = $testInputFile.Tasks[0].Actions[1].Parameter.SftpPath
            FileName   = 'sftp file.txt'
            FileLength = 3KB
            DateTime   = Get-Date
            Downloaded = $true
            Action     = @('file downloaded', 'file removed')
            Error      = $null
        }
    )

    $testExportedExcelRows = @(
        [PSCustomObject]@{
            SftpServer   = $testInputFile.Tasks[0].Sftp.ComputerName
            ComputerName = $testInputFile.Tasks[0].Actions[0].Parameter.ComputerName
            Source       = $testData[0].LocalPath
            Destination  = $testData[0].SftpPath
            FileName     = $testData[0].FileName
            FileSize     = $testData[0].FileLength / 1KB
            DateTime     = $testData[0].DateTime
            Type         = 'Upload'
            Successful   = $true
            Action       = $testData[0].Action -join ', '
            Error        = $null
        }
        [PSCustomObject]@{
            SftpServer   = $testInputFile.Tasks[0].Sftp.ComputerName
            ComputerName = $testInputFile.Tasks[0].Actions[1].Parameter.ComputerName
            Source       = $testData[1].SftpPath
            Destination  = $testData[1].LocalPath
            FileName     = $testData[1].FileName
            FileSize     = $testData[1].FileLength / 1KB
            DateTime     = $testData[1].DateTime
            Type         = 'Download'
            Successful   = $true
            Action       = $testData[1].Action -join ', '
            Error        = $null
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
    Mock Invoke-Command {
        $testData[0]
    } -ParameterFilter {
        $FilePath -eq $testParams.Path.UploadScript
    }
    Mock Invoke-Command {
        $testData[1]
    } -ParameterFilter {
        $FilePath -eq $testParams.Path.DownloadScript
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
                'MaxConcurrentJobs', 'Tasks', 'SendMail', 'ExportExcelFile'
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
                'TaskName', 'Sftp', 'Actions'
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
                'UserName'
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
            Context 'Tasks.Sftp.Credential' {
                It 'Password or PasswordKeyFile are missing' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Sftp.Credential.Password = $null
                    $testNewInputFile.Tasks[0].Sftp.Credential.PasswordKeyFile = $null

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and
                            ($Message -like "*$ImportFile*Property 'Tasks.Sftp.Credential.Password' or 'Tasks.Sftp.Credential.PasswordKeyFile' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Password and PasswordKeyFile used at the same time' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Sftp.Credential.Password = 'a'
                    $testNewInputFile.Tasks[0].Sftp.Credential.PasswordKeyFile = 'b'

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and
                            ($Message -like "*$ImportFile*Property 'Tasks.Sftp.Credential.Password' and 'Tasks.Sftp.Credential.PasswordKeyFile' cannot be used at the same time*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'PasswordKeyFile does not exist' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Sftp.Credential.Password = $null
                    $testNewInputFile.Tasks[0].Sftp.Credential.PasswordKeyFile = 'a'

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and
                            ($Message -like "*$ImportFile*Failed converting the task.Sftp.Credential.PasswordKeyFile*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'PasswordKeyFile is an empty file' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Sftp.Credential.Password = $null
                    $testNewInputFile.Tasks[0].Sftp.Credential.PasswordKeyFile = (New-Item 'TestDrive:\a.pub' -ItemType File).FullName

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and
                            ($Message -like "*$ImportFile*Failed converting the task.Sftp.Credential.PasswordKeyFile*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
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
                    'SftpPath', 'Path', 'Option',
                    'PartialFileExtension'
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
                    'RemoveFailedPartialFiles'
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
            }
            Context "Tasks.Actions.Type is 'Upload'" {
                It 'Tasks.Actions.Parameter.<_> not found' -ForEach @(
                    'SftpPath', 'Path', 'Option', 'PartialFileExtension'
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
                    'RemoveFailedPartialFiles'
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
                    'PathIsNotFound'
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
            It 'SendMail.<_> not found' -ForEach @(
                'To', 'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'ExportExcelFile.<_> not found' -ForEach @(
                'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.ExportExcelFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'ExportExcelFile.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'ExportExcelFile.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.ExportExcelFile.When = 'wrong'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'ExportExcelFile.When' with value 'wrong' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.When = 'wrong'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.When' with value 'wrong' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
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
            It 'Tasks.Actions.Parameter.PartialFileExtension does not start with a dot' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Actions[0].Parameter.PartialFileExtension = 'txt'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'Tasks.Actions.Parameter.PartialFileExtension' needs to start with a dot. For example: '.txt', '.xml'*")
                }
            }
            It 'Tasks.Actions.Parameter.FileExtension does not start with a dot' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Actions[0].Parameter.FileExtensions = @('txt', '.xml')

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'Tasks.Actions.Parameter.FileExtensions' needs to start with a dot. For example: '.txt', '.xml'*")
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
Describe 'correct the import file' {
    It 'add trailing slashes to SFTP path when they are not there' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Actions[0].Parameter.SftpPath = '/a'
        $testNewInputFile.Tasks[0].Actions[1].Parameter.SftpPath = '\b/'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        $Tasks[0].Actions[0].Parameter.SftpPath | Should -Be '/a/'
        $Tasks[0].Actions[1].Parameter.SftpPath | Should -Be '/b/'
    }
}
Describe 'execute the SFTP script' {
    BeforeAll {
        $testJobArguments = @(
            {
                ($FilePath -eq $testParams.Path.UploadScript) -and
                ($ArgumentList[0] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Path) -and
                ($ArgumentList[1] -eq $testInputFile.Tasks[0].Sftp.ComputerName) -and
                ($ArgumentList[2] -eq $testInputFile.Tasks[0].Actions[0].Parameter.SftpPath) -and
                ($ArgumentList[3] -eq 'bobUserName') -and
                ($ArgumentList[4] -eq $testInputFile.Tasks[0].Actions[0].Parameter.PartialFileExtension) -and
                ($ArgumentList[5] -eq 'bobPasswordEncrypted') -and
                (-not $ArgumentList[6]) -and
                ($ArgumentList[7] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Option.OverwriteFile) -and
                ($ArgumentList[8] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Option.ErrorWhen.PathIsNotFound) -and
                ($ArgumentList[9] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Option.RemoveFailedPartialFiles) -and
                ($ArgumentList[10] -eq $testInputFile.Tasks[0].Actions[0].Parameter.FileExtensions)
            }
            {
                ($FilePath -eq $testParams.Path.DownloadScript) -and
                ($ArgumentList[0] -eq $testInputFile.Tasks[0].Actions[1].Parameter.Path) -and
                ($ArgumentList[1] -eq $testInputFile.Tasks[0].Sftp.ComputerName) -and
                ($ArgumentList[2] -eq $testInputFile.Tasks[0].Actions[1].Parameter.SftpPath) -and
                ($ArgumentList[3] -eq 'bobUserName') -and
                ($ArgumentList[4] -eq $testInputFile.Tasks[0].Actions[1].Parameter.PartialFileExtension) -and
                ($ArgumentList[5] -eq 'bobPasswordEncrypted') -and
                (-not $ArgumentList[6]) -and
                ($ArgumentList[7] -eq $testInputFile.Tasks[0].Actions[1].Parameter.FileExtensions) -and
                ($ArgumentList[8] -eq $testInputFile.Tasks[0].Actions[1].Parameter.Option.OverwriteFile) -and
                ($ArgumentList[9] -eq $testInputFile.Tasks[0].Actions[1].Parameter.Option.RemoveFailedPartialFiles)
            }
        )
    }
    Context "for Tasks.Actions.Type 'Upload'" {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].Actions = @(
                $testNewInputFile.Tasks[0].Actions[0]
            )
        }
        It 'with Invoke-Command when Tasks.Actions.Parameter.ComputerName is not the localhost' {
            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
                (& $testJobArguments[0]) -and
                ($ComputerName -eq $testNewInputFile.Tasks[0].Actions[0].Parameter.ComputerName)
            }
        }
        It 'with a scriptblock when Tasks.Actions.Parameter.ComputerName is the localhost' {
            $testNewInputFile.Tasks[0].Actions[0].Parameter.ComputerName = 'localhost'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Invoke-Command
        }
    }
    Context "for Tasks.Actions.Type 'Download'" {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].Actions = @(
                $testNewInputFile.Tasks[0].Actions[1]
            )
        }
        It 'with Invoke-Command when Tasks.Actions.Parameter.ComputerName is not the localhost' {
            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
                (& $testJobArguments[1]) -and
                ($ComputerName -eq $testNewInputFile.Tasks[0].Actions[0].Parameter.ComputerName)
            }
        }
        It 'with a scriptblock when Tasks.Actions.Parameter.ComputerName is the localhost' {
            $testNewInputFile.Tasks[0].Actions[0].Parameter.ComputerName = 'localhost'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Invoke-Command
        }
    }
    It 'with Tasks.Sftp.Credential.PasswordKeyFile and a blank secure string for Tasks.Sftp.Credential.Password' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Actions[0].Parameter.ComputerName = 'PC1'
        $testNewInputFile.Tasks[0].Sftp.Credential.Password = $null
        $testNewInputFile.Tasks[0].Sftp.Credential.PasswordKeyFile = (New-Item 'TestDrive:\key' -ItemType File).FullName

        'passKeyContent' | Out-File -LiteralPath $testNewInputFile.Tasks[0].Sftp.Credential.PasswordKeyFile

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
            ($FilePath -eq $testParams.Path.UploadScript) -and
            ($ArgumentList[0] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Path) -and
            ($ArgumentList[1] -eq $testInputFile.Tasks[0].Sftp.ComputerName) -and
            ($ArgumentList[2] -eq $testInputFile.Tasks[0].Actions[0].Parameter.SftpPath) -and
            ($ArgumentList[3] -eq 'bobUserName') -and
            ($ArgumentList[4] -eq $testInputFile.Tasks[0].Actions[0].Parameter.PartialFileExtension) -and
            ($ArgumentList[5] -is 'SecureString') -and
            ($ArgumentList[6] -eq 'passKeyContent') -and
            ($ArgumentList[7] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Option.OverwriteFile) -and
            ($ArgumentList[8] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Option.ErrorWhen.PathIsNotFound) -and
            ($ArgumentList[9] -eq $testInputFile.Tasks[0].Actions[0].Parameter.Option.RemoveFailedPartialFiles) -and
            ($ArgumentList[10] -eq $testInputFile.Tasks[0].Actions[0].Parameter.FileExtensions)
        }
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
                    $_.Source -eq $testRow.Source
                }
                $actualRow.SftpServer | Should -Be $testRow.SftpServer
                $actualRow.ComputerName | Should -Be $testRow.ComputerName
                $actualRow.Destination | Should -Be $testRow.Destination
                $actualRow.DateTime.ToString('yyyyMMdd') |
                Should -Be $testRow.DateTime.ToString('yyyyMMdd')
                $actualRow.Action | Should -Be $testRow.Action
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Type | Should -Be $testRow.Type
                $actualRow.Successful | Should -Be $testRow.Successful
                $actualRow.FileName | Should -Be $testRow.FileName
                $actualRow.FileSize | Should -Be $testRow.FileSize
            }
        } -Tag test
    }
    Context 'send an e-mail' {
        It 'with attachment to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.SendMail.To) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq '1 uploaded, 1 downloaded') -and
            ($Attachments -like '*- Log.xlsx') -and
            ($Message -like "*table*$($testInputFile.Tasks[0].TaskName)*SFTP Server*$($testInputFile.Tasks[0].Sftp.ComputerName)*ComputerName*Source*Destination*Result*$($testInputFile.Tasks[0].Actions[0].ComputerName)*$($testInputFile.Tasks[0].Actions[0].Parameter.Path)*$($testInputFile.Tasks[0].Actions[0].Parameter.SftpPath)*1 uploaded*$($testInputFile.Tasks[0].Actions[1].ComputerName)*$($testInputFile.Tasks[0].Actions[1].Parameter.SftpPath)*$($testInputFile.Tasks[0].Actions[1].Parameter.Path)*1 downloaded*")
            }
        }
    }
}
Describe 'ExportExcelFile.When' {
    Context 'create no Excel file' {
        It "'Never'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'Never'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
        It "'OnlyOnError' and no errors are found" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are no errors and no actions" {
            Mock Invoke-Command {
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
    }
    Context 'create an Excel file' {
        It "'OnlyOnError' and there are errors" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Action   = @()
                    Error    = 'oops'
                }
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are actions but no errors" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Uploaded = $true
                    Action   = @('upload')
                    Error    = $null
                }
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are errors but no actions" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    Uploaded = $false
                    DateTime = Get-Date
                    Action   = @()
                    Error    = 'oops'
                }
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnErrorOrAction'

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
            ParameterFilter = { $To -eq $testNewInputFile.SendMail.To }
        }
    }
    Context 'send no e-mail to the user' {
        It "'Never'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'Never'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnError' and no errors are found" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC
        }
        It "'OnlyOnErrorOrAction' and there are no errors and no actions" {
            Mock Invoke-Command {
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC
        }
    }
    Context 'send an e-mail to the user' {
        It "'OnlyOnError' and there are errors" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Action   = @()
                    Error    = 'oops'
                }
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are actions but no errors" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Uploaded = $true
                    Action   = @('upload')
                    Error    = $null
                }
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are errors but no actions" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Action   = @()
                    Error    = 'oops'
                }
            } -ParameterFilter {
                ($FilePath -eq $testParams.Path.DownloadScript) -or
                ($FilePath -eq $testParams.Path.UploadScript)
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
    }
}