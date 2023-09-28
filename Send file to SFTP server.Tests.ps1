#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $testFile = @(
        New-Item 'TestDrive:/file1.txt' -ItemType File
        New-Item 'TestDrive:/file2.txt' -ItemType File
    )

    $testInputFile = @{
        SendMail        = @{
            To   = 'bob@contoso.com'
            When = 'Always'
        }
        ExportExcelFile = @{
            When = 'Always'
        }
        Upload          = @(
            @{
                Type        = 'File'
                Source      = @(
                    $testFile[0].FullName
                    $testFile[1].FullName
                )
                Destination = '/SFTP/folder/'
                Option      = @{
                    OverwriteDestinationData = $false
                    RemoveSourceAfterUpload  = $false
                    ErrorWhen                = @{
                        SourceIsNotFound = $false
                    }
                }
            }
        )
        Sftp            = @{
            ComputerName = 'PC1'
            Credential   = @{
                UserName = 'envVarBob'
                Password = 'envVarPasswordBob'
            }
        }
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@conotoso.com'
    }

    Function Get-EnvironmentVariableValueHC {
        Param(
            [String]$Name
        )
    }
    
    Mock Get-EnvironmentVariableValueHC {
        'bob'
    } -ParameterFilter {
        $Name -eq $testInputFile.SFtp.Credential.UserName
    }
    Mock Get-EnvironmentVariableValueHC {
        'PasswordBob'
    } -ParameterFilter {
        $Name -eq $testInputFile.SFtp.Credential.Password
    }
    Mock Set-SFTPItem
    Mock New-SFTPSession {
        [PSCustomObject]@{
            SessionID = 1
        }
    }
    Mock Test-SFTPPath {
        $true
    }
    Mock Remove-SFTPSession
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
                'SendMail', 'Upload', 'ExportExcelFile'
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
            It 'ExportExcelFile.<_> not found' -ForEach @(
                'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.ExportExcelFile.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'ExportExcelFile.When' with value 'wrong' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrUpload'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail.<_> not found' -ForEach @(
                'To', 'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
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
            It 'SendMail.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.When = 'wrong'
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'SendMail.When' with value 'wrong' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrUpload'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Upload.<_> not found' -ForEach @(
                'Type', 'Source', 'Destination', 'Option'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Upload[0].$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Upload.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Upload.Option.<_> is not a boolean' -ForEach @(
                'OverwriteDestinationData', 
                'RemoveSourceAfterUpload'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Upload[0].Option.$_ = 2
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Upload.Option.$_' is not a boolean value*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Upload.Option.ErrorWhen.<_> is not a boolean' -ForEach @(
                'SourceIsNotFound'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Upload[0].Option.ErrorWhen.$_ = 2
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Upload.Option.ErrorWhen.$_' is not a boolean value*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Upload.Option.ErrorWhen.<_> is not a boolean' -ForEach @(
                'SourceFolderIsEmpty'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Upload[0].Option.ErrorWhen.$_ = 2
                $testNewInputFile.Upload[0].Type = 'FolderContent'
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams

                .$testScript @testParams

                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Upload[0].Option.ErrorWhen.$_ = 2
                $testNewInputFile.Upload[0].Type = 'Folder'
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 2 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Upload.Option.ErrorWhen.$_' is not a boolean value*")
                }
                Should -Invoke Write-EventLog -Exactly 2 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Sftp.<_> not found' -ForEach @(
                'ComputerName'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Sftp.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Sftp.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Sftp.Credential.<_> not found' -ForEach @(
                'UserName', 'Password'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Sftp.Credential.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'Sftp.Credential.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
    It 'authentication to the SFTP server fails' {
        Mock New-SFTPSession {
            throw 'Failed authenticating'
        }

        $testInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Failed creating an SFTP session to '$($testInputFile.sftp.ComputerName)'*")
        }
    }
}
Describe 'send an error e-mail to the user when' {
    BeforeAll {
        $MailUserParams = {
            ($To -eq $testInputFile.SendMail.To) -and 
            ($Bcc -eq $testParams.ScriptAdmin) -and 
            ($Priority -eq 'High') -and 
            ($Subject -like '*error*')
        }    
    }
    It 'the SFTP upload destination folder does not exist' {
        Mock Test-SFTPPath {
            $false
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Upload[0].Destination = '/notExisting'

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Times 1 -Exactly -ParameterFilter {
            (&$MailUserParams) -and 
            ($Message -like "*Upload destination folder '/notExisting' not found on SFTP server*")
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        $testInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams
    }
    It 'upload each source to the SFTP server' {
        @(
            $testInputFile.Upload[0].Source[0]
            $testInputFile.Upload[0].Source[1]
            
        ) | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -Scope 'Describe' -ParameterFilter {
                $Path -eq $_
            }
        }
    }
    It 'close the SFTP session' {
        Should -Invoke Remove-SFTPSession -Times 1 -Exactly -Scope Describe
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                [PSCustomObject]@{
                    Type          = $testInputFile.Upload[0].Type
                    Source        = $testFile[0].FullName
                    Destination   = $testInputFile.Upload[0].Destination
                    UploadedItems = 1
                    UploadedOn    = Get-Date
                    Info          = ''
                    Error         = $null
                }
                [PSCustomObject]@{
                    Type          = $testInputFile.Upload[0].Type
                    Source        = $testFile[1].FullName
                    Destination   = $testInputFile.Upload[0].Destination
                    UploadedItems = 1
                    UploadedOn    = Get-Date
                    Info          = ''
                    Error         = $null
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
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
                $actualRow.UploadedOn.ToString('yyyyMMdd') | 
                Should -Be $testRow.UploadedOn.ToString('yyyyMMdd')
                $actualRow.Type | Should -Be $testRow.Type
                $actualRow.UploadedItems | Should -Be $testRow.UploadedItems
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Info | Should -Be $testRow.Info
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
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
}
Describe 'when OverwriteDestinationData is' {
    It 'true the file on the SFTP server is overwritten' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Upload[0].Option.OverwriteDestinationData = $true

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        $testNewInputFile.Upload[0].Source | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -eq $_) -and
                ($Force)
            }
        }
    }
    It 'false the file on the SFTP server is not overwritten' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Upload[0].Option.OverwriteDestinationData = $false

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        $testNewInputFile.Upload[0].Source | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -eq $_) -and
                (-not $Force)
            }
        }
    }
}
Describe 'when RemoveSourceAfterUpload is' {
    It 'true the uploaded source is removed' {
        $testSourceFiles = @(
            (New-Item 'TestDrive:/file3.txt' -ItemType File).FullName
            (New-Item 'TestDrive:/file4.txt' -ItemType File).FullName
        )

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Upload[0].Source = $testSourceFiles
        $testNewInputFile.Upload[0].Option.RemoveSourceAfterUpload = $true

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        $testSourceFiles | Should -Not -Exist
    }
    It 'false the uploaded source is not removed' {
        $testSourceFiles = @(
            (New-Item 'TestDrive:/file3.txt' -ItemType File).FullName
            (New-Item 'TestDrive:/file4.txt' -ItemType File).FullName
        )

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Upload[0].Source = $testSourceFiles
        $testNewInputFile.Upload[0].Option.RemoveSourceAfterUpload = $false

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        $testSourceFiles | Should -Exist
    }
}