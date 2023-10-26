#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        Path             = @(
            (New-Item 'TestDrive:/a.txt' -ItemType 'File').FullName
            (New-Item 'TestDrive:/b.txt' -ItemType 'File').FullName
        )
        SftpComputerName = 'PC1'
        SftpPath         = '/out'
        SftpUserName     = 'bob'
        SftpPassword     = 'pass' | ConvertTo-SecureString -AsPlainText -Force
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
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @(
        'Path',
        'SftpComputerName', 
        'SftpUserName', 
        'SftpPassword', 
        'SftpPath'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'generate an error when' {
    It 'authentication to the SFTP server fails' {
        Mock New-SFTPSession {
            throw 'Failed authenticating'
        }

        $testResult = .$testScript @testParams

        $testResult.Error | Should -Be "Failed creating an SFTP session to '$($testParams.SftpComputerName)': Failed authenticating"
    }
    It 'the upload path on the SFTP server does not exist' {
        Mock Test-SFTPPath {
            $false
        }

        $testResult = .$testScript @testParams

        $testResult.Error | Should -Be "Path '$($testParams.SftpPath)' not found on SFTP server"
    }
    It 'Path does not exist and ErrorWhenUploadPathIsNotFound is true' {      
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = 'c:\doesNotExist'
        $testNewParams.ErrorWhenUploadPathIsNotFound = $true

        $testResult = .$testScript @testNewParams

        $testResult.Error | 
        Should -Be "Path '$($testNewParams.Path)' not found"
    }
    It 'the upload fails' {
        Mock Set-SFTPItem {
            throw 'upload failed'
        }

        $testNewParams = $testParams.Clone()
        $testNewParams.Path = $testParams.Path[0]

        $testResult = .$testScript @testNewParams

        $testResult.Error | Should -Be 'upload failed'
    }
}
Describe 'do not start an SFTP sessions when' {
    It 'there is nothing to upload' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = (New-Item 'TestDrive:/f' -ItemType 'Directory').FullName

        .$testScript @testNewParams

        $testNewParams = $testParams.Clone()
        $testNewParams.Path = 'c:\doesNotExist.txt'
        $testNewParams.ErrorWhenUploadPathIsNotFound = $true

        .$testScript @testNewParams

        $testNewParams = $testParams.Clone()
        $testNewParams.Path = 'c:\doesNotExist.txt'
        $testNewParams.ErrorWhenUploadPathIsNotFound = $false

        .$testScript @testNewParams

        Should -Not -Invoke Set-SFTPItem
        Should -Not -Invoke New-SFTPSession
        Should -Not -Invoke Test-SFTPPath
        Should -Not -Invoke Remove-SFTPSession
    }
}
Describe 'upload to the SFTP server' {
    BeforeAll {
        $testFolder = (New-Item 'TestDrive:/a' -ItemType 'Directory').FullName

        $testFiles = @('file1.txt', 'file2.txt', 'file3.txt') | ForEach-Object {
            (New-Item "$testFolder\$_" -ItemType 'File').FullName
        }
    }
    It 'all files in a folder when Path is a folder' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = $testFolder

        .$testScript @testNewParams

        $testFiles | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -eq $_) -and
                ($Destination -eq $testNewParams.SftpPath) -and
                ($SessionId -eq 1)
            }
        }
    }
    It 'all files defined in Path' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = $testFiles

        .$testScript @testNewParams

        $testFiles | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -eq $_) -and
                ($Destination -eq $testNewParams.SftpPath) -and
                ($SessionId -eq 1)
            }
        }
    }
    It 'Return an object with results' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = $testFiles

        $testResults = .$testScript @testNewParams

        $testResults | Should -HaveCount $testFiles.Count

        $testResults | ForEach-Object {
            $_.Path | Should -Not -BeNullOrEmpty
            $_.Uploaded | Should -BeTrue
            $_.DateTime | Should -Not -BeNullOrEmpty
            $_.Action | Should -Be 'file uploaded'
            $_.Error | Should -BeNullOrEmpty
        }
    }
}
Describe 'OverwriteFileOnSftpServer' {
    Context 'when true' {
        BeforeAll {
            $testNewParams = $testParams.Clone()
            $testNewParams.OverwriteFileOnSftpServer = $true

            .$testScript @testNewParams
        }
        It 'the file on the SFTP server is overwritten' {
            $testNewParams.Path | ForEach-Object {
                Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -eq $_) -and
                ($Destination -eq $testNewParams.SftpPath) -and
                ($SessionId -eq 1) -and
                ($Force )
                } -Scope 'Context'
            }
        }
    }
    Context 'when false' {
        BeforeAll {
            $testNewParams = $testParams.Clone()
            $testNewParams.OverwriteFileOnSftpServer = $false

            .$testScript @testNewParams
        }
        It 'the file on the SFTP server is not overwritten' {
            $testNewParams.Path | ForEach-Object {
                Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -eq $_) -and
                ($Destination -eq $testNewParams.SftpPath) -and
                ($SessionId -eq 1) -and
                (-not $Force)
                } -Scope 'Context'
            }
        }
    }
}
Describe 'RemoveFileAfterUpload ' {
    Context 'when false' {
        BeforeAll {
            $testNewParams = $testParams.Clone()
            $testNewParams.RemoveFileAfterUpload = $false
    
            $testResults = .$testScript @testNewParams
        }
        It 'the uploaded file is not removed' {
            $testNewParams.Path | ForEach-Object {
                Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                    ($Path -eq $_) -and
                    ($Destination -eq $testNewParams.SftpPath) -and
                    ($SessionId -eq 1)
                } -Scope 'Context'
    
                $_ | Should -Exist
            }
        }
        It 'return an object with results' {
            $testResults | ForEach-Object {
                $_.Path | Should -Not -BeNullOrEmpty
                $_.Uploaded | Should -BeTrue
                $_.DateTime | Should -Not -BeNullOrEmpty
                $_.Action | Should -Be 'file uploaded'
                $_.Error | Should -BeNullOrEmpty
            }
        }
    }
    Context 'when true' {
        BeforeAll {
            $testNewParams = $testParams.Clone()
            $testNewParams.RemoveFileAfterUpload = $true
    
            $testResults = .$testScript @testNewParams
        }
        It 'the uploaded file is removed' {
            $testNewParams.Path | ForEach-Object {
                Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                    ($Path -eq $_) -and
                    ($Destination -eq $testNewParams.SftpPath) -and
                    ($SessionId -eq 1)
                } -Scope 'Context'
    
                $_ | Should -Not -Exist
            }
        }
        It 'return an object with results' {
            $testResults | ForEach-Object {
                $_.Path | Should -Not -BeNullOrEmpty
                $_.Uploaded | Should -BeTrue
                $_.DateTime | Should -Not -BeNullOrEmpty
                $_.Action[0] | Should -Be 'file uploaded'
                $_.Action[1] | Should -Be 'file removed'
                $_.Error | Should -BeNullOrEmpty
            }
        }
    }
}
