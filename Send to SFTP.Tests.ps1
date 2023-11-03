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
        SftpPath         = '/out/'
        SftpUserName     = 'bob'
        SftpPassword     = 'pass' | ConvertTo-SecureString -AsPlainText -Force
    }

    Mock Set-SFTPItem
    Mock Rename-SFTPFile
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

        $testResult.Error | Should -Be "General error: Failed creating an SFTP session to '$($testParams.SftpComputerName)': Failed authenticating"
    }
    It 'the upload path on the SFTP server does not exist' {
        Mock Test-SFTPPath {
            $false
        }

        $testResult = .$testScript @testParams

        $testResult.Error | Should -Be "General error: Path '$($testParams.SftpPath)' not found on SFTP server"
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
Describe 'when a file is uploaded it is' {
    BeforeAll {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = (New-Item 'TestDrive:/c.txt' -ItemType 'File').FullName

        $testResults = .$testScript @testNewParams
    }
    It 'renamed to extension .UploadInProgress' {
        'TestDrive:/c.txt' | Should -Not -Exist
        $testResults.Action | Should -Contain 'temp file created'
    }
    It 'uploaded to the SFTP server with extension .UploadInProgress' {
        Should -Invoke Set-SFTPItem -Times 1 -Exactly -Scope 'Describe' -ParameterFilter {
            ($Path -like '*\c.txt.UploadInProgress') -and
            ($Destination -eq $testNewParams.SftpPath) -and
            ($SessionId -eq 1)
        }
        $testResults.Action | Should -Contain 'temp file uploaded'
    }
    It 'renamed on the SFTP server to its original name' {
        Should -Invoke Rename-SFTPFile -Times 1 -Exactly -Scope 'Describe' -ParameterFilter {
            ($NewName -eq 'c.txt') -and
            ($Path -eq ($testNewParams.SftpPath + 'c.txt.UploadInProgress')) -and
            ($SessionId -eq 1)
        }
        $testResults.Action | Should -Contain 'temp file renamed on SFTP server'
    }
    It 'removed after a successful upload' {
        'TestDrive:/c.txt.UploadInProgress' | Should -Not -Exist 
        $testResults.Action | Should -Contain 'temp file removed'
    }
    It 'reports a successful upload' {
        $testResults.Uploaded | Should -BeTrue
        $testResults.Action | Should -Contain 'file successfully uploaded'
    }
}
Describe 'upload to the SFTP server' {
    BeforeAll {
    }
    BeforeEach {
        $testFolder = 'TestDrive:/a'
        Remove-Item $testFolder -Recurse -ErrorAction Ignore
        $null = New-Item $testFolder -ItemType 'Directory' 

        $testFiles = @('file1.txt', 'file2.txt', 'file3.txt') | ForEach-Object {
            New-Item "$testFolder\$_" -ItemType 'File'
        }
    }
    It 'all files in a folder when Path is a folder' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = $testFolder

        .$testScript @testNewParams

        $testFiles | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -like "*\$($_.Name).UploadInProgress") -and
                ($Destination -eq $testNewParams.SftpPath) -and
                ($SessionId -eq 1)
            }
        }
    }
    It 'all files defined in Path' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = $testFiles.FullName

        .$testScript @testNewParams

        $testFiles | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -like "*\$($_.Name).UploadInProgress") -and
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
            $_.LocalPath | Should -Not -BeNullOrEmpty
            $_.SftpPath | Should -Be $testNewParams.SftpPath
            $_.FileName | Should -Not -BeNullOrEmpty
            $_.Uploaded | Should -BeTrue
            $_.DateTime | Should -Not -BeNullOrEmpty
            $_.Action | Should -Not -BeNullOrEmpty
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
                ($Path -like '*.UploadInProgress') -and
                ($Destination -eq $testNewParams.SftpPath) -and
                ($SessionId -eq 1) -and
                ($Force)
                } -Scope 'Context'
            }
        }
    }  -Tag test
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
