#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $testData = @{
        Folder = (New-Item 'TestDrive:/folder' -ItemType 'Directory').FullName
        File   = @(
            (New-Item 'TestDrive:/folder/a.txt' -ItemType 'File').FullName
            (New-Item 'TestDrive:/folder/b.txt' -ItemType 'File').FullName
        )
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        Path                 = $testData.Folder
        SftpComputerName     = 'PC1'
        SftpPath             = '/out/'
        FileExtensions       = @()
        PartialFileExtension = '.UploadInProgress'
        SftpUserName         = 'bob'
        SftpPassword         = 'pass' | ConvertTo-SecureString -AsPlainText -Force
    }

    Mock Get-SFTPChildItem
    Mock Set-SFTPItem
    Mock Rename-SFTPFile
    Mock Remove-SFTPItem
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
        'SftpPath',
        'PartialFileExtension'
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
    It 'Path does not exist' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = 'c:\doesNotExist'

        $testResult = .$testScript @testNewParams

        $testResult.Error |
        Should -Be "Path '$($testNewParams.Path)' not found"
    }
    It 'the upload fails' {
        Mock Set-SFTPItem {
            throw 'Oops'
        }

        $Error.Clear()

        $testResult = .$testScript @testParams

        $testResult[0].Error | Should -BeLike "*$($testData.File[0])*Oops"
        $testResult[1].Error | Should -BeLike "*$($testData.File[1])*Oops"

        $error | Should -HaveCount 0
    }
}
Describe 'do not start an SFTP sessions when' {
    It 'there is nothing to upload' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = (New-Item 'TestDrive:/emptyFolder' -ItemType 'Directory').FullName

        .$testScript @testNewParams

        Should -Not -Invoke Set-SFTPItem
        Should -Not -Invoke New-SFTPSession
        Should -Not -Invoke Test-SFTPPath
        Should -Not -Invoke Remove-SFTPSession
    }
}
Describe 'when a file is uploaded' {
    BeforeAll {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = (New-Item 'TestDrive:/c' -ItemType 'Directory').FullName

        $testFile = (New-Item "$($testNewParams.Path)\a.txt" -ItemType 'File')

        $testResults = .$testScript @testNewParams
    }
    It 'it is renamed to extension .UploadInProgress' {
        $testFile.FullName | Should -Not -Exist
    }
    It 'it is uploaded to the SFTP server with extension .UploadInProgress' {
        Should -Invoke Set-SFTPItem -Times 1 -Exactly -Scope 'Describe' -ParameterFilter {
            ($Path -eq "$($testFile.FullName).UploadInProgress") -and
            ($Destination -eq $testNewParams.SftpPath) -and
            ($SessionId -eq 1)
        }
    }
    It 'it is renamed on the SFTP server to its original name' {
        Should -Invoke Rename-SFTPFile -Times 1 -Exactly -Scope 'Describe' -ParameterFilter {
            ($NewName -eq $testFile.Name) -and
            ($Path -eq ($testNewParams.SftpPath + $testFile.Name + '.UploadInProgress')) -and
            ($SessionId -eq 1)
        }
    }
    It 'it is removed after a successful upload' {
        "$($testFile.FullName).UploadInProgress" | Should -Not -Exist
        $testFile.FullName | Should -Not -Exist
    }
    Context 'an object is returned with property' {
        It 'DateTime' {
            $testResults.DateTime.ToString('yyyyMMdd') |
            Should -Be (Get-Date).ToString('yyyyMMdd')
        }
        Context 'Action' {
            It "<_>" -ForEach @(
                'temp file created',
                'temp file uploaded',
                'temp file removed',
                'temp file renamed on SFTP server',
                'file successfully uploaded'
            ) {
                $testResults.Action | Should -Contain $_
            }
        }
        It 'Uploaded' {
            $testResults.Uploaded | Should -BeTrue
        }
        It 'LocalPath' {
            $testResults.LocalPath | Should -Be $testNewParams.Path
        }
        It 'FileName' {
            $testResults.FileName | Should -Be $testFile.Name
        }
        It 'SftpPath' {
            $testResults.SftpPath | Should -Be $testNewParams.SftpPath
        }
        It 'Error' {
            $testResults.Error | Should -BeNullOrEmpty
        }
        It 'FileLength' {
            $testResults.FileLength | Should -Not -BeNullOrEmpty
            $testResults.FileLength | Should -BeOfType [long]
        }
    }
}
Describe 'upload to the SFTP server' {
    BeforeAll {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = (New-Item 'TestDrive:/z' -ItemType 'Directory').FullName

        $testFiles = @('file1.txt', 'file2.txt', 'file3.txt') | ForEach-Object {
            New-Item "$($testNewParams.Path)\$_" -ItemType 'File'
        }

        $testResults = .$testScript @testNewParams
    }
    It 'all files in the Path folder' {
        $testFiles | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($Path -eq "$($_.FullName).UploadInProgress") -and
                ($Destination -eq $testNewParams.SftpPath) -and
                ($SessionId -eq 1)
            }
        }
    }
    It 'Return an object with results' {
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
Describe 'OverwriteFile' {
    BeforeAll {
        $testSFtpFile = @{
            Name     = 'a.txt'
            FullName = 'sftpPath\a.txt'
        }

        Mock Get-SFTPChildItem {
            $testSFtpFile
        }
    }
    It 'when true the file on the SFTP server is overwritten' {
        $testNewParams = $testParams.Clone()
        $testNewParams.OverwriteFile = $true
        $testNewParams.Path = (New-Item 'TestDrive:/y' -ItemType 'Directory').FullName

        $testFile = New-Item "$($testNewParams.Path)\$($testSFtpFile.Name)" -ItemType 'File'

        .$testScript @testNewParams

        Should -Invoke Remove-SFTPItem -Times 1 -Exactly -ParameterFilter {
            ($Path -eq $testSFtpFile.FullName) -and
            ($SessionId -eq 1)
        }
    }
    It 'when false the file on the SFTP server is not overwritten' {
        $testNewParams = $testParams.Clone()
        $testNewParams.OverwriteFile = $false

        .$testScript @testNewParams

        Should -Not -Invoke Remove-SFTPItem
    }
}
Describe 'when RemoveFailedPartialFiles is true' {
    Context 'remove partial files that are not completely uploaded' {
        It 'from the folder in Path' {
            $testNewParams = $testParams.Clone()
            $testNewParams.RemoveFailedPartialFiles = $true
            $testNewParams.Path = (New-Item 'TestDrive:\k' -ItemType 'Directory').FullName

            $testFiles = @(
                Join-Path $testNewParams.Path "file.txt"
                Join-Path $testNewParams.Path "file.txt$($testParams.PartialFileExtension)"
            ) | ForEach-Object {
                New-Item -Path $_ -ItemType 'File'
            }

            $testResults = .$testScript @testNewParams

            $testFiles[1].FullName | Should -Not -Exist

            $testResult = $testResults | Where-Object {
                $_.FileName -eq $testFiles[1].Name
            }

            $testResult.Action | Should -Be "removed failed uploaded partial file '$($testFiles[1].FullName)'"
        }
        It 'from the SFTP server' {
            $testSFtpFile = @{
                Name     = "c.txt$($testParams.PartialFileExtension)"
                FullName = "sftpPath\c.txt$($testParams.PartialFileExtension)"
            }

            Mock Get-SFTPChildItem {
                $testSFtpFile
            }

            $testNewParams = $testParams.Clone()
            $testNewParams.RemoveFailedPartialFiles = $true

            $testNewParams.Path = (New-Item 'TestDrive:/p' -ItemType 'Directory').FullName

            $testFile = New-Item "$($testNewParams.Path)\b.txt" -ItemType 'File'

            $testResults = .$testScript @testNewParams

            $testResult = $testResults | Where-Object {
                $_.FileName -eq $testSFtpFile.Name
            }

            $testResult.Action |
            Should -Be "removed partial file '$($testSFtpFile.FullName)'"
        }
    }
}
Describe 'when FileExtensions is' {
    BeforeAll {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = (New-Item 'TestDrive:\Upload' -ItemType 'Directory').FullName
    }
    It 'empty, all files are uploaded' {
        $testNewParams.FileExtensions = @()

        $testFiles = @(
            'file.txt'
            'file.xml'
            'file.jpg'
        ) | ForEach-Object {
            New-Item -Path (Join-Path $testNewParams.Path $_) -ItemType 'File'
        }

        .$testScript @testNewParams

        foreach ($testFile in $testFiles) {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Destination -eq $testNewParams.SftpPath) -and
                ($Path -eq ($testFile.FullName + $testNewParams.PartialFileExtension))
            }
        }
    }
    It 'not empty, only specific files are uploaded' {
        $testNewParams.FileExtensions = @('.txt', '.xml')

        $testFiles = @(
            'file.txt'
            'file.xml'
            'file.jpg'
        ) | ForEach-Object {
            New-Item -Path (Join-Path $testNewParams.Path $_) -ItemType 'File'
        }

        .$testScript @testNewParams

        foreach ($testFile in $testFiles) {
            if ($testFile.Extension -eq '.jpg') {
                Should -Not -Invoke Set-SFTPItem -ParameterFilter {
                    ($Destination -eq $testNewParams.SftpPath) -and
                    ($Path -eq ($testFile.FullName + $testNewParams.PartialFileExtension))
                }
                Continue
            }

            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Destination -eq $testNewParams.SftpPath) -and
                ($Path -eq ($testFile.FullName + $testNewParams.PartialFileExtension))
            }
        }
    }
}
Describe 'when SftpOpenSshKeyFile is used' {
    It 'New-SFTPSession is called with the correct arguments' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = (New-Item 'TestDrive:\Upload' -ItemType 'Directory').FullName
        New-Item -Path (Join-Path $testNewParams.Path 'k.txt') -ItemType 'File'

        $testNewParams.SftpOpenSshKeyFile = @('a')

        .$testScript @testNewParams

        Should -Invoke New-SFTPSession -Times 1 -Exactly -ParameterFilter {
            ($ComputerName -eq $testNewParams.SftpComputerName) -and
            ($Credential -is 'System.Management.Automation.PSCredential') -and
            ($AcceptKey) -and
            ($KeyString -eq 'a')
        }
    }
}