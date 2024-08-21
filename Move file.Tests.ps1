#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        SftpComputerName  = 'PC1'
        SftpUserName      = 'bob'
        Paths             = @(
            @{
                Source      = (New-Item 'TestDrive:/f1' -ItemType 'Directory').FullName
                Destination = 'sftp:/data/'
            }
            @{
                Source      = 'sftp:/report/'
                Destination = (New-Item 'TestDrive:/f2' -ItemType 'Directory').FullName
            }
        )
        MaxConcurrentJobs = 1
        SftpPassword      = 'pass' | ConvertTo-SecureString -AsPlainText -Force
        FileExtensions    = @()
        OverwriteFile     = $false
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
        'Paths',
        'SftpComputerName',
        'SftpUserName',
        'MaxConcurrentJobs'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'Upload to SFTP server' {
    BeforeAll {
        $testSource = @{
            Folder   = (New-Item 'TestDrive:/f3' -ItemType 'Directory').FullName
            FileName = 'a.txt'
        }

        $testSource.FileFullName = (New-Item -Path $testSource.Folder -Name $testSource.FileName -ItemType File).FullName
    }
    Context 'create an object with Error property when' {
        It 'Paths.Source does not exist' {
            $testNewParams = $testParams.Clone()
            $testNewParams.Paths[0].Source = 'c:\doesNotExist'

            $testResult = .$testScript @testNewParams

            $testResult.Error |
            Should -Be "Source folder '$($testNewParams.Paths[0].Source)' not found"
        }
        It 'the SFTP path does not exist' {
            $testNewParams = $testParams.Clone()
            $testNewParams.Paths = @(
                @{
                    Source      = $testSource.Folder
                    Destination = 'sftp:/notExisting/'
                }
            )

            Mock Test-SFTPPath {
                $false
            }

            $testResult = .$testScript @testNewParams

            $testResult.Error |
            Should -Be "Failed upload: Path '/notExisting/' not found on the SFTP server"
        }
        It 'the upload fails' {
            $testNewParams = $testParams.Clone()
            $testNewParams.Paths = @(
                @{
                    Source      = $testSource.Folder
                    Destination = 'sftp:/data/'
                }
            )

            Mock Set-SFTPItem {
                throw 'Oops'
            }

            $error.Clear()

            $testResult = .$testScript @testNewParams

            $testResult.Error | Should -BeLike "*$($testSource.FileName)*Oops"
            $testSource.FileFullName | Should -Exist

            $error | Should -HaveCount 0
        }
    }
    Context 'throw a terminating error when' {
        It 'authentication to the SFTP server fails' {
            $testNewParams = $testParams.Clone()
            $testNewParams.Paths = @(
                @{
                    Source      = $testSource.Folder
                    Destination = 'sftp:/data/'
                }
            )

            Mock New-SFTPSession {
                throw 'Failed authenticating'
            }

            { .$testScript @testNewParams } | Should -Throw "Failed creating an SFTP session to '$($testNewParams.SftpComputerName)': Failed authenticating"
        }
    }
    Context 'do not start an SFTP sessions when' {
        It 'there is nothing to upload' {
            $testNewParams = $testParams.Clone()
            $testNewParams.Paths = @{
                Source      = (New-Item 'TestDrive:/emptyFolder' -ItemType 'Directory').FullName
                Destination = 'sftp:/data/'
            }

            .$testScript @testNewParams

            Should -Not -Invoke Set-SFTPItem
            Should -Not -Invoke New-SFTPSession
            Should -Not -Invoke Test-SFTPPath
            Should -Not -Invoke Remove-SFTPSession
        }
    }
    Context 'when files are found in the source folder' {
        BeforeAll {
            $testNewParams = $testParams.Clone()
            $testNewParams.Paths = @(
                @{
                    Source      = (New-Item 'TestDrive:/z' -ItemType 'Directory').FullName
                    Destination = 'sftp:/data/'
                }
            )

            $testFiles = @('file1.txt', 'file2.txt', 'file3.txt') | ForEach-Object {
                New-Item "$($testNewParams.Paths[0].Source)\$_" -ItemType 'File'
            }

            $testResults = .$testScript @testNewParams
        }
        It 'call Set-SFTPItem to upload a temp file' {
            $testFiles | ForEach-Object {
                Should -Invoke Set-SFTPItem -Times 1 -Exactly -Scope Context -ParameterFilter {
                    ($Path -eq "$($_.FullName).UploadInProgress") -and
                    ($Destination -eq $testNewParams.Paths[0].Destination.TrimStart('sftp:')) -and
                    ($SessionId -eq 1)
                }
            }
        }
        It 'call Rename-SFTPFile to rename the temp file' {
            $testFiles | ForEach-Object {
                Should -Invoke Rename-SFTPFile -Times 1 -Exactly -Scope Context -ParameterFilter {
                    ($Path -eq ($testNewParams.Paths[0].Destination.TrimStart('sftp:') + $_.Name + ".UploadInProgress")) -and
                    ($NewName -eq $_.Name) -and
                    ($SessionId -eq 1)
                }
            }
        }
        It 'remove the local temp file' {
            $testFiles | ForEach-Object {
                $_.FullName | Should -Not -Exist
            }
        }
        It 'Return an object with results' {
            $testResults | Should -HaveCount $testFiles.Count

            foreach ($testFile in $testFiles) {
                $actual = $testResults.where(
                    { $_.FileName -eq $testFile.Name }
                )

                $actual.DateTime | Should -Not -BeNullOrEmpty
                $actual.Source | Should -Be $testNewParams.Paths[0].Source
                $actual.Destination | Should -Be $testNewParams.Paths[0].Destination
                $actual.FileLength | Should -Not -BeNullOrEmpty
                $actual.Action | Should -Be 'File moved'
                $actual.Error | Should -BeNullOrEmpty
            }
        }
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