#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML, Toolbox.General
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        Path                 = (New-Item 'TestDrive:/Download' -ItemType 'Directory').FullName
        SftpComputerName     = 'PC1'
        SftpPath             = '/out/'
        FileExtensions       = @()
        PartialFileExtension = '.DownloadInProgress'
        SftpUserName         = 'bob'
        SftpPassword         = 'pass' | ConvertTo-SecureString -AsPlainText -Force
    }

    $testData = @(
        [PSCustomObject]@{
            Name     = 'file.txt'
            FullName = $testParams.SftpPath + 'file.txt'
            Length   = 1KB
        }
    )

    Mock Get-SFTPChildItem {
        $testData
    }
    Mock Get-SFTPItem
    Mock Rename-SFTPFile
    Mock Rename-Item
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

        $testResult.Error | Should -Be "General error: Failed creating an SFTP session to '$($testParams.SftpComputerName)': Failed authenticating"
    }
    It 'the download path on the SFTP server does not exist' {
        Mock Test-SFTPPath {
            $false
        }

        $testResult = .$testScript @testParams

        $testResult.Error | Should -Be "General error: Path '$($testParams.SftpPath)' not found on SFTP server"
    }
    It 'Path does not exist' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = 'c:\doesNotExist'

        $testResult = .$testScript @testNewParams

        $testResult.Error |
        Should -BeLike "*Path '$($testNewParams.Path)' not found"
    }
    It 'the download fails' {
        Mock Get-SFTPItem {
            throw 'download failed'
        }

        $testResult = .$testScript @testParams

        $testResult.Error | Should -BeLike '*download failed'
    }
    It 'the file cannot be removed from the SFTP server' {
        Mock Remove-SFTPItem {
            throw 'removal failed'
        }

        $testResult = .$testScript @testParams

        $testResult.Error | Should -BeLike "*removal failed"
    }
    It 'the file cannot be renamed on the file system' {
        Mock Rename-Item {
            throw 'rename failed'
        }

        $testResult = .$testScript @testParams

        $testResult.Error | Should -BeLike "*rename failed"
    }
}
Describe 'when a file is downloaded' {
    BeforeAll {
        $testResults = .$testScript @testParams
    }
    It 'it is renamed on the SFTP server to extension .DownloadInProgress' {
        Should -Invoke Rename-SFTPFile -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Path -eq $testParams.SftpPath + $testData[0].Name) -and
            ($NewName -eq $testData[0].Name + $testParams.PartialFileExtension)
        }
    }
    It 'it is downloaded from the SFTP server with extension .DownloadInProgress' {
        Should -Invoke Get-SFTPItem -Times 1 -Exactly -Scope 'Describe' -ParameterFilter {
            ($Path -eq $testParams.SftpPath + $testData[0].Name + $testParams.PartialFileExtension) -and
            ($Destination -eq $testParams.Path) -and
            ($SessionId -eq 1)
        }
    }
    It 'it is removed from the SFTP server after a successful download' {
        Should -Invoke Remove-SFTPItem -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Path -eq $testParams.SftpPath + $testData[0].Name + $testParams.PartialFileExtension) -and
            ($SessionId -eq 1)
        }
    }
    It 'it is renamed on the local file system to its original name' {
        Should -Invoke Rename-Item -Times 1 -Exactly -Scope 'Describe' -ParameterFilter {
            ($NewName -eq $testData[0].Name) -and
            ($LiteralPath -eq ($testParams.Path + '\' + $testData[0].Name + $testParams.PartialFileExtension))
        }
    }
    Context 'an object is returned with property' {
        It 'DateTime' {
            $testResults.DateTime.ToString('yyyyMMdd') |
            Should -Be (Get-Date).ToString('yyyyMMdd')
        }
        Context 'Action' {
            It "<_>" -ForEach @(
                'temp file created',
                'temp file downloaded',
                'temp file removed',
                'temp file renamed',
                'file successfully downloaded'
            ) {
                $testResults.Action | Should -Contain $_
            }
        }
        It 'Downloaded' {
            $testResults.Downloaded | Should -BeTrue
        }
        It 'LocalPath' {
            $testResults.LocalPath | Should -Be $testParams.Path
        }
        It 'FileName' {
            $testResults.FileName | Should -Be $testData[0].Name
        }
        It 'SftpPath' {
            $testResults.SftpPath | Should -Be $testParams.SftpPath
        }
        It 'Error' {
            $testResults.Error | Should -BeNullOrEmpty
        }
        It 'FileLength' {
            $testResults.FileLength | Should -Not -BeNullOrEmpty
            $testResults.FileLength | Should -BeOfType [int]
        }
    }
}
Describe 'OverwriteFile' {
    BeforeAll {
        Mock Remove-Item
        Mock Get-SFTPChildItem {
            [PSCustomObject]@{
                Name     = 'x.txt'
                FullName = $testParams.SftpPath + 'x.txt'
                Length   = 1KB
            }
        }

        $testFile = New-Item -Path $testParams.Path -Name 'x.txt' -ItemType File
    }
    It 'when true the file on the local file system is overwritten' {
        $testNewParams = $testParams.Clone()
        $testNewParams.OverwriteFile = $true

        .$testScript @testNewParams

        Should -Invoke Remove-Item -Times 1 -Exactly -ParameterFilter {
            ($LiteralPath -eq $testFile.FullName)
        }
    }
    It 'when false the file on the local file system is not overwritten' {
        $testNewParams = $testParams.Clone()
        $testNewParams.OverwriteFile = $false

        .$testScript @testNewParams

        Should -Not -Invoke Remove-Item -Times 1 -Exactly -ParameterFilter {
            ($LiteralPath -eq $testFile.FullName)
        }
    }
}
Describe 'when RemoveFailedPartialFiles is true' {
    Context 'remove partial files that are not completely downloaded' {
        It 'from the local folder in Path' {
            $testNewParams = $testParams.Clone()
            $testNewParams.RemoveFailedPartialFiles = $true

            $testFiles = @(
                Join-Path $testNewParams.Path "file.txt"
                Join-Path $testNewParams.Path "file.txt$($testParams.PartialFileExtension)"
            ) | ForEach-Object {
                New-Item -Path $_ -ItemType 'File'
            }

            $testResults = .$testScript @testNewParams

            $testFiles[1].FullName | Should -Not -Exist

            $testResult = $testResults.where(
                { $_.FileName -eq $testFiles[1].Name }
            )

            $testResult.Action | Should -Be "removed failed downloaded partial file '$($testFiles[1].FullName)'"
        }
        It 'from the SFTP server' {
            $testFile = [PSCustomObject]@{
                Name     = "file.txt$($testParams.PartialFileExtension)"
                FullName = $testParams.SftpPath + "file.txt$($testParams.PartialFileExtension)"
            }

            Mock Get-SFTPChildItem {
                $testFile
            }

            $testNewParams = $testParams.Clone()
            $testNewParams.RemoveFailedPartialFiles = $true

            $testResults = .$testScript @testNewParams

            $testResult = $testResults | Where-Object {
                $_.FileName -eq $testFile.Name
            }

            $testResult.Action |
            Should -Be "removed failed downloaded partial file '$($testFile.FullName)'"
        }
    }
}
Describe 'when FileExtensions is' {
    BeforeAll {
        $testFiles = @(
            [PSCustomObject]@{
                Name     = 'file.txt'
                FullName = $testParams.SftpPath + 'file.txt'
                Length   = 1KB
            }
            [PSCustomObject]@{
                Name     = 'file.xml'
                FullName = $testParams.SftpPath + 'file.xml'
                Length   = 1KB
            }
            [PSCustomObject]@{
                Name     = 'file.jpg'
                FullName = $testParams.SftpPath + 'file.jpg'
                Length   = 1KB
            }
        )
        Mock Get-SFTPChildItem {
            $testFiles
        }

        $testNewParams = $testParams.Clone()
    }
    It 'empty, all files are downloaded' {
        $testNewParams.FileExtensions = @()

        .$testScript @testNewParams

        foreach ($testFile in $testFiles) {
            Should -Invoke Get-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Destination -eq $testNewParams.Path) -and
                ($Path -eq ($testFile.FullName + $testNewParams.PartialFileExtension))
            }
        }
    }
    It 'not empty, only specific files are downloaded' {
        $testNewParams.FileExtensions = @('.txt', '.xml')

        .$testScript @testNewParams

        foreach ($testFile in $testFiles.where({ $_.Name -notLike "*.jpg" })) {
            Should -Invoke Get-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Destination -eq $testNewParams.Path) -and
                ($Path -eq ($testFile.FullName + $testNewParams.PartialFileExtension))
            }
        }
        foreach ($testFile in $testFiles.where({ $_.Name -Like "*.jpg" })) {
            Should -Invoke Get-SFTPItem -Not -ParameterFilter {
                ($Destination -eq $testNewParams.Path) -and
                ($Path -eq ($testFile.FullName + $testNewParams.PartialFileExtension))
            }
        }
    }
}
Describe 'when SftpOpenSshKeyFile is used' {
    It 'New-SFTPSession is called with the correct arguments' {
        $testNewParams = $testParams.Clone()
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