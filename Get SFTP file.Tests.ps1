#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
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
        'SftpPassword', 
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
    } -Tag test 
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
    }  -Tag test 
}
Describe 'download to the SFTP server' {
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
                ($Path -like "*\$($_.Name).DownloadInProgress") -and
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
                ($Path -like "*\$($_.Name).DownloadInProgress") -and
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
            $_.Downloaded | Should -BeTrue
            $_.DateTime | Should -Not -BeNullOrEmpty
            $_.Action | Should -Not -BeNullOrEmpty
            $_.Error | Should -BeNullOrEmpty
        }
    }
}
Describe 'OverwriteFileOnSftpServer' {
    BeforeEach {
        $testParams.Path | ForEach-Object {
            New-Item $_ -ItemType 'File' -ErrorAction Ignore
        }
    }
    It 'when true the file on the SFTP server is overwritten' {
        $testNewParams = $testParams.Clone()
        $testNewParams.OverwriteFileOnSftpServer = $true

        .$testScript @testNewParams

        $testNewParams.Path | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -like "$_.DownloadInProgress") -and
                ($Destination -eq $testNewParams.SftpPath) -and
                ($SessionId -eq 1) -and
                ($Force)
            }
        }
    }
    It 'when false the file on the SFTP server is not overwritten' {
        $testNewParams = $testParams.Clone()
        $testNewParams.OverwriteFileOnSftpServer = $false

        .$testScript @testNewParams

        $testNewParams.Path | ForEach-Object {
            Should -Invoke Set-SFTPItem -Times 1 -Exactly -ParameterFilter {
                ($Path -like "$_.DownloadInProgress") -and
                ($Destination -eq $testNewParams.SftpPath) -and
                ($SessionId -eq 1) -and
                (-not $Force)
            }
        }
    }
}
Describe 'when RemoveFailedPartialFiles is true' {
    Context 'remove partial files that are not completely downloaded' {
        It 'from the local folder in Path' {
            $testNewParams = $testParams.Clone()
            $testNewParams.RemoveFailedPartialFiles = $true
            $testNewParams.Path = (New-Item 'TestDrive:\download' -ItemType 'Directory').FullName

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
        It 'with the same name as a file in Path' {
            $testNewParams = $testParams.Clone()
            $testNewParams.RemoveFailedPartialFiles = $true
            $testNewParams.Path = (New-Item 'TestDrive:\u.txt' -ItemType 'File').FullName

            $testFile = New-Item -Path "$($testNewParams.Path)$($testParams.PartialFileExtension)" -ItemType 'File'
            
            $testResults = .$testScript @testNewParams

            $testFile.FullName | Should -Not -Exist

            $testResult = $testResults.where(
                { $_.FileName -eq $testFile.Name }
            )

            $testResult.Action | Should -Be "removed failed downloaded partial file '$($testFile.FullName)'" 
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
            $testNewParams.Path = (New-Item 'TestDrive:\o.txt' -ItemType 'File').FullName

            $testResults = .$testScript @testNewParams

            $testResult = $testResults.where(
                { $_.FileName -eq $testFile.Name }
            )

            $testResult.Action | Should -Be "removed failed downloaded partial file '$($testFile.FullName)'" 
        }
    }
}
Describe 'when FileExtensions is' {
    BeforeAll {
        $testNewParams = $testParams.Clone()
        $testNewParams.Path = (New-Item 'TestDrive:\download' -ItemType 'Directory').FullName
    }
    It 'empty, all files are downloaded' {
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
    It 'not empty, only specific files are downloaded' {
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