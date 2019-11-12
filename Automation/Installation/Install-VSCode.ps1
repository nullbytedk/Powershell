<#PSScriptInfo

.VERSION 1.0

.AUTHOR Microsoft

.MODIFIEDBY NullbyteDK

.COPYRIGHT (c) Microsoft Corporation

.TAGS install vscode installer

.LICENSEURI https://github.com/PowerShell/vscode-powershell/blob/develop/LICENSE.txt

.ORIGINALPROJECTURI https://github.com/PowerShell/vscode-powershell/blob/develop/scripts/Install-VSCode.ps1

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
    Initial release.
#>

<#
.SYNOPSIS
    Installs Visual Studio Code, the PowerShell extension, and optionally
    a list of additional extensions.

.DESCRIPTION
    This script can be used to easily install Visual Studio Code and the
    PowerShell extension on your machine.  You may also specify additional
    extensions to be installed using the -AdditionalExtensions parameter.
    The -LaunchWhenDone parameter will cause VS Code to be launched as
    soon as installation has completed.

    This project is a modified version of the original found on the following GitHub:
    https://github.com/PowerShell/vscode-powershell/blob/develop/scripts/Install-VSCode.ps1

.PARAMETER Architecture
    A validated string defining the bit version to download. Values can be either 64-bit or 32-bit.
    If 64-bit is chosen and the OS Architecture does not match, then the 32-bit build will be
    downloaded instead. If parameter is not used, then 64-bit is used as default.

.PARAMETER BuildEdition
    A validated string defining which build edition or "stream" to download:
    Stable or Insiders Edition (system install or user profile install).
    If the parameter is not used, then stable is downloaded as default.


.PARAMETER AdditionalExtensions
    An array of strings that are the fully-qualified names of extensions to be
    installed in addition to the PowerShell extension.  The fully qualified
    name is formatted as "<publisher name>.<extension name>" and can be found
    next to the extension's name in the details tab that appears when you
    click an extension in the Extensions panel in Visual Studio Code.

.PARAMETER LaunchWhenDone
    When present, causes Visual Studio Code to be launched as soon as installation
    has finished.

.PARAMETER EnableContextMenus
    When present, causes the installer to configure the Explorer context menus

.EXAMPLE
    Install-VSCode.ps1 -Architecture 32-bit

    Installs Visual Studio Code (32-bit) and the powershell extension.
.EXAMPLE
    Install-VSCode.ps1 -LaunchWhenDone

    Installs Visual Studio Code (64-bit) and the PowerShell extension and then launches
    the editor after installation completes.

.EXAMPLE
    Install-VSCode.ps1 -AdditionalExtensions 'eamodio.gitlens', 'vscodevim.vim'

    Installs Visual Studio Code (64-bit), the PowerShell extension, and additional
    extensions.

.EXAMPLE
    Install-VSCode.ps1 -BuildEdition Insider-User -LaunchWhenDone

    Installs Visual Studio Code Insiders Edition (64-bit) to the user profile and then launches the editor
    after installation completes.

.NOTES
    This script is licensed under the MIT License:

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [parameter()]
    [ValidateSet('64-bit', '32-bit')]
    [string]$Architecture = '64-bit',

    [parameter()]
    [ValidateSet('Stable-System', 'Stable-User', 'Insider-System', 'Insider-User')]
    [string]$BuildEdition = "Stable-System",

    [Parameter()]
    [ValidateNotNull()]
    [string[]]$AdditionalExtensions = @(),

    [switch]$LaunchWhenDone,

    [switch]$EnableContextMenus
)


function Test-IsOsArchX64 {
    if ($PSVersionTable.PSVersion.Major -lt 6) {
        return (Get-CimInstance -ClassName Win32_OperatingSystem).OSArchitecture -eq '64-bit'
    }

    return [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq [System.Runtime.InteropServices.Architecture]::X64
}

function Get-CodePlatformInformation {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('32-bit', '64-bit')]
        [string]
        $Bitness,

        [Parameter(Mandatory=$true)]
        [ValidateSet('Stable-System', 'Stable-User', 'Insider-System', 'Insider-User')]
        [string]
        $BuildEdition
    )


    if ($Bitness -ne '64-bit' -and $os -ne 'Windows') {
        throw "Non-64-bit *nix systems are not supported"
    }

    if ($BuildEdition.EndsWith('User') -and $os -ne 'Windows') {
        throw 'User builds are not available for non-Windows systems'
    }

    switch ($BuildEdition) {
        'Stable-System' {
            $appName = "Visual Studio Code ($Bitness)"
            break
        }

        'Stable-User' {
            $appName = "Visual Studio Code ($($Architecture) - User)"
            break
        }

        'Insider-System' {
            $appName = "Visual Studio Code - Insiders Edition ($Bitness)"
            break
        }

        'Insider-User' {
            $appName = "Visual Studio Code - Insiders Edition ($($Architecture) - User)"
            break
        }
    }

    
    $ext = 'exe'
    switch ($Bitness) {
        '32-bit' {
            $platform = 'win32'

            if (Test-IsOsArchX64) {
                $installBase = ${env:ProgramFiles(x86)}
                break
            }

            $installBase = ${env:ProgramFiles}
            break
        }

        '64-bit' {
            $installBase = ${env:ProgramFiles}

            if (Test-IsOsArchX64) {
                $platform = 'win32-x64'
                break
            }

            Write-Warning '64-bit install requested on 32-bit system. Installing 32-bit VSCode'
            $platform = 'win32'
            break
        }
    }  

    switch ($BuildEdition) {
        'Stable-System' {
            $exePath = "$installBase\Microsoft VS Code\bin\code.cmd"
            $channel = 'stable'
            break
        }

        'Stable-User' {
            $exePath = "${env:LocalAppData}\Programs\Microsoft VS Code\bin\code.cmd"
            $channel = 'stable'
            $platform += '-user'
            break
        }

        'Insider-System' {
            $exePath = "$installBase\Microsoft VS Code Insiders\bin\code-insiders.cmd"
            $channel = 'insider'
            break
        }

        'Insider-User' {
            $exePath = "${env:LocalAppData}\Programs\Microsoft VS Code Insiders\bin\code-insiders.cmd"
            $channel = 'insider'
            $platform += '-user'
            break
        }
    }

    $info = @{
        AppName = $appName
        ExePath = $exePath
        Platform = $platform
        Channel = $channel
        FileUri = "https://vscode-update.azurewebsites.net/latest/$platform/$channel"
        Extension = $ext
    }


    return $info
}

function Save-WithBitsTransfer {
    param(
        [Parameter(Mandatory=$true)]
        [string]
        $FileUri,

        [Parameter(Mandatory=$true)]
        [string]
        $Destination,

        [Parameter(Mandatory=$true)]
        [string]
        $AppName
    )

    Write-Host "`nDownloading latest $AppName..." -ForegroundColor Yellow

    Remove-Item -Force $Destination -ErrorAction SilentlyContinue

    $bitsDl = Start-BitsTransfer $FileUri -Destination $Destination -Asynchronous

    while (($bitsDL.JobState -eq 'Transferring') -or ($bitsDL.JobState -eq 'Connecting')) {
        Write-Progress -Activity "Downloading: $AppName" -Status "$([math]::round($bitsDl.BytesTransferred / 1mb))mb / $([math]::round($bitsDl.BytesTotal / 1mb))mb" -PercentComplete ($($bitsDl.BytesTransferred) / $($bitsDl.BytesTotal) * 100 )
    }

    switch ($bitsDl.JobState) {

        'Transferred' {
            Complete-BitsTransfer -BitsJob $bitsDl
            break
        }

        'Error' {
            throw 'Error downloading installation media.'
        }
    }
}


try {
    $prevProgressPreference = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'

    # Get information required for installation
    $codePlatformInfo = Get-CodePlatformInformation -Bitness $Architecture -BuildEdition $BuildEdition

    # Download the installer
    $tmpdir = [System.IO.Path]::GetTempPath()

    $ext = $codePlatformInfo.Extension
    $installerName = "vscode-install.$ext"

    $installerPath = [System.IO.Path]::Combine($tmpdir, $installerName)

    if ($PSVersionTable.PSVersion.Major -le 5) {
        Save-WithBitsTransfer -FileUri $codePlatformInfo.FileUri -Destination $installerPath -AppName $codePlatformInfo.AppName
    }
    # We don't want to use RPM packages -- see the installation step below
    elseif ($codePlatformInfo.Extension -ne 'rpm') {
        if ($PSCmdlet.ShouldProcess($codePlatformInfo.FileUri, "Invoke-WebRequest -OutFile $installerPath")) {
            Invoke-WebRequest -Uri $codePlatformInfo.FileUri -OutFile $installerPath
        }
    }

    # Install VSCode
    $exeArgs = '/verysilent /tasks=addtopath'
    if ($EnableContextMenus) {
        $exeArgs = '/verysilent /tasks=addcontextmenufiles,addcontextmenufolders,addtopath'
    }

    if (-not $PSCmdlet.ShouldProcess("$installerPath $exeArgs", 'Start-Process -Wait')) {
        break
    }

    Start-Process -Wait $installerPath -ArgumentList $exeArgs
    $env:Path += ";C:\Program Files\Microsoft VS Code\bin"

    $codeExePath = $codePlatformInfo.ExePath

    # Install any extensions
    $extensions = @("ms-vscode.PowerShell","pkief.material-icon-theme") + $AdditionalExtensions
    if ($PSCmdlet.ShouldProcess(($extensions -join ','), "$codeExePath --install-extension")) {    
        foreach ($extension in $extensions) {
            Write-Host "`nInstalling extension $extension..." -ForegroundColor Yellow
            & $codeExePath --install-extension $extension
        }
    }


    # Install custom settings - DRAFT
    <#$filepath = "$($env:APPDATA)\Code\User\settings.json" 

    $script:content = $null
    if(Test-Path -Path $filepath -PathType Leaf){
        $item = Get-Item $filepath
        $script:content = Get-Content -Path $item.FullName
    }else{
        $item = New-Item $filepath | Out-Null
        $script:content = "{`n}"
        Add-Content -Path $filepath -Value $script:content
    }
    
    $settings = $script:content | ConvertFrom-Json

    #Specify settings
        #$settings | Add-Member -MemberType NoteProperty -Name "editor.autoClosingBrackets" -Value "always"
        #$settings | Add-Member -MemberType NoteProperty -Name "editor.aaaa" -Value "never"
        #$settings | Add-Member -MemberType NoteProperty -Name "editor.bbbb" -Value "maybe"
        #$settings | Add-Member -MemberType NoteProperty -Name "editor.cccc" -Value "yes"

    $json = $settings | ConvertTo-Json
    Set-Content -Path $filepath -Value $json#>


    # Launch if requested
    if ($LaunchWhenDone) {
        $appName = $codePlatformInfo.AppName

        if (-not $PSCmdlet.ShouldProcess($appName, "Launch with $codeExePath")) {
            return
        }

        Write-Host "`nInstallation complete, starting $appName...`n`n" -ForegroundColor Green
        & $codeExePath
        return
    }

    if ($PSCmdlet.ShouldProcess('Installation complete!', 'Write-Host')) {
        Write-Host "`nInstallation complete!`n`n" -ForegroundColor Green
    }
}
finally {
    $ProgressPreference = $prevProgressPreference
}
