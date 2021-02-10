<#
PowerShell library for performing geo IP lookups.

Requirements:
- PowerShell 7.0.0 or Windows PowerShell 5.1.18362.628
  https://github.com/PowerShell/PowerShell/releases

- nuget 5.4.0.6315
  Invoke-WebRequest -Uri 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -OutFile (Join-Path -Path ([Environment]::GetFolderPath('UserProfile')) -ChildPath 'Downloads\nuget.exe')

- MaxMind.DB 2.6.1 nuget package
  Install the nuget package on Windows:
  & (Join-Path -Path ([Environment]::GetFolderPath('UserProfile')) -ChildPath 'Downloads\nuget.exe') Install MaxMind.Db -Version 2.6.1 -OutputDirectory (Join-Path -Path ([Environment]::GetFolderPath('UserProfile')) -ChildPath 'Downloads\.nuget')

  Install the nuget package on macOS with mono:
  & mono (Join-Path -Path ([Environment]::GetFolderPath('UserProfile')) -ChildPath 'Downloads\nuget.exe') Install MaxMind.Db -Version 2.6.1 -OutputDirectory (Join-Path -Path ([Environment]::GetFolderPath('UserProfile')) -ChildPath 'Downloads\.nuget')

- GeoLite2 City and Country binary databases
  Requires free registration to download.
  https://www.maxmind.com

Installation:
- You may need to change your execution policy to import unsigned modules.
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

- Download the module:
  Invoke-WebRequest -Uri 'https://gist.githubusercontent.com/dindoliboon/32d9aa78842d33359c5ce624c570ca96/raw/libMaxMindGeoIp2V1.psm1' -OutFile (Join-Path -Path ([Environment]::GetFolderPath('UserProfile')) -ChildPath 'Downloads\libMaxMindGeoIp2V1.psm1')

- Import the module:
  Import-Module -Name (Join-Path -Path ([Environment]::GetFolderPath('UserProfile')) -ChildPath 'Downloads\libMaxMindGeoIp2V1.psm1')
#>

New-Module -Name libMaxMindGeoIp2V1 -ScriptBlock {
    function New-GeoIp2Reader {
        param (
            $Library,
            $Database
        )

        Add-Type -Path $Library | Out-Null

        return [MaxMind.Db.Reader]::new($Database)
    }

    function Close-GeoIp2Reader {
        param (
            [ref]$Reader
        )

        $Reader.Value.Dispose()
        $Reader.Value = $null
    }

    function Find-GeoIp2 {
        param (
            $Reader,
            $IpAddress,
            $Library,
            $Database
        )

        $results = $null
        $useInternalReader = $Reader -eq $null -and $Library -ne $null -and $Database -ne $null

        $ip = [System.Net.IPAddress]$IpAddress

        if ($useInternalReader) {
            $Reader = New-GeoIp2Reader -Library $Library -Database $Database
        }

        if ($Reader) {
            # Use the first Find method and tell it to return type Dictionary<string, object>.
            $oldMethod = ($Reader.GetType().GetMethods() |? {$_.Name -eq 'Find'})[0]
            $newMethod = $oldMethod.MakeGenericMethod(@([System.Collections.Generic.Dictionary`2[System.String,System.Object]]))

            # Call our new method, T Find[T](ipaddress ipAddress, MaxMind.Db.InjectableValues injectables)
            $results = $newMethod.Invoke($Reader, @($ip, $null))

            if ($useInternalReader) {
                Close-GeoIp2Reader -Reader ([ref]$Reader)
            }
        }
        else {
            throw 'MaxMind.Db.Reader not defined.'
        }

        return $results
    }

    Export-ModuleMember -Function 'New-GeoIp2Reader', 'Find-GeoIp2', 'Close-GeoIp2Reader'
}
