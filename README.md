# simple powershell office 365 offline downloader and installer

```
# Get-ODTURL function was taken from Github: mallockey / Install-Office365Suite

function Get-ODTURL {

  [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'

  $MSWebPage | ForEach-Object {
    if ($_ -match 'url=(https://.*officedeploymenttool.*\.exe)') {
      $matches[1]
    }
  }

}
cls
mkdir C:\ODT -ErrorAction SilentlyContinue
write-host " "
Invoke-WebRequest -Uri (Get-ODTURL) -OutFile "C:\ODT\Setup.exe"
write-host " "`nYour browser will open to a page where you customize the office install.`nExport the config file to C:\ODT and save it as "Configuration.xml"`n" "
read-host Press any key to launch https://config.office.com/deploymentsettings
start https://config.office.com/deploymentsettings
write-host " "`n"Press any key once you've saved the Configuration.xml file to C:\ODT"`n" "`nThe download is now beginning...
invoke-command {cmd C:\Setup /download Configuration.xml}
write-host " "`nDownload complete!`n" "
read-host Press any key to begin installation
write-host " "
invoke-command {cmd C:\Setup /configure Configuration.xml}
write-host " "`nOffice installation complete!`n" "
```
