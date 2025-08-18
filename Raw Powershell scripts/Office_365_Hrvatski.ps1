md "C:\Office_instalacija"
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Batman123n/Office-installation-script/refs/heads/main/Croatian%20configurations/Configuration_O365%20Hrvatski.xml" -OutFile "C:\Office_instalacija\Configuration_O365.xml"
Invoke-WebRequest -Uri "https://github.com/Batman123n/Office123/raw/refs/heads/main/setup_office.exe" -OutFile "C:\Office_instalacija\setup_office.exe"
Start-Process -FilePath "C:\Office_instalacija\setup_office.exe" -ArgumentList "/configure C:\Office_instalacija\Configuration_O365.xml" -Verb RunAs -Wait
Remove-Item "C:\Office_instalacija" -Recurse -Force
