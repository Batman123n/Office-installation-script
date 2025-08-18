md "C:\Office_installation"
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Batman123n/Office-installation-script/refs/heads/main/English%20configurations/Configuration_O2021%20English.xml" -OutFile "C:\Office_installation\Configuration_O2021.xml"
Invoke-WebRequest -Uri "https://aka.ms/ODTDownload" -OutFile "C:\Office_installation\setup_office.exe"
Start-Process -FilePath "C:\Office_installation\setup_office.exe" -ArgumentList "/configure C:\Office_installation\Configuration_O2021.xml" -Verb RunAs -Wait
Remove-Item "C:\Office_installation" -Recurse -Force
