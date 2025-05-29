import urllib.request
import subprocess
import os

Office_2019 = {"tool":"https://github.com/Batman123n/Office-installation-script/releases/download/office_installers_v2/Office.2019.eng.exe"}
Office_2021 = {"tool":"https://github.com/Batman123n/Office-installation-script/releases/download/office_installers_v2/Office.2021.eng.exe"}
Office_2024 = {"tool":"https://github.com/Batman123n/Office-installation-script/releases/download/office_installers_v2/Office.2024.eng.exe"}
Office_365 = {"tool":"https://github.com/Batman123n/Office-installation-script/releases/download/office_installers_v2/Office.365.eng.exe"}

office_version = {
    "2019": Office_2019,
    "2021": Office_2021,
    "2024": Office_2024,
    "365": Office_365
}

Response = input("Please chose a version of office. (2019 / 2021 / 2024 / 365)").strip()

if Response in office_version:
    url = office_version [Response]["tool"]
    filename = os.path.basename(url)
    
    print(f"Downloading {filename} ...")
    urllib.request.urlretrieve(url, filename)
    print(f"Download complete.: {filename}")
    
    print("Pokrećem instalaciju...")
    subprocess.run([filename])
else:
    print("Unknown version. Please type one of these: 2019, 2021, 2024 ili 365.")



