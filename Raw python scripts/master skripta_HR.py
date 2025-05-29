import urllib.request
import subprocess
import os

Office_2019 = {"tool":"https://github.com/Batman123n/Office-installation-script/releases/download/Office_installers-v1/Office_2019_Hrvatski.exe"}
Office_2021 = {"tool":"https://github.com/Batman123n/Office-installation-script/releases/download/Office_installers-v1/Office_2021_Hrvatski.exe"}
Office_2024 = {"tool":"https://github.com/Batman123n/Office-installation-script/releases/download/Office_installers-v1/Office_2024_Hrvatski.exe"}
Office_365 = {"tool":"https://github.com/Batman123n/Office-installation-script/releases/download/Office_installers-v1/Office_365_Hrvatski.exe"}

office_verzije = {
    "2019": Office_2019,
    "2021": Office_2021,
    "2024": Office_2024,
    "365": Office_365
}

Odgovor = input("Odaberite verziju Office-a. (2019 / 2021 / 2024 / 365)").strip()

if Odgovor in office_verzije:
    url = office_verzije[Odgovor]["tool"]
    filename = os.path.basename(url)
    
    print(f"Preuzimam {filename} ...")
    urllib.request.urlretrieve(url, filename)
    print(f"Preuzimanje završeno: {filename}")
    
    print("Pokrećem instalaciju...")
    subprocess.run([filename])
else:
    print("Nepoznata verzija. Molimo unesite 2019, 2021, 2024 ili 365.")



