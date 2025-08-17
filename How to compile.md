#To compile you need VS 2022 or your compiler of choice
To compile in VS 2022 you select "windows desktop application" and paste the code from the main branch from folder:
"C++ GUI Installers" and build it.
On command line you need to manualy install c++ compiler and compile it yourself.
With cl you use something like this: cl "path to DarioInstallsOffice.cpp" /EHsc /Fe:"path to DarioInstallsOffice.exe" user32.lib gdi32.lib shell32.lib urlmon.lib
You need to install VS build tools to use cl compiler.
I recomend VS 2022 method which is way easier.
