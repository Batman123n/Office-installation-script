# Office Installation Script
This tool uses low level C++ Win32 API and Qt to provide a custom GUI for configuring and installing Microsoft Office using the Office Deployment Tool (ODT).

## Features:
  * Dynamic XML configuration generation
  * Support for multiple Office versions (2019, 2021, 2024, 365)
  * Custom application selection (Word, Excel, PowerPoint, etc.)
  * Embedded `setup_office.exe` for a single-executable experience
  * Automatic cleanup of temporary installation files
  * Statically linked for maximum portability

## Requirements: 
* Visual Studio 2022 with "Desktop development with C++" workload installed.
* Static Qt 6 installation (default path: `D:/QtStatic`).
* CMake.

## How to build:
First clone my repo with this command: \
`git clone https://github.com/Batman123n/Office-installation-script.git`

### Build instructions:
  1. Ensure `setup_office.exe` is present in the project root to embed it.
  2. Run `build.bat`.
 * This will automatically configure the project with CMake, compile the resources, and perform an LTO-optimized build.
 * The final standalone executable will be copied to the project root.

### Disclaimer:
#### Make sure you dont have MS office already installed
#### Use this tool at your own risk. Always ensure you have the appropriate licenses for Microsoft Office.
#### This tool is provided for administrative convenience and follows Microsoft's official deployment patterns.
