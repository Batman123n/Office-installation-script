#include "mainwindow.h"
#include <QScrollArea>
#include <QFileDialog>
#include <QFile>
#include <QTextStream>
#include <QMessageBox>
#include <QLabel>
#include <QDir>
#include <QResource>
#include <QScreen>
#include <QCoreApplication>
#include <windows.h>
#include <urlmon.h>
#include <shellapi.h>

#pragma comment(lib, "urlmon.lib")
#pragma comment(lib, "shell32.lib")

void InstallerThread::run() {
    SHELLEXECUTEINFO sei = { sizeof(sei) };
    sei.fMask = SEE_MASK_NOCLOSEPROCESS;
    sei.hwnd = (HWND)parentWidget->winId();
    sei.lpVerb = L"runas";
    sei.lpFile = (LPCWSTR)targetExePath.utf16();
    QString args = "/configure " + xmlPath;
    sei.lpParameters = (LPCWSTR)args.utf16();
    sei.nShow = SW_HIDE;

    if (ShellExecuteEx(&sei)) {
        WaitForSingleObject(sei.hProcess, INFINITE);
        CloseHandle(sei.hProcess);
        QDir("C:\\Office_instalacija").removeRecursively();
        emit finished(true);
    } else {
        emit finished(false);
    }
}

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent) {
    setupUI();
    updateXML();
    int width = 420;
    int length = 500;
    resize(width, length);
}

void MainWindow::setupUI() {
    QWidget *centralWidget = new QWidget(this);
    QVBoxLayout *mainLayout = new QVBoxLayout(centralWidget);

    QLabel *versionLabel = new QLabel("<b>Select office version:</b>");
    mainLayout->addWidget(versionLabel);

    versionBox = new QComboBox(this);
    versionBox->addItems({"2019", "2021", "2024", "365"});
    connect(versionBox, &QComboBox::currentIndexChanged, this, &MainWindow::updateXML);
    mainLayout->addWidget(versionBox);

    // List of apps in specific order
    QList<QPair<QString, QString>> apps = {
        {"Word", "Word"}, {"PowerPoint", "PowerPoint"}, {"Excel", "Excel"},
        {"Publisher", "Publisher"}, {"Outlook", "Outlook"}, {"Access", "Access"},
        {"Teams", "Teams"}, {"OneDrive", "OneDrive"}, {"OneNote", "OneNote"}
    };

    QLabel *header = new QLabel("<b>Select Apps to install:</b>");
    mainLayout->addWidget(header);

    // Group apps into pairs for 2 columns
    for (int i = 0; i < apps.size(); i += 2) {
        QHBoxLayout *rowLayout = new QHBoxLayout();
        for (int j = 0; j < 2 && (i + j) < apps.size(); ++j) {
            QPair<QString, QString> app = apps[i + j];
            QCheckBox *cb = new QCheckBox(app.first, this);

            // Uncheck these by default
            bool defaultChecked = (app.second != "Teams" && app.second != "OneDrive" && app.second != "OneNote" && app.second != "Outlook");
            cb->setChecked(defaultChecked);

            connect(cb, &QCheckBox::stateChanged, this, &MainWindow::updateXML);
            appCheckBoxes.insert(app.second, cb);
            rowLayout->addWidget(cb);
        }
        mainLayout->addLayout(rowLayout);
    }

    xmlPreview = new QPlainTextEdit();
    xmlPreview->setReadOnly(true);
    xmlPreview->setMaximumHeight(180);
    mainLayout->addWidget(xmlPreview);

    QPushButton *installBtn = new QPushButton("Install", this);
    connect(installBtn, &QPushButton::clicked, this, &MainWindow::runInstallation);
    mainLayout->addWidget(installBtn);

    setCentralWidget(centralWidget);
    setWindowTitle("Dario installs MS Office");
}

QString MainWindow::generateXML() {
    QString version = versionBox->currentText();
    QString xml;

    // Common forced exclusions
    QString forcedExclusions = "      <ExcludeApp ID=\"Groove\" />\n"
                               "      <ExcludeApp ID=\"Lync\" />\n"
                               "      <ExcludeApp ID=\"OneDrive\" />\n";

    if (version == "2019") {
        xml = "<Configuration>\n";
        xml += "  <Add OfficeClientEdition=\"64\" Channel=\"PerpetualVL2019\">\n";
        xml += "    <Product ID=\"ProPlus2019Volume\">\n";
        xml += "      <Language ID=\"hr-hr\" />\n";
        xml += forcedExclusions;

        for (auto it = appCheckBoxes.begin(); it != appCheckBoxes.end(); ++it) {
            if (!it.value()->isChecked() && it.key() != "Groove" && it.key() != "Lync" && it.key() != "OneDrive") {
                xml += QString("      <ExcludeApp ID=\"%1\" />\n").arg(it.key());
            }
        }
        xml += "    </Product>\n";
        xml += "  </Add>\n";
        xml += "  <RemoveMSI />\n";
        xml += "</Configuration>";
    } else if (version == "2021" || version == "365") {
        xml = "<Configuration ID=\"dcffe8aa-735c-4778-a114-dac4b35064e8\">\n";
        xml += "  <Add OfficeClientEdition=\"64\" Channel=\"Current\">\n";
        xml += "    <Product ID=\"O365ProPlusEEANoTeamsRetail\">\n";
        xml += "      <Language ID=\"hr-hr\" />\n";
        xml += forcedExclusions;

        for (auto it = appCheckBoxes.begin(); it != appCheckBoxes.end(); ++it) {
            if (!it.value()->isChecked() && it.key() != "Groove" && it.key() != "Lync" && it.key() != "OneDrive") {
                xml += QString("      <ExcludeApp ID=\"%1\" />\n").arg(it.key());
            }
        }
        xml += "    </Product>\n";
        xml += "  </Add>\n";
        xml += "  <Updates Enabled=\"TRUE\" />\n";
        xml += "  <RemoveMSI />\n";
        xml += "  <AppSettings>\n";
        xml += "    <User Key=\"software\\microsoft\\office\\16.0\\excel\\options\" Name=\"defaultformat\" Value=\"51\" Type=\"REG_DWORD\" App=\"excel16\" Id=\"L_SaveExcelfilesas\" />\n";
        xml += "    <User Key=\"software\\microsoft\\office\\16.0\\powerpoint\\options\" Name=\"defaultformat\" Value=\"27\" Type=\"REG_DWORD\" App=\"ppt16\" Id=\"L_SavePowerPointfilesas\" />\n";
        xml += "    <User Key=\"software\\microsoft\\office\\16.0\\word\\options\" Name=\"defaultformat\" Value=\"\" Type=\"REG_SZ\" App=\"word16\" Id=\"L_SaveWordfilesas\" />\n";
        xml += "  </AppSettings>\n";
        xml += "</Configuration>";
    } else if (version == "2024") {
        xml = "<Configuration>\n";
        xml += "  <Add OfficeClientEdition=\"64\" Channel=\"Current\">\n";
        xml += "    <Product ID=\"ProPlus2024Retail\">\n";
        xml += "      <Language ID=\"hr-hr\" />\n";
        xml += forcedExclusions;

        for (auto it = appCheckBoxes.begin(); it != appCheckBoxes.end(); ++it) {
            if (!it.value()->isChecked() && it.key() != "Groove" && it.key() != "Lync" && it.key() != "OneDrive") {
                xml += QString("      <ExcludeApp ID=\"%1\" />\n").arg(it.key());
            }
        }
        xml += "    </Product>\n";
        xml += "  </Add>\n";
        xml += "  <Property Name=\"SharedComputerLicensing\" Value=\"0\" />\n";
        xml += "  <Property Name=\"FORCEAPPSHUTDOWN\" Value=\"FALSE\" />\n";
        xml += "  <Property Name=\"DeviceBasedLicensing\" Value=\"0\" />\n";
        xml += "  <Property Name=\"SCLCacheOverride\" Value=\"0\" />\n";
        xml += "  <Property Name=\"AUTOACTIVATE\" Value=\"0\" />\n";
        xml += "  <Updates Enabled=\"TRUE\" />\n";
        xml += "  <RemoveMSI />\n";
        xml += "  <AppSettings>\n";
        xml += "    <User Key=\"software\\microsoft\\office\\16.0\\excel\\options\" Name=\"defaultformat\" Value=\"51\" Type=\"REG_DWORD\" App=\"excel16\" Id=\"L_SaveExcelfilesas\" />\n";
        xml += "    <User Key=\"software\\microsoft\\office\\16.0\\powerpoint\\options\" Name=\"defaultformat\" Value=\"27\" Type=\"REG_DWORD\" App=\"ppt16\" Id=\"L_SavePowerPointfilesas\" />\n";
        xml += "    <User Key=\"software\\microsoft\\office\\16.0\\word\\options\" Name=\"defaultformat\" Value=\"\" Type=\"REG_SZ\" App=\"word16\" Id=\"L_SaveWordfilesas\" />\n";
        xml += "  </AppSettings>\n";
        xml += "</Configuration>";
    } else {
        xml = "<Configuration>\n";
        xml += "  <Add OfficeClientEdition=\"64\">\n";
        xml += "    <Product ID=\"O365ProPlusRetail\">\n";
        xml += "      <Language ID=\"hr-hr\" />\n";

        for (auto it = appCheckBoxes.begin(); it != appCheckBoxes.end(); ++it) {
            if (!it.value()->isChecked()) {
                xml += QString("      <ExcludeApp ID=\"%1\" />\n").arg(it.key());
            }
        }
        xml += "    </Product>\n";
        xml += "  </Add>\n";
        xml += "</Configuration>";
    }
    return xml;
}

void MainWindow::updateXML() {
    xmlPreview->setPlainText(generateXML());
}

void MainWindow::runInstallation() {
    QString installDir = "C:\\Office_instalacija";
    QDir().mkpath(installDir);

    QString xmlPath = installDir + "\\Configuration_O2019.xml";
    QString targetExePath = installDir + "\\setup_office.exe";

    // Save XML
    QFile xmlFile(xmlPath);
    if (!xmlFile.open(QIODevice::WriteOnly | QIODevice::Text)) {
        QMessageBox::critical(this, "Error", "Could not create configuration file.");
        return;
    }
    xmlFile.write(generateXML().toUtf8());
    xmlFile.close();

    // Copy from internal resources (embedded)
    QString sourcePath = ":/setup_office.exe";

    if (!QFile::exists(sourcePath)) {
        QMessageBox::critical(this, "Error", "The 'setup_office.exe' was not embedded in this build. Please check your resources.");
        return;
    }

    // Force kill and cleanup to ensure we can overwrite
    system("taskkill /F /IM setup_office.exe /T 2>nul");
    Sleep(500);

    if (QFile::exists(targetExePath)) {
        if (!QFile::remove(targetExePath)) {
            // Try one more time after a short delay
            Sleep(500);
            QFile::remove(targetExePath);
        }
    }

    if (!QFile::copy(sourcePath, targetExePath)) {
        QMessageBox::critical(this, "Error", "Could not extract the embedded setup file to " + targetExePath);
        return;
    }
    InstallerThread *thread = new InstallerThread();
    thread->xmlPath = xmlPath;
    thread->targetExePath = targetExePath;
    thread->parentWidget = this;
    connect(thread, &InstallerThread::finished, this, &MainWindow::onInstallationFinished);
    connect(thread, &InstallerThread::finished, thread, &InstallerThread::deleteLater);
    thread->start();
}

void MainWindow::onInstallationFinished(bool success) {
    if (!success) {
        QMessageBox::critical(this, "Error", "Failed to run setup.");
    }
}
