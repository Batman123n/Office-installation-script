#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QCheckBox>
#include <QPlainTextEdit>
#include <QVBoxLayout>
#include <QPushButton>
#include <QMap>
#include <QLabel>
#include <QComboBox>
#include <QThread>

class InstallerThread : public QThread {
    Q_OBJECT
public:
    QString xmlPath;
    QString targetExePath;
    QString installDir;
    QWidget *parentWidget;

protected:
    void run() override;

signals:
    void finished(bool success);
};

class MainWindow : public QMainWindow {
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);

private slots:
    void updateXML();
    void runInstallation();
    void onInstallationFinished(bool success);

private:
    QMap<QString, QCheckBox*> appCheckBoxes;
    QComboBox *versionBox;
    QComboBox *languageBox;
    QPlainTextEdit *xmlPreview;

    void setupUI();
    QString generateXML();
    void checkOfficeInstallation();
};

#endif // MAINWINDOW_H
