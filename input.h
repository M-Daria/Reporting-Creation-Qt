#ifndef INPUT_H
#define INPUT_H

#include <QWidget>
#include <QDir>
#include <QAxObject>

QT_BEGIN_NAMESPACE
class QLabel;
class QPushButton;
class QRadioButton;
class QTextEdit;
class QLineEdit;
class QDateEdit;
QT_END_NAMESPACE

class Input : public QWidget
{
    Q_OBJECT

public:
    Input(QWidget *parent = 0);

    void addWord(int, QString, QString, int, int, QString, QString, QDate, QString);
    int addExcel(QString, QString, QDate, QString, QString, QString);
    void replaceMark(QAxObject *, QString, QString);
    void startExcel(QAxObject *, QString, int);

public slots:
    void addData();
    void noData();
    void Data();
    void chExPath();

private:
    QLineEdit *numText;
    QLineEdit *summText;
    QLineEdit *fileText;
    QLineEdit *nameParText;
    QLineEdit *phoneText;
    QLineEdit *nameChildText;
    QDateEdit *bdChildText;
    QDateEdit *dateText;

    QLabel *numLabel;
    QLabel *summLabel;
    QLabel *dateLabel;
    QLabel *exLabel;
    QLabel *nameParLabel;
    QLabel *phoneLabel;
    QLabel *nameChildLabel;
    QLabel *bdChildLabel;

    QPushButton *addButton;
    QPushButton *excelPath;
    QRadioButton *butNoData;
    QRadioButton *butData;

    QString pathToExcel;
};

#endif // INPUT_H
