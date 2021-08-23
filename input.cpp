#include "input.h"

#include <QtWidgets>
#include "stdlib.h"

QString toWord(int num)
{
    switch (num)
    {
        case 1: return "января";
        case 2: return "февраля";
        case 3: return "марта";
        case 4: return "апреля";
        case 5: return "мая";
        case 6: return "июня";
        case 7: return "июля";
        case 8: return "августа";
        case 9: return "сентября";
        case 10: return "октября";
        case 11: return "ноября";
        case 12: return "декабря";
        default: return "";
    }
}

Input::Input(QWidget *parent)
    : QWidget(parent)
{
    QFont font("Helvetica", 10);
    int i = 0;

    numLabel = new QLabel(tr("Номер договора:"));
    numLabel->setFont(font);
    dateLabel = new QLabel(tr("Дата заключения:"));
    dateLabel->setFont(font);
    nameParLabel = new QLabel(tr("ФИО заказчика:"));
    nameParLabel->setFont(font);
    phoneLabel = new QLabel(tr("Телефон:"));
    phoneLabel->setFont(font);
    nameChildLabel = new QLabel(tr("ФИО обучающегося:"));
    nameChildLabel->setFont(font);
    bdChildLabel = new QLabel(tr("Дата рождения:"));
    bdChildLabel->setFont(font);
    summLabel = new QLabel(tr("Сумма:"));
    summLabel->setFont(font);
    exLabel = new QLabel(tr("Выбор файла Excel:"));
    exLabel->setFont(font);

    butNoData = new QRadioButton(tr("Новый файл"));
    butNoData->setFont(font);
    butData = new QRadioButton(tr("Выбрать файл:"));
    butData->setChecked(true);
    butData->setFont(font);

    excelPath = new QPushButton(tr("Выбрать"), this);
    excelPath->setStyleSheet("padding: 5px 15px");
    excelPath->setFont(font);

    addButton = new QPushButton(tr("Done"));
    addButton->setStyleSheet("margin-top: 30px; padding: 5px 15px");
    addButton->setFont(font);
    addButton->setMinimumSize(excelPath->size());

    numText = new QLineEdit;
    numText->setPlaceholderText("Л2");

    dateText = new QDateEdit(QDate::currentDate());
    dateText->setCalendarPopup(true);
    bdChildText = new QDateEdit;
    bdChildText->setCalendarPopup(true);

    nameParText = new QLineEdit;
    nameParText->setPlaceholderText("Иванов Иван Иванович");
    phoneText = new QLineEdit;
    phoneText->setPlaceholderText("+7(909)999-00-00");
    nameChildText = new QLineEdit;
    nameChildText->setPlaceholderText("Иванов Алексей Иванович");
    summText = new QLineEdit;
    summText->setPlaceholderText("19 500,00 (Девятнадцать тысяч пятьсот)");

    fileText = new QLineEdit;
    fileText->setPlaceholderText("Выбрать путь...");

    QGridLayout *gLayout = new QGridLayout;
    QVBoxLayout *vlayout = new QVBoxLayout;
    QHBoxLayout *hlayout1 = new QHBoxLayout, *hlayout2 = new QHBoxLayout, *hlayout3 = new QHBoxLayout;

    gLayout->addWidget(numLabel, i, 0, Qt::AlignRight);
    hlayout1->addWidget(numText, 1);
    hlayout1->addWidget(dateLabel, 1, Qt::AlignRight);
    hlayout1->addWidget(dateText, 1);
    gLayout->addLayout(hlayout1, i, 1);

    gLayout->addWidget(nameParLabel, ++i, 0, Qt::AlignRight);
    hlayout2->addWidget(nameParText, 3);
    hlayout2->addWidget(phoneLabel, 1, Qt::AlignRight);
    hlayout2->addWidget(phoneText, 2);
    gLayout->addLayout(hlayout2, i, 1);

    gLayout->addWidget(nameChildLabel, ++i, 0, Qt::AlignRight);
    hlayout3->addWidget(nameChildText, 1);
    hlayout3->addWidget(bdChildLabel, 1, Qt::AlignRight);
    hlayout3->addWidget(bdChildText, 1);
    gLayout->addLayout(hlayout3, i, 1);

    gLayout->addWidget(summLabel, ++i, 0, Qt::AlignTop | Qt::AlignRight);
    gLayout->addWidget(summText, i, 1);

    gLayout->addWidget(exLabel, ++i, 0, Qt::AlignTop | Qt::AlignRight);
    vlayout->addWidget(butNoData);
    vlayout->addWidget(butData);
    vlayout->addWidget(fileText);
    gLayout->addWidget(excelPath, ++i, 1, Qt::AlignRight);

    gLayout->addLayout(vlayout, --i, 1);
    gLayout->addWidget(addButton, i += 2, 1, Qt::AlignRight);

    setLayout(gLayout);

    connect(addButton, &QAbstractButton::clicked, this, &Input::addData);
    connect(butData, &QAbstractButton::clicked, this, &Input::Data);
    connect(butNoData, &QAbstractButton::clicked, this, &Input::noData);
    connect(excelPath, &QAbstractButton::clicked, this, &Input::chExPath);
}

void Input::noData()
{
    pathToExcel = "";
    fileText->setEnabled(false);
    excelPath->setEnabled(false);
}

void Input::Data()
{
    fileText->setEnabled(true);
    excelPath->setEnabled(true);
}

void Input::chExPath()
{
    QString open_path = QFileDialog::getOpenFileName(0, tr("Выбрать файл excel"), "");
    if (open_path.isNull()) return;
    pathToExcel = open_path;
    fileText->setText(pathToExcel);
}

void Input::addData()
{
    if (pathToExcel.isEmpty() && butData->isChecked())
    {
        QMessageBox::information(0, "Сообщение", "Путь к файлу excel не выбран!");
        return;
    }

    int count = addExcel(numText->text(), summText->text(), dateText->date(), nameParText->text(), phoneText->text(), nameChildText->text());
    addWord(count, numText->text(), summText->text(), dateText->date().day(), dateText->date().month(),
            nameParText->text(), nameChildText->text(), bdChildText->date(), phoneText->text());
}

int Input::addExcel(QString num, QString summ, QDate date, QString namePar, QString phone, QString nameCh)
{
    QAxObject *excel = new QAxObject("Excel.Application");
    QAxObject *workbook = excel->querySubObject("Workbooks");

    if (!pathToExcel.isEmpty()) workbook = workbook->querySubObject("Open(path)", pathToExcel);
    else workbook = workbook->querySubObject("Add()");

    excel->setProperty("Visible", true);

    int count;

    QAxObject *sheet = workbook->querySubObject("Sheets");
    sheet = sheet->querySubObject("Item(int)", 1);

    if (pathToExcel.isEmpty())
    {
        startExcel(sheet, "№", 1);
        startExcel(sheet, "№ Договора", 1);
        startExcel(sheet, "Дата", 1);
        startExcel(sheet, "Ф.И.О заключившего договор", 1);
        startExcel(sheet, "Телефон родителей", 1);
        startExcel(sheet, "Ф.И.О. обучающегося", 1);
        startExcel(sheet, "Сумма", 1);
        startExcel(sheet, "Программа обучения", 1);
        startExcel(sheet, "Заметки", 1);
    }

    QAxObject *range;

    range = sheet->querySubObject("Range(const QVariant&)", QVariant(QString("A1:A65536")));
    range = range->querySubObject("SpecialCells(int)", 2);
    count = range->dynamicCall("Count()").toInt();

    startExcel(sheet, QString::number(count), count + 1);
    startExcel(sheet, num, count + 1);
    startExcel(sheet, date.toString("dd.MM.yyyy"), count + 1);
    startExcel(sheet, namePar, count + 1);
    startExcel(sheet, phone, count + 1);
    startExcel(sheet, nameCh, count + 1);
    startExcel(sheet, summ.left(summ.indexOf('(')), count + 1);
    startExcel(sheet, "", count + 1);
    startExcel(sheet, "", count + 1);

    delete range;
    delete sheet;
    delete workbook;
    delete excel;

    return count;
}

void Input::addWord(int count, QString num, QString summ, int day, int month, QString namePar, QString nameCh, QDate bd, QString phone)
{
    QAxObject *word = new QAxObject("Word.Application");

    word->setProperty("Visible", true);

    QAxObject *doc = word->querySubObject("Documents");
    doc = doc->querySubObject("Open(path)", QDir::toNativeSeparators(QDir::currentPath() + "/Договор Шаблон.docx"));

    QAxObject *wordSelection = word->querySubObject("Selection");
    QAxObject *find = wordSelection->querySubObject("Find");

    find->dynamicCall("ClearFormatting()");

    QString sCount;
    if (count < 10) sCount = "000" + QString::number(count);
    else if (count < 100) sCount = "00" + QString::number(count);
    else if (count < 1000) sCount = "0" + QString::number(count);

    replaceMark(find, "%Номер_договора%", sCount + "/ " + num);
    replaceMark(find, "%Заказчик%", namePar);
    //replaceMark(find, "%ФамилияИО%", );
    replaceMark(find, "%Телефон%", phone);
    replaceMark(find, "%Ребёнок%", nameCh);
    replaceMark(find, "%ДР_Ребёнка%", bd.toString("dd.MM.yyyy"));
    replaceMark(find, "%Сумма%", summ);
    replaceMark(find, "%Ч%", "0" + QString::number(day));
    replaceMark(find, "%М%", toWord(month));

    delete find;
    delete wordSelection;

    QString save_path = QFileDialog::getSaveFileName(0, tr("Сохранить договор"), "");
    if (save_path.isEmpty()) return;
    doc->dynamicCall("SaveAs(const QString&)", save_path);
    word->dynamicCall("Quit()");

    delete doc;
    delete word;
}

void Input::replaceMark(QAxObject *find, QString umark, QString utext)
{
    QList <QVariant> params;
    params.operator << (QVariant(umark));
    params.operator << (QVariant("0"));
    params.operator << (QVariant("0"));
    params.operator << (QVariant("0"));
    params.operator << (QVariant("0"));
    params.operator << (QVariant("0"));
    params.operator << (QVariant(true));
    params.operator << (QVariant("0"));
    params.operator << (QVariant("0"));
    params.operator << (QVariant(utext));
    params.operator << (QVariant("2"));
    params.operator << (QVariant("0"));
    params.operator << (QVariant("0"));
    params.operator << (QVariant("0"));
    params.operator << (QVariant("0"));
    find->dynamicCall("Execute(const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,const QVariant&)",
                      params);
}

void Input::startExcel(QAxObject *sheet, QString text, int count)
{
    QAxObject *range;
    static int i = 0;

    range = sheet->querySubObject("Cells(int, int)", count, i % 9 + 1);
    range->dynamicCall("setValue(const QVariant&)", QVariant(text));
    if (count == 1)
    {
        QAxObject *font = range->querySubObject("Font");
        font->setProperty("Bold", true);
        delete font;
    }

    QAxObject *borders = range->querySubObject("Borders");
    borders->setProperty("LineStyle", true);

    QAxObject *col = range->querySubObject("EntireColumn");
    col->dynamicCall("AutoFit()");
    col->dynamicCall("HorizontalAlignment", -4108);
    col->dynamicCall("VerticalAlignment", -4108);

    i++;
    delete borders;
    delete col;
    delete range;
}
