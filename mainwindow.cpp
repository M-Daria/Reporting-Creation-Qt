#include "mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    input = new Input;
    input->setMinimumHeight(input->height() - 150);
    input->setMinimumWidth(700);

    setCentralWidget(input);
    setWindowTitle(tr("Заполнение отчётности"));
}

MainWindow::~MainWindow()
{
}
