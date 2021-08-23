#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include "input.h"

#include <QMainWindow>
#include <QLineEdit>

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = 0);
    ~MainWindow();

private:
    Input *input;

};

#endif // MAINWINDOW_H
