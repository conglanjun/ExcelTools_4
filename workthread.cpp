#include "workthread.h"
#include "qt_windows.h"
#include <QDebug>

void setCellValue(QAxObject *work_sheet, int row, int column, QString data)
{
    QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", row, column);
    if(row == 12){
        if(column == 5 || column == 6){
            QAxObject* interior = cell->querySubObject("Interior");
            interior->setProperty("Color", QColor(255, 255, 0));   //设置单元格背景色（黄色）
            QAxObject *font = cell->querySubObject("Font");
            font->setProperty("Color", QColor(0, 0, 0));
        }else{
            QAxObject* interior = cell->querySubObject("Interior");
            interior->setProperty("Color", QColor(0, 0, 255));   //设置单元格背景色（绿色）
            QAxObject *font = cell->querySubObject("Font");
            font->setProperty("Color", QColor(255, 255, 255));
        }
    }
    // 寄件公司地址 收件公司地址 两个成本中心
    cell->setProperty("Value", data);  //设置单元格值
    if(row == 12){
        if(column == 1 || column == 3){
            cell->setProperty("ColumnWidth", 60);
        }else if (column == 2 || column == 4 || column == 5 || column == 6){
            cell->setProperty("ColumnWidth", 16);
        }
    }
}


void WorkThread::run ()
{
    CoInitializeEx(NULL, COINIT_MULTITHREADED);

    QString path = getPath();//得到用户选择的文件名
    QAxObject * excel = new QAxObject("Excel.Application");
//        excel.setProperty("Visible", false);
    QAxObject *work_books = excel->querySubObject("WorkBooks");
    work_books->dynamicCall("Open(const QString&)", path);
    excel->setProperty("Caption", "Qt Excel");
    QAxObject *work_book = excel->querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");  //Sheets也可换用WorkSheets
    //删除工作表（删除第一个）
//        QAxObject *first_sheet = work_sheets->querySubObject("Item(int)", 1);
//        first_sheet->dynamicCall("delete");

    // 原始账单
    QAxObject *data_sheet = work_sheets->querySubObject("Item(int)", 1);
    // 人名和成本中心对应
    QAxObject *person_center_sheet = work_sheets->querySubObject("Item(int)", 2);
    // 新旧成本中心替换
//    QAxObject *new_old_sheet = work_sheets->querySubObject("Item(int)", 3);
    // 计算发票号个数
    // 排序后的发票号
    int code_row = 12;// 记录数据行数，为了显示进度条
    int code_column = 23;
    while(true)
    {
        QAxObject *send_address_cell = data_sheet->querySubObject("Cells(int,int)", code_row, code_column);
        QString send_address_str = send_address_cell->dynamicCall("Value2()").toString();
        if(send_address_str.isEmpty()) break;
        ++code_row;
    }
    emit send_excel_row_count(code_row);
    //插入工作表（插入至最后一行）
    int sheet_count = work_sheets->property("Count").toInt();
    QAxObject *last_sheet = work_sheets->querySubObject("Item(int)", sheet_count);
    QAxObject *work_sheet = work_sheets->querySubObject("Add(QVariant)", last_sheet->asVariant());
    last_sheet->dynamicCall("Move(QVariant)", work_sheet->asVariant());

    work_sheet->setProperty("Name", "制作后发客户");  //设置工作表名称

    // 新旧成本中心替换
//    QVector<QString> oldList;
//    QVector<QString> newList;
//    int new_old_row = 2;
//    int old_column = 1;
//    int new_column = 2;
//    while(true){
//        QString old_str = data_sheet->querySubObject("Cells(int,int)", new_old_row, old_column)->dynamicCall("Value2()").toString();
//        QString new_str = data_sheet->querySubObject("Cells(int,int)", new_old_row, new_column)->dynamicCall("Value2()").toString();
//        if(old_str.isEmpty() || new_str.isEmpty()){
//            break;
//        }
//        oldList.append(old_str);
//        newList.append(new_str);
//    }
    // 记录word sheet中的行数。
    int word_row = 12;
    int word_column = 1;
    int username_column = 12;// 经手人
    // 原始账单 寄件公司地址
    int data_row = 12;
    int send_address = 23;
    int person_center = 1;
    int center_code = 2;
    // 收件公司地址
    int recieve_address = 32;
    while(true)
    {
        // 获取data中的值。
        QAxObject *send_address_cell = data_sheet->querySubObject("Cells(int,int)", data_row, send_address);
        QString send_address_str = send_address_cell->dynamicCall("Value2()").toString();
        QAxObject *recieve_address_cell = data_sheet->querySubObject("Cells(int,int)", data_row, recieve_address);
        QString recieve_address_str = recieve_address_cell->dynamicCall("Value2()").toString();
        if(send_address_str.isEmpty() || recieve_address_str.isEmpty()) break;

        // set 寄件地址值到work
        setCellValue(work_sheet, word_row, word_column, send_address_str);
        // set 收件地址到work
        setCellValue(work_sheet, word_row, word_column + 2, recieve_address_str);

        // 找出地址连续六位数字，排除电话号码
        QRegExp rx("^((\\w*\\D+\\d{6})|(\\w*\\D+\\d{6}\\D+\\w*)|(\\d{6}\\D+\\w*))$");
        bool send_flag = rx.exactMatch(send_address_str.remove(QRegExp("\\s")));
        QRegExp rx_index("\\d{6}");
        int send_index = rx_index.indexIn(send_address_str);
        bool recieve_flag = rx.exactMatch(recieve_address_str.remove(QRegExp("\\s")));
        int recieve_index = rx_index.indexIn(recieve_address_str);
        QString send_number = "";
        QString recieve_number = "";
        if(send_flag){
            send_number = send_address_str.mid(send_index, 6);
        }
        if(recieve_flag){
            recieve_number = recieve_address_str.mid(recieve_index, 6);
        }

        // set 成本中心
        if(data_row == 12){
            setCellValue(work_sheet, word_row, word_column + 1, "寄件地址成本中心");
            setCellValue(work_sheet, word_row, word_column + 3, "收件地址成本中心");
        }else if (data_row > 12){
            if(!send_number.isEmpty()){
                setCellValue(work_sheet, word_row, word_column + 1, send_number);
            }
            if(!recieve_number.isEmpty()){
                setCellValue(work_sheet, word_row, word_column + 3, recieve_number);
            }
        }

        // 没有匹配上需要去 经手人 成本中心关系表中查找
        if(data_row == 12){
            setCellValue(work_sheet, word_row, word_column + 4, "人员v成本中心");
        }else if (data_row > 12){
            if(!send_flag && !recieve_flag){// 去v表查
                // 经手人中文名
                QAxObject *username_cell = data_sheet->querySubObject("Cells(int,int)", data_row, username_column);
                QString username_str = username_cell->dynamicCall("Value2()").toString();
                int person_row = 2;
                while(true){
                    QAxObject *person_center_cell = person_center_sheet->querySubObject("Cells(int,int)", person_row, person_center);
                    QString person_center_str = person_center_cell->dynamicCall("Value2()").toString();
                    if(username_str.compare(person_center_str) == 0){// 找到人对应的成本号
                        QAxObject *center_code_cell = person_center_sheet->querySubObject("Cells(int,int)", person_row, center_code);
                        QString center_code_str = center_code_cell->dynamicCall("Value2()").toString();
                        setCellValue(work_sheet, word_row, word_column + 4, center_code_str);
                        break;
                    }
                    if(person_center_str.isEmpty()){
                        break;
                    }
                    ++ person_row;
                }
            }
        }

        // 需要替换旧码成为新码
//        if(data_row == 12){
//            setCellValue(work_sheet, word_row, word_column + 5, "新旧码替换");
//        }else if (data_row > 12){

//        }


        ++ data_row;
        ++word_row;
        emit send_excel_row_done();
    }
    work_book->dynamicCall("Save()");  //保存文件（为了对比test与下面的test2文件，这里不做保存操作） work_book->dynamicCall("SaveAs(const QString&)", "E:\\test2.xlsx");  //另存为另一个文件
    work_book->dynamicCall("Close(Boolean)", false);  //关闭文件
    excel->dynamicCall("Quit(void)");  //退出
    emit send_export_signal(m_path);
//    send();

//    delete excel;

//    QMessageBox::information(getExcel(), tr("Information"), "完事啦，去：" + path + "，查看文件吧！");
}

//void WorkThread::send()
//{
//    emit send_export_signal();
//}

