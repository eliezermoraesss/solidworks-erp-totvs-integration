def engenharia_qss():
    return """* {
                background-color: #363636;
            }

            QLabel, QCheckBox {
                color: #EEEEEE;
                font-size: 11px;
                font-weight: bold;
            }
            
            QLabel#logo-enaplic {
                margin: 5px 0;
            }
            
            QLabel#label-line-number {
                font-size: 16px;
                font-weight: normal;
            }

            QLineEdit {
                background-color: #DFE0E2;
                padding: 5px;
                border-radius: 8px;
            }
            
            QDateEdit, QComboBox {
                background-color: #DFE0E2;
                border: 1px solid #262626;
                padding: 5px 10px;
                border-radius: 10px;
                height: 20px;
                font-size: 16px;
            }
            
            QComboBox QAbstractItemView {
                background-color: #EEEEEE;
                color: #000000; /* Cor do texto para garantir legibilidade */
                selection-background-color: #0a79f8; /* Cor de seleção quando passa o mouse */
                selection-color: #FFFFFF; /* Cor do texto quando selecionado */
                border: 1px solid #393E46;
            }
    
            QDateEdit::drop-down, QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 30px;
                border-left-width: 1px;
                border-left-color: darkgray;
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
            }
    
            QDateEdit::down-arrow, QComboBox::down-arrow {
                image: url(../resources/images/arrow.png);
                width: 10px;
                height: 10px;
            }   

            QPushButton {
                background-color: #0a79f8;
                color: #fff;
                padding: 5px 15px;
                border: 2px;
                border-radius: 8px;
                font-size: 11px;
                height: 20px;
                font-weight: bold;
                margin: 10px 5px;
            }

            QPushButton#btn_home {
                background-color: #c1121f;
            }

            QPushButton:hover, QPushButton#btn_home:hover {
                background-color: #fff;
                color: #0a79f8
            }

            QPushButton:pressed, QPushButton#btn_home:pressed{
                background-color: #6703c5;
                color: #fff;
            }

            QTableWidget {
                border: 1px solid #000000;
                background-color: #363636;
                padding-left: 10px;
                margin-bottom: 15px;
            }

            QTableWidget QHeaderView::section {
                background-color: #262626;
                color: #A7A6A6;
                padding: 5px;
                height: 18px;
            }

            QTableWidget QHeaderView::section:horizontal {
                border-top: 1px solid #333;
            }

            QTableWidget::item {
                /* background-color: #363636; */
                color: #f9feff;
                font-weight: bold;
                padding-right: 8px;
                padding-left: 8px;
            }

            QTableWidget::item:selected {
                background-color: #000000;
                color: #f9feff;
                font-weight: bold;
            }"""
