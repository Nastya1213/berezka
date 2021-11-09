import sys
from PyQt5 import uic  # импортируем класс uic
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5 import QtWidgets
from unicodedata import normalize
from openpyxl import load_workbook
from random import shuffle


class Orthoepy(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(Orthoepy, self).__init__(parent)
        uic.loadUi('dialog.ui', self)

        self.users_answers = []
        self.temporary_lst = []
        self.right_answers = ['мусоропрово́д', 'не ровён час','предвосхи́тить', 'обеспе́чение',
                              'балова́ть (балу́ю, избало́ванный)', 'диспансе́р', 'о́бнял', 'то́рты',
                              'ша́рфы', 'кварта́л', 'облегчи́ть', 'зави́дно', 'ла́тте', 'свёкла',
                              'граффи́ти']
        self.wrong_answers = ['Мусоропро́вод', 'Не ро́вен час', 'Предвосхити́ть', 'Обеспече́ние',
                              'Ба́ловать (ба́лую, изба́лованный)', 'Диспа́нсер', 'обня́л', 'торты́',
                              'Шарфы́', 'Ква́ртал', 'Обле́гчить','За́видно', 'Латте́', 'Свекла́',
                              'Гра́ффити']


        self.pushButton.clicked.connect(self.run)
        self.rb1.toggled.connect(self.answer_user)
        self.rb2.toggled.connect(self.answer_user)
        self.rb3.toggled.connect(self.answer_user)
        self.rb3.setChecked(True)

    def answer_user(self):
        self.temporary_lst.append(self.sender().text())

    def run(self):
        self.options.append(self.right_answers[self.count])
        self.options.append(self.wrong_answers[self.count])
        shuffle(self.options)
        self.rb1.setText(self.options[0])
        self.rb2.setText(self.options[1])
        self.options.clear()
        self.users_answers.append(self.temporary_lst[-1])
        print(self.users_answers)




class ClssDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        self.wb = load_workbook('words.xlsx')
        self.sheet = self.wb['Лист1']
        self.right_answers = []
        self.wrong_answers = []
        self.explanation = []
        # получаем правильные ответы из столбца B
        for i in range(2, 64):
            self.a = self.sheet[f'B{i}'].value
            self.clean_a = normalize("NFKD", self.a)
            self.right_answers.append(self.clean_a)
        # получаем неверные отыеты из столбца C
        for i in range(2, 64):
            self.a = self.sheet[f'C{i}'].value
            self.clean_a = normalize("NFKD", self.a)
            self.wrong_answers.append(self.clean_a)
        # получаем поячнение из столбца D
        for i in range(2, 64):
            self.a = self.sheet[f'D{i}'].value
            self.clean_a = normalize("NFKD", self.a)
            self.explanation.append(self.clean_a)

        self.options = []
        self.count = 0
        self.right_count = 0
        self.users_answers = []
        self.temporary_lst = []

        super(ClssDialog, self).__init__()
        uic.loadUi('dialog.ui', self)

        self.options.append(self.right_answers[self.count])
        self.options.append(self.wrong_answers[self.count])
        shuffle(self.options)
        self.rb1.setText(self.options[0])
        self.rb2.setText(self.options[1])
        self.count +=1
        self.options.clear()

        self.pushButton.clicked.connect(self.run)
        self.rb1.toggled.connect(self.answer_user)
        self.rb2.toggled.connect(self.answer_user)
        self.rb3.toggled.connect(self.answer_user)
        self.rb3.setChecked(True)

    def answer_user(self):
        self.temporary_lst.append(self.sender().text())

    def run(self):

        if len(self.users_answers) == 58:
            self.pushButton.setText('Завершить и узнать результат!')
            self.options.append(self.right_answers[self.count])
            self.options.append(self.wrong_answers[self.count])
            shuffle(self.options)
            self.rb1.setText(self.options[0])
            self.rb2.setText(self.options[1])
            self.options.clear()
            self.count += 1
            #print(self.temporary_lst[-1])
            self.users_answers.append(self.temporary_lst[-1])
            print(self.users_answers)
            #print(self.count)
            #self.pushButton.setText('Завершить и узнать результат!')
        # если все все вопросы решены
        if len(self.users_answers) == 60:
            #пробегаем по ответам пользователя
            for i in self.users_answers:
                #если ответ пользователя есть в правильных ответах,
                # то прибавляем 1 к счетчику
                if i in self.right_answers:
                    self.right_count += 1
                #если такого ответа нет, выводим правильный ответ
                else:
                    print(self.right_answers[self.users_answers.index(i)])
            #выводим количество правильных ответов
            print(self.right_count)
            #закрываем диалоговое окно
            self.dialog.close()
        #остались ещё вопросы
        else:
            self.options.append(self.right_answers[self.count])
            self.options.append(self.wrong_answers[self.count])
            shuffle(self.options)
            self.rb1.setText(self.options[0])
            self.rb2.setText(self.options[1])
            self.options.clear()

            self.count += 1
            self.users_answers.append(self.temporary_lst[-1])
            self.rb3.setChecked(True)


class MyWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        self.runner = False
        uic.loadUi('design.ui', self)  # Загружаем дизайн
        self.test_60.clicked.connect(self.run_test)
        #self.test_orthoepy.clicked.connect(self.run_test_orthoepy)

    def run_test(self):
        #ex.close()
        self.runner = True
        test = ClssDialog(self)
        test.exec_()

    def except_hook(cls, exception, traceback):
        """Функция для отслеживания ошибок PyQt5"""
        sys.__excepthook__(cls, exception, traceback)


    """
    def run_test_orthoepy(self):
        self.runner = True
        test_orthoepy = Orthoepy(self)
        test_orthoepy.exec_()
    """


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWidget()
    ex.show()
    sys.exit(app.exec_())
    sys.excepthook = except_hook
    sys.exit(app.exec_())