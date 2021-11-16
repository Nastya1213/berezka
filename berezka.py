import sys
from PyQt5 import uic  # импортируем класс uic
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5 import QtWidgets
from unicodedata import normalize
from openpyxl import load_workbook
from random import shuffle

#Извлекаем и таблицы Excel нужные данные для теста "60 частых ошибок"
wb = load_workbook('words.xlsx')
sheet = wb['Лист1']
right_answers_60 = [] #правильные ответы на тест "60 частых ошибок"
wrong_answers_60 = [] #неправильные ответы на тест "60 частых ошибок"
explanation_60 = [] #пояснения для теста "60 частых ошибок"

#Функция для извлечения данных из столбца и записи в список
#1-й аргумент - это буквенный номер столбца
#2-й аргумент - это в какой список нужно поместить
#3-й и 4-й аргументы - это с какой по счету строки начинаются данные и какой заканчиваются
def answr(column, what_list, start, finish):
    for i in range(start, finish): #цикл чтобы пройти от начала и до конца
        a = sheet[f'{column}{i}'].value # Sheet у нас это лист, из него мы получаем значение в столбце,
        # который мы указали в аргкменте, по номеру i
        what_list.append(normalize("NFKD", a))#в таблице, имелись неразрывные пробелы,
        # которые в списке выглядили как \xa0


answr('B', right_answers_60, 2, 64)
answr('C', wrong_answers_60, 2, 64)
answr('D', explanation_60, 2, 64)

#Извлекаем и таблицы Excel нужные данные для теста "Орфоэпия"
tab = load_workbook('Orthoepy.xlsx')
sheet = tab['Лист1']
right_answers_ort = []
wrong_answers_ort = []
answr('A', right_answers_ort, 1, 15)
answr('B', wrong_answers_ort, 1, 15)


class Orthoepy(QtWidgets.QDialog):
    def __init__(self):
        super(Orthoepy, self).__init__()
        #загружаем дизайн
        uic.loadUi('ort_dialog.ui', self)

        self.users_answers = [] #список для ответов пользователя
        self.temporary_lst = [] #времменный список, чтобы в users_answers вносить не каждый ответ,
        # который,еще думая, выбирал пользователь, а последний его ответ перед нажатием
        # на кнопку (pushButton) 'следующий вопрос'
        self.right_answers = right_answers_ort
        self.wrong_answers = wrong_answers_ort
        self.options = [self.right_answers[0], self.wrong_answers[0]]
        shuffle(self.options)
        self.rb1.setText(self.options[0])
        self.rb2.setText(self.options[1])
        self.rb3.setText('Пропустить')
        self.options.clear()

        self.pushButton.clicked.connect(self.run)
        self.rb1.toggled.connect(self.answer_user)
        self.rb2.toggled.connect(self.answer_user)
        self.rb3.toggled.connect(self.answer_user)
        self.rb3.setChecked(True)

    def answer_user(self):
        self.temporary_lst.append(self.sender().text())

    def run(self):
        self.options.append(self.right_answers[len(self.users_answers)])
        self.options.append(self.wrong_answers[len(self.users_answers)])
        shuffle(self.options)
        self.rb1.setText(self.options[0])
        self.rb2.setText(self.options[1])
        self.options.clear()
        self.users_answers.append(self.temporary_lst[-1])
        print(self.users_answers)


class ClssDialog(QtWidgets.QDialog):
    def __init__(self):
        self.right_answers = right_answers_60
        self.wrong_answers = wrong_answers_60
        self.explanation = explanation_60
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
        self.count += 1
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
            self.users_answers.append(self.temporary_lst[-1])
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
                    print(self.right_answers[i.index])
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
        uic.loadUi('design.ui', self)  # Загружаем дизайн
        self.test_60.clicked.connect(self.run_test_60)
        self.test_orthoepy.clicked.connect(self.run_test_orthoepy)

    def run_test_60(self):
        test = ClssDialog()
        test.exec_()

    def run_test_orthoepy(self):
        test = Orthoepy()
        test.exec_()


def except_hook(cls, exception, traceback):
    """Функция для отслеживания ошибок PyQt5"""
    sys.__excepthook__(cls, exception, traceback)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWidget()
    ex.show()
    sys.exit(app.exec_())
    sys.excepthook = except_hook
    sys.exit(app.exec_())
