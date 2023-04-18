import pandas as pd
import openpyxl
from openpyxl import load_workbook
import sys
import time

class try_toconnect():
    try:
        global wb
        global my_sheet
        global df
        df = pd.read_excel('./DataBase.xlsx')
        wb = openpyxl.load_workbook('./DataBase.xlsx')
        my_sheet = wb.active
    except:
        print('Неудалось подключиться к БД')
        time.sleep(5)
        sys.exit()
def main():
    print ('Сортировка таблицы: 1. По убыванию\n2. По возрастанию\n3. Нынешняя база данных\n4. Количество значений\n5. Имя по алфавиту\n6. Сначала...\n7. Назад')
    value = int(input("Введите значение: "))
    if value == 1:
        print('Сортировка как?\n1. По году вступления \n2. По убыванию возраста\n3. По убыванию возраста и года вступления\n4. Назад')
        mind = int(input('Введите значение: '))
        if mind == 1:
            print(df.sort_values(["Year_to_join"], ascending=False))
        elif mind == 2:
            print(df.sort_values(["Age"]), ascending=False)
        elif mind == 3:
            print(df.sort_values(["Age","Year_to_join"], ascending=[False, False]))
        else:
            return main()
    elif value == 2:
        print("Сортировка как?\n1. По году вступления\n2. По возрастанию возраста\n3. По возрастанию возраста и года вступления\n4. Назад")
        value1 = int(input('Введите значение: '))
        if value1 == 1:
            print(df.sort_values(["Year_to_join"]))
        elif value1 == 2:
            print(df.sort_values(["Age"]))
        elif value1 == 3:
            print(df.sort_values(["Age","Year_to_join"]))
        else:
            return main()
    elif value == 3:
        print(df)
        return main()
    elif value == 4:
        print('1. Количество обладателей одного бренда смартфона\n2. Количество каких-либо одногорожан\n3. Количество людей одинакового пола\n4. Количество одногодок\n5. Количество обладателей одной OS\n6. Количество одногодок вступителей в беседу\n7. Количество людей схожей национальности\n8. Количество клиентов одного банка\n9. Назад')
        value2 = int(input('Введите значение: '))
        if value2 == 1:
            print(df["Brand of phone"].value_counts())
        elif value2 == 2:
            print(df["City"].value_counts())
        elif value2 == 3:
            print(df["Sex"].value_counts())
        elif value2 == 4:
            print(df["Age"].value_counts())
        elif value2 == 5:
            print(df["OS"].value_counts())
        elif value2 == 6:
            print(df["Year_to_join"].value_counts())
        elif value2 == 7:
            print(df["National"].value_counts())
        elif value2 == 8:
            print(df["Bank"].value_counts())
        else:
            return main()
    elif value == 5:
        print("1. От A до Z\n2. Oт Z до A\n3. Назад")
        value3 = int(input("Введите значение: "))
        if value3 == 1:
            print(df.sort_values(["Name"]))
        elif value3 == 2:
            print(df.sort_values(["Name"], ascending=False))
        else:
            return main()

    elif value == 6:
        def value6():
            print("Сначала что:\n1. Сначала жители России/Казахстана\n2. Сначала Android/iOS\n3. Сначала высокие/низкие\n4. Назад")
            num = int(input("Введите значение: "))
            if num == 1:
                print('1. Сначала жители России\n2. Сначала жители Казахстана')
                numc = int(input("Введите значение: "))
                if numc == 1:
                    print(df.sort_values(["Country"], ascending=False))
                elif numc == 2:
                    print(df.sort_values(["Country"]))
                else:
                    return value6()
            elif num == 2:
                print('1. Сначала Android\n2. Сначала iOS\n3. Назад')
                numo = int(input("Введите значение: "))
                if numo == 1:
                    print(df.sort_values(["OS"]))
                elif numo == 2:
                    print(df.sort_values(["OS"], ascending=False))
                else: 
                    return value6()
            elif num == 3:
                print('1. Сначала высокие\n2. Сначала низкие\n3. Назад')
                numh = int(input("Введите значение"))
                if numh == 1:
                    print(df.sort_values(["Height"]))
                elif numh == 2:
                    print(df.sort_values(["Height"], ascending=False))
                else:
                    return value6()
            else:
                return main()
        value6()
    else:
        return

def redactor():
    print('Что вы хотите сделать?\n1.Редактирование/добавление данных\n2.Удаление строки\n3.Назад')
    main = int(input("Введите значение: "))
    if main == 1:
        print (pd.read_excel("./DataBase.xlsx"))
        data = []
        index = []
        
        Name = str(input('Введите имя и фамилию: '))
        Age = int(input("Введите возраст: "))
        City = str(input("Введите город: "))
        Country = str(input("Введите страну: "))
        Sex = str(input("Введите пол: "))
        Height = int(input("Введите рост: "))
        Grade = int(input("Введите школьный класс(если закончили школу введите 0): "))
        OS = str(input("Введите свою операционную систему на телефоне(iOS/Android): "))
        Brand_of_phone = str(input("Введите марку своего смартфона: "))
        National = str(input("Введите свою национальность: "))
        Birth = str(input("Введите дату своего рождения(пример: 01.01.2021): "))
        Bank = str(input("Введите свой основной банк: "))
        Year_to_join = int(input("Введите год, когда вы вступили в беседу(пример: 2020): "))


        inti = int(input("Укажите номер строки(В случае, если вы хотите отредактировать данные, просто укажите строку с существующими данными): "))
        inti = inti + 2

        print('Вы действительно хотите отредактировать строку ', inti-2, '\n1.Да\n2.Нет')
        amogus = int(input("Введите значение: "))
        if amogus == 1:
            c1 = my_sheet.cell(row = inti, column = 1)
            c1.value = Name
            c2 = my_sheet.cell(row= inti, column = 2)
            c2.value = Age
            c3 = my_sheet.cell(row=inti, column=3)
            c3.value = City
            c4 = my_sheet.cell(row=inti,column=4)
            c4.value = Country
            c5 = my_sheet.cell(row=inti,column=5)
            c5.value = Sex
            c6 = my_sheet.cell(row=inti,column=7)
            c6.value = Height
            c7 = my_sheet.cell(row=inti, column=8)
            c7.value = Grade
            c8 = my_sheet.cell(row=inti, column=9)
            c8.value = OS
            c9 = my_sheet.cell(row=inti,column=10)
            c9.value = Brand_of_phone
            c10 = my_sheet.cell(row=inti, column=11)
            c10.value = National
            c11 = my_sheet.cell(row=inti, column=12)
            c11.value = Birth
            c12 = my_sheet.cell(row=inti, column=13)
            c12.value = Bank
            c13 = my_sheet.cell(row=inti, column=14)
            c13.value = Year_to_join
            c14 = my_sheet.cell(row=inti, column=6)
            c14.value = 'Non admin'
            wb.save("./DataBase.xlsx")
            print(pd.read_excel('./DataBase.xlsx'))
        else:
            return redactor()

    elif main == 2:
        print(df)
        a = int(input("Введите номер строки: "))
        print('Вы уверены, что хотите удалить строку ',a,'?\n1.Да\n2.Нет')
        mind = int(input('Введите значение:'))
        if mind == 1:
            a = a+2
            my_sheet.delete_rows(a)
            wb.save('./DataBase.xlsx')

            print(pd.read_excel('./DataBase.xlsx'))
        else:
            return redactor()
    else:
        return

        
class somain:
    wb = openpyxl.load_workbook('DataBase.xlsx')
    wb.save("./DataBase.xlsx")
    df = pd.read_excel('./DataBase.xlsx')
    print('Добро пожаловать в консольный сортировщик и редактор БД и ДТ.\nЧто вы хотите сделать?\n1.Отсортировать\n2.Редактировать')
    a = int(input("Введите значение: "))
    if a==1:
        main()
    elif a==2:
        redactor()
    else:
        print("Неверное значение")

while True:
    somain()
    try_toconnect()
