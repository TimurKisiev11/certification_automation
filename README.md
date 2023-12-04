# certification_automation
Проект, призванный автоматизировать процесс аттестации чаптера инженеров данных и разработчиков.<br>
Процесс аттестации, который автоматически проводиться (пункты 3 и 4) путем запуска из командной строки скрипта **certification_script.py** заключается в следующем: </br>
#### 1. Аттестуемый заполняет таблицу с некоторым набором критериев, самостоятельно оценивая себя по 5-балльной шкале.
#### 2. Затем Тимлид заполняет таблицу своими оценками и общая оценка по каждому критерию считается как среднее арифметическое двух оценок (собственной и тимлида).

![image](https://github.com/TimurKisiev11/certification_automation/assets/113093142/20e501b8-401f-4415-b773-fa154187d745)

#### 3. Далее полученные оценки нужно сравнить с целевыми значениями для каждого профессионального уровня, если более 80% критериев проходят, то уровень считается присвоенным.

![image](https://github.com/TimurKisiev11/certification_automation/assets/113093142/6ec2a8e4-0f2f-4920-8e1c-9d69931fb971)

#### 4. По каждому аттестуемому создается новая таблица с результатами тестирования, которая выглядит вот так:

![image](https://github.com/TimurKisiev11/certification_automation/assets/113093142/3fe4f75c-6051-4e5a-9b60-f86478ae587f)

## Как пользоваться скриптом **certification_script.py**?

#### 1. Для начала нужно скачать скрипт, скачивать и запускать весь проект в IDE необязательно.
#### 2. Далее нужно собрать в одну директорию собранные в процессе тестирования файлы с оценками (свой для каждого аттестуемого)

  ![image](https://github.com/TimurKisiev11/certification_automation/assets/113093142/895342f4-9702-42eb-9602-f272e9f0bc7d)
  
#### 3. Также желательно создать отдельную директорию для результирующих файлов.
#### 4. Теперь запускаем командную строку (от имени администратора)

![image](https://github.com/TimurKisiev11/certification_automation/assets/113093142/429f2660-d785-414d-9098-fcfe80ff9909)

#### 5. И поочередно указываем пути:
1. Путь к самому скрипту.
2. Путь к директории с xlsx файлами, в которых содержатся оценки (Оценка_DE_сам.оценка+TL оценка.xlsx).
3. Путь к xlxs файлу в котором содержатся целевые значения для проведения сертификации (Компетенции_по_шкале_DE.xlsx).
4. Путь к директории, в которую будут сохранятся результирующие xlsx файлы по одному на каждый файл с оценками.
   ![image](https://github.com/TimurKisiev11/certification_automation/assets/113093142/44d6e643-d9ae-41cc-b34f-8e121615d27c)
и нажимаем Enter
5. В директории, указанной ранее, появятся результирующие файлы.
   ![image](https://github.com/TimurKisiev11/certification_automation/assets/113093142/477398c8-f326-4b43-8dbc-36d1212f49c5)
