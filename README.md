# Справка для расчетов за выполненные работы. Форма ЭСМ-7

## Даталогическая модель
<p align="center">
   <img width="468" height="426" alt="image" src="https://github.com/user-attachments/assets/afcc7519-5f4a-4378-bd17-8b6898b92d9d" />
</p>

## Блок-схема
<p align="center">
   <img width="297" height="495" alt="image" src="https://github.com/user-attachments/assets/a32e3155-79cb-4179-b7f7-6dd686559458" />
</p>
1. Программа начинается с ввода реквизитов документа.
2. Далее определяется количество оказанных услуг, после чего перечисляются данные по каждой услуге, данные записываются в массив.
3. Вводится информация о простоях по вине заказчика.
4. И наконец, программа суммирует стоимость каждой услуги, прибавляет налог NDS и выводит итоговую сумму SUM.

## Диаграмма компонентов (component diagram)
<p align="center">
   <img width="430" height="124" alt="image" src="https://github.com/user-attachments/assets/84f1e580-b524-46aa-b17a-12a9253c7343" />
</p>

## Use Case диаграмма
<p align="center">
   <img width="375" height="412" alt="image" src="https://github.com/user-attachments/assets/ddd250e8-9304-491f-9d14-6c950189027d" />
</p>

## Описание
Работа представляет из себя программу-приложение с GUI, с помощью которой составляются справки ЭСМ-7 (справка для расчетов за выполненные услуги). В подписании документа принимают участие две стороны: исполнитель и заказчик. Справка формируется на стороне исполнителя.

### Структура
1. Основной файл с программой main.py (написан на Python с использованием библиотеки Tkinter)
2. Библиотека «Справочники» содержит в себе файлы csv:
   - 5 справочников для работы приложения
   - Файл с сохраненными настройками программы setup.csv
3. Файл-шаблон template.xlsx для экспорта готовых справок
<p align="center">
   <img width="180" height="187" alt="image" src="https://github.com/user-attachments/assets/c84fb68b-7763-4b23-9f5e-9710f6a41af8" />
</p>

## Приложение

При запуске кода открывается приложение.

### Страница 1:
<p align="center">
   <img width="468" height="231" alt="image" src="https://github.com/user-attachments/assets/4c66f60d-9bc3-47ad-89c6-83e5b45259c3" />
</p>

На странице 1:
1. Часть данных (данные об организации-исполнителе) заполнена автоматически из файла setup.csv
2. Дата составления документа заполняется автоматически сегодняшняя
3. Остальные данные вводятся с клавиатуры («ручной ввод»)
4. Поля «Машина» и «Машинисты» выбираются из выпадающего списка
После заполнения всех полей на странице 1, пользователь нажимает кнопку «Далее» и переходит на страницу 2

### Страница 2:
<p align="center">
   <img width="468" height="229" alt="image" src="https://github.com/user-attachments/assets/18a3f7b4-c743-4c11-9bf8-9af70689d264" />
</p>

На странице 2:
1. Секция добавления услуг с кнопкой «Добавить»
2. Секция добавление простоя с кнопкой «Добавить простои»
3. Таблица отображения результатов
4. Кнопка «Экспорт» - экспортирует справку и завершает программу
5. Поля «Код и наименование работы» и «Причина простоя» выбираются из выпадающего списка

Пользователь добавляет услуги, стоимость и время, а также простои по вине заказчика. После чего по кнопке «Экспорт» производится экспорт готовой справки в формате xlsx, программа завершается. 

### Выпадающие списки (Combobox)
1. Поля «Машина» и «Машинисты» на странице 1 заполняются выбором соответствующей ячейки из выпадающего списка (Combobox). Данные для выбора взяты из справочников mashines.csv и stuff.csv
<p align="center">
   <img width="199" height="63" alt="image" src="https://github.com/user-attachments/assets/59a5065a-c7ed-41af-a8e4-0d8ce70994ce" />  <img width="200" height="102" alt="image" src="https://github.com/user-attachments/assets/59c64aac-8fdb-48f3-8ad8-677a79d0310c" />
</p>

3. Поля «Код и наименование работы» и «Причина простоя» на странице 2. Данные для выбора взяты из справочников works.csv и downtime.csv
<p align="center">
   <img width="188" height="86" alt="image" src="https://github.com/user-attachments/assets/481c57e4-5f3c-4bce-a44e-ffccfc720ec1" />  <img width="242" height="77" alt="image" src="https://github.com/user-attachments/assets/64252c23-aaa5-4c65-aace-85212c1c984f" />
</p>

### CSV-файлы
1. Справочник numbers.csv нужен для конвертации чисел из цифр в пропись. Количество часов работы
<p align="center">
   <img width="94" height="266" alt="image" src="https://github.com/user-attachments/assets/f49d9995-8326-4c42-a9fc-950300615ff1" />
</p>

2. Справочник mahines.csv: содержит информацию о машинах организации-исполнителя: наименование, марка, номер
<p align="center">
   <img width="362" height="89" alt="image" src="https://github.com/user-attachments/assets/e391f39a-9ba6-43c6-89a5-189429472c9a" />
</p>

3. Справочник stuff.csv: машинисты организации-исполнителя
<p align="center">
   <img width="149" height="164" alt="image" src="https://github.com/user-attachments/assets/334834bd-97e4-4b35-9958-e82171a87a4a" />
</p>

4. Справочник works.csv: коды и виды работ
<p align="center">
   <img width="295" height="267" alt="image" src="https://github.com/user-attachments/assets/a01fc085-8719-46b9-8ab4-fa5b45387140" />
</p>

5. Справочник downtime.csv: коды и виды простоев
<p align="center">
   <img width="298" height="115" alt="image" src="https://github.com/user-attachments/assets/9853d696-c686-45b7-8e7d-485205ba4f82" />
</p>

6. Файл с настройками программы: данные об организации-исполнителе
<p align="center">
   <img width="468" height="15" alt="image" src="https://github.com/user-attachments/assets/396d962f-e56e-45d9-af59-97141827e34b" />
</p>

7. Файл-шаблон xlsx
<p align="center">
   <img width="454" height="265" alt="image" src="https://github.com/user-attachments/assets/45535bfe-b86b-4478-9e37-9b675c2838fa" />
</p>

### Форматно-логический контроль ввода
1. Поля на обеих страницах, которые состоят только из цифр, не реагируют на ввод иных символов с клавиатуры
2. Все поля должны на странице 1 должны быть заполнены. В ином случае выводится сообщение:
<p align="center">
   <img width="149" height="103" alt="image" src="https://github.com/user-attachments/assets/bb7f9723-5e73-49b8-95e6-a9a1f5f5fba3" />
</p>

3. Поля, где нужно ввести дату, не реагируют на ввод символов, кроме цифр и точки. Даты вводятся в определенном формате. В ином случае выводится сообщение:
<p align="center">
   <img width="141" height="108" alt="image" src="https://github.com/user-attachments/assets/b5325fe5-1800-4e94-a798-7d4be753e839" />
</p>

4. На странице 2 нужно добавить хотя бы одну услугу, чтобы экспортировать справку
<p align="center">
   <img width="152" height="107" alt="image" src="https://github.com/user-attachments/assets/e3ac6e79-450d-4510-8a63-1e37f83c39d2" />
</p>

### Тестирование программы
#### Вводные данные:
##### Страница 1:
<p align="center">
   <img width="852" height="423" alt="image" src="https://github.com/user-attachments/assets/b935f19f-9bf1-44d8-9f54-ff0bfe55b558" />
</p>

##### Страница 2:
<p align="center">
   <img width="858" height="425" alt="image" src="https://github.com/user-attachments/assets/a42087bd-70fa-47fb-a768-1cd21f33ddc5" />
</p>
 
## Итоговая справка:
<p align="center">
   <img width="1468" height="793" alt="image" src="https://github.com/user-attachments/assets/67ef316d-af06-428d-be26-006cc42732e2" />
</p>
