### Как подключить пример
1. Откройте Excel и перейдите в редактор Visual Basic (в ленте Разработчик > Visual Basic) или комбинация клавиш (Alt+F11)
2. Откройте список ссылок: в меню Tools > References...
3. Отметьте пункты «Microsoft Scripting Runtime» и «Microsoft VBScript Regular Expressions 5.5».
4. Нажмите OK.
5. Откройте модуль Sheet1.
6. Замените значение константы API_KEY на ваш API-ключ.
### Как работает пример
Введите ИНН в соответствующую колонку, нажмите кнопку «Start». В соседних ячейках появятся реквизиты компании из Подсказок. Код перебирает каждую строку пока не дойдет до последней строки с ИНН.
