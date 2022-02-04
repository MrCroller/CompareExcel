# Программа для объединения файлов Excel :capital_abcd:
### Соединение всех Exel таблиц по месяцам :shipit:

Запуск осуществляется через cmd со следующими параматерами:

>- `-p` :inbox_tray: Папка для чтения
>- `-o` :outbox_tray: Папка вывода объединенного файла
>- `-i` :speech_balloon: Уровень информирования, выбор из 3 доступных:
>    - `None` - Без консольного вывода, 
>    - `Main` - главные процессы [по умолчанию], 
>    - `All` - Полный вывод, аж на каждую ячеечку.


##### Пути к папкам в PowerShell окружать символом `` ` ``. Пример:
```
.\UseExcel.exe -p `C:\Users\Username\Downloads\resources` -o `C:\Users\Username\Downloads\resources two\chek 1`
```
