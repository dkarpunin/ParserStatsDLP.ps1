# ParserStatsDLP.ps1
Парсинг файла статистики DLP Solar Dozor и сохранение результата в CSV

Коннектор db-conn-3

1. Запрашиваем статистику коннектора db-conn-3 и сохраняем её в файл db-conn-3.txt:
  /opt/dozor/bin/shell show-stats -2 -con db-conn-3 > db-conn-3.txt
2. Помещаем db-conn-3.txt в папку с файлом ParseStatsDLP.ps1
3. Запускаем парсер:
  powershell ParseStatsDLP.ps1
4. На выходе получаем файл db-conn-3.csv
