# Chart_Gantt
Диаграмма Ганта — это график работ по проекту, который строится в виде таблицы с этапами и ответственными за их выполнение.
Данная программа автоматически заполняет Excel-файл на основе результата sql-запроса базы данных

Необходимый пакет для корректной работы приложения :
Среда выполнения .NET 8.0.7 (https://dotnet.microsoft.com/ru-ru/download/dotnet/8.0)

Путь к SQL-скрипту, из которого берётся выборка - "~\bin\Debug\net8.0\ОтчётLotuscvb.sql" (Кодировка UTF-8)
Excel-файл для вывода - "~\\Исходник\\Отчёт_Gant.xlsx"

*Образец принимаемых данных*:
| Название выполненных работ | Длительность выполнения работ | Время начала работ | Время окончания работ |
|-----------------------------|-------------------------------|--------------------|-----------------------|
| Проверка серверов           | 2 часа                        | 17:00              | 19:00                 |
| Обновление ПО               | 1.5 часа                      | 19:30              | 21:00                 |
| Резервное копирование данных| 3 часа                        | 21:30              | 00:30                 |
| Тестирование системы        | 2 часа                        | 01:00              | 03:00                 |
| Настройка сети              | 1 час                         | 03:30              | 04:30                 |
| Мониторинг безопасности     | 2 часа                        | 05:00              | 07:00                 |
| Оптимизация базы данных     | 1.5 часа                      | 17:00              | 18:30                 |
| Установка обновлений        | 2 часа                        | 19:00              | 21:00                 |
| Проверка логов              | 1 час                         | 21:30              | 22:30                 |
| Техническое обслуживание    | 2 часа                        | 23:00              | 01:00                 |

По окончанию выполнения отобразится сообщение "Отчёт Lotus заполен."
При возникновении ошибок пользователю будет отображено сообщение "Отчёт Lotus не заполнен. Необходимо заполнить вручную" и отчёт НЕ БУДЕТ ЗАПОЛНЕН!!!