# Data of project budgets

* Описание проекта
* Используемые технологии
* Запуск

## Описание проекта

Данный проект является сведением трат бюджетов различных производственных проектов на зарплаты и премии сотрудникам за календарный год.
В проекте есть возможность:

* выбрать рассматриваемый календарный год;
* создать файл с 2 листами:
  * лист с пустой таблицей с актуальным списком сотрудников для заполнения зарплат и премий за календарный год;
  * лист с пустой таблицей с актуальным списком производственных проектов для заполнения бюджетов на начало года и распределения выделений премий;
* обновить файл, актуализировав список сотрудников и проектов;
* сгенерировать итоговый файл данных, в котором будут отражены:
  * траты бюджетов проектов на зарплаты и премии сотрудникам по каждому месяцу года;
  * сумма трат бюджетов проектов на зарплаты за год (в виде формулы excel);
  * сумма трат бюджетов проектов на премии за год (в виде формулы excel);
  * остаток бюджетов проектов на конец календарного года

## Используемые технологии

|Библиотеки|
|:---|
| Python 3.11.1   |  
| Openpyxl 3.1.5  |

## Запуск

Для начала необходимо сгенерировать файл выгрузки "Отчёт за текущий год.xlsx" особого вида и положить его в корневую директорию.  
Остаётся запустить файл main.py, находящийся в корневой директории.
