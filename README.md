# kai-software-economics

Розрахункова модель до лабораторних робіт з дисципліни «Економіка програмного забезпечення», варіант 14: «ЕнергоПульс», моніторинг енергоспоживання будівлі (PropTech). Валюта моделі: EUR.

## Структура

* `econ_model.py`: єдина модель для ЛР № 1...5 (параметри, сегменти, стейкхолдери, scope MVP, витрати, реєстр припущень).
* `lab1/build_lab1.py`: генерація робочої книги з формулами, перевірка перерахунку в LibreOffice, графіки та схеми.
* `lab1/out`: результати (xlsx, png).
* `lab2/build_lab2.py`: ЛР № 2, WBS, CPM, PERT, симуляція Монте-Карло, backlog і velocity, UCP, ресурсний календар, baseline; робоча книга з формулами і її перевірка в LibreOffice.
* `lab2/out`: результати ЛР № 2 (xlsx, png, dot).
* `lab3/build_lab3.py`: ЛР № 3, склад loaded rates, labor cost за фазами, cost baseline MVP і contingency з двох ризиків, cash burn M1...M12, TCO3 cloud і fixed hosting, чутливість; робоча книга з формулами і її перевірка в LibreOffice.
* `lab3/out`: результати ЛР № 3 (xlsx, png).

## Запуск

```
pip install -r requirements.txt
python3 lab1/build_lab1.py
python3 lab2/build_lab2.py
python3 lab3/build_lab3.py
```

Потрібні LibreOffice (шлях задається змінною SOFFICE) і Graphviz (`dot`).
