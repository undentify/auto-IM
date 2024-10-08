Данный скрипт предназначен для автогенерации Инсталляционных карт по шаблону.
Сам шаблон выполнен в нотации JINJA2 и приложен к скрипту.
Вы можете менять стилевое оформление скрипта как угодно, менять местами куски текста и таблицы по вашему вкусу.
Главное - не трогать переменные JINJA2 (имеют вид "{{ something }}" )

Как пользоваться скриптом:
1. Копируем всю папку "auto-IM" куда-нибудь к себе
2. Для работы необходим установленный python 3.9 (на других 3.x тоже должно работать, но не проверялось), также необходим pip (как правило устанавливается вместе с python)
3. Необходимо наличие пакетов pandas~=2.2.2, openpyxl~=3.1.5, numpy~=2.0.1, docxtpl~=0.18.0
    Установка из командной строки:
    - открываем командную строку
    - переходим в папку "auto-IM"
    - выполняем команду: pip install -r requirements.txt
    - готово
4. Заполняем базовую информацию в файле "properties.ini"
5. Кладем все нужные файлы в подпапку "../auto-IM/res".
    Скрипту нужны следующие файлы:
    - NET_PASSPORT,
    - Кабельный журнал "...Сеть...",
    - Спецификация,
    - файл "RAID",
    - файлы lsblk.txt, vm-hw.txt, fstab.txt (опционально вместо этих файлов можно положить архив report.zip в котором присутстсвуют данные файлы),
    - рисунки: "Размещение...", "_irack", "_rack", "_L2 L3" в форматах jpg/png
6. Если каких-то файлов не будет найденов в подпаке /res - эти данные не попадут в итоговый отчет и вместо них будут значения UNKNOWN.
7. В командной строке, находясь в папке "auto-IM" запускаем скрипт: python autoim.py
8. Следим за выполнением программы, до появления сообщения FINISHED. В случае вылетов и ошибок - пишите мне, возможно скрипт пока сыроват.