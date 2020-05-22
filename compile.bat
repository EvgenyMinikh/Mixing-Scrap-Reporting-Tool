pyinstaller --clean -F -i label.ico --noconsole --additional-hooks-dir=hooks Mixing_Scrap_Reporting.py
copy config.cfg .\dist\
copy list_items.json .\dist\
copy main_window.ui .\dist\
copy mixing_label_template.svg .\dist\
copy shift_supervisors.csv .\dist\
copy operators.csv .\dist\