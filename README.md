pyinstaller --onefile --windowed --icon=image.ico --name "ReporteISP" main.py


pyinstaller --onefile --windowed --icon=image.ico --add-data ".env;." --add-data "image.ico;." --name "ReporteISP" main.py

pyinstaller --onefile --windowed --icon=image.ico --name "ReporteISP" main.py