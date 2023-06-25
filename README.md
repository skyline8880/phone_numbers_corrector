# phone_numbers_corrector

Application for correcting list of phone numbers

### create virtual environment

```bash
python -m venv venv
```

make shure your environment is activated

### requirements

```bash
pip install -r requirements.txt
```

### packaging using pyinstaller

write command below in terminal

```bash
pyinstaller --onefile --windowed --add-data 'sp_icon.png' --icon=sp.ico main.py
```
