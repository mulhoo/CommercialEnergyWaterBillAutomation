@echo off
echo Downloading dependencies...

mkdir dependencies
cd dependencies

REM Download Tesseract portable
curl -L -o tesseract.zip "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3/tesseract-ocr-w64-setup-5.3.3.20231005.exe"

REM Download Poppler
curl -L -o poppler.zip "https://github.com/oschwartz10612/poppler-windows/releases/download/v23.08.0-0/Release-23.08.0-0.zip"
powershell -command "Expand-Archive poppler.zip -DestinationPath ."

cd ..

echo Building executable...
pip install pyinstaller
pyinstaller --onefile --windowed --name=WaterBillProcessor ^
  --add-data="config.py;." ^
  --add-data="*.xlsx;." ^
  --add-binary="dependencies/poppler-23.08.0/Library/bin/*;poppler/" ^
  main.py

echo Creating distribution...
mkdir distribution
copy dist\WaterBillProcessor.exe distribution\
copy Install.bat distribution\
copy README-Windows.txt distribution\

echo Done! Check the distribution folder.