python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python run.py


python -m pip install -r requirements.txt
python init_db.py
python run.py

####Ticketing app###########
rmdir /s /q build
rmdir /s /q dist
python -m PyInstaller --noconfirm --clean TicketingApp.spec
####################################

####Repair Tracker#########
python -m PyInstaller launcher.py --noconfirm --onedir --windowed --name RepairTracker `
--collect-all reportlab `
--collect-all openpyxl `
--collect-all webview `
--collect-all PIL `
--hidden-import=webview `
--hidden-import=openpyxl.cell._writer `
--hidden-import=openpyxl.worksheet._writer `
--hidden-import=reportlab.graphics.barcode.code128 `
--hidden-import=reportlab.graphics.barcode.code39 `
--hidden-import=reportlab.graphics.barcode.code93 `
--hidden-import=reportlab.graphics.barcode.common `
--hidden-import=reportlab.graphics.barcode.codabar `
--hidden-import=reportlab.graphics.barcode.eanbc `
--hidden-import=reportlab.graphics.barcode.usps `
--hidden-import=reportlab.graphics.barcode.usps4s `
--hidden-import=reportlab.graphics.barcode.widgets `
--hidden-import=reportlab.graphics.barcode.qr `
--add-data "app;app"
#################################