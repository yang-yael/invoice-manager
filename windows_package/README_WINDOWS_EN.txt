Invoice Manager Windows Package

Use on a computer WITHOUT Python:
1. Build web_app.exe on one Windows computer that has Python.
2. To build, double-click build_exe.cmd.
3. After web_app.exe is created, copy this whole folder to other Windows computers.
4. On other computers, double-click start.cmd.

Use on a computer WITH Python:
1. Double-click start.cmd.

Important files:
- invoice_system/web_app.py
- invoice_system/invoice_manager.py
- invoice_template: 发票模板.xlsx
- uploads: original uploaded files
- outputs: generated results

If PDF text cannot be read, add Poppler here:
tools/poppler/Library/bin/pdftotext.exe
tools/poppler/Library/bin/pdftoppm.exe
