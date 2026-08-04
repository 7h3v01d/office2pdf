# Office2PDF icon integration

The supplied `badge.ico` has been preserved unchanged as `assets/office2pdf.ico`. The ICO directory contains six embedded Windows icon images, including 128×128 and 256×256 variants.

## Integrated surfaces

- PyQt application icon
- Main-window title-bar icon
- Windows taskbar identity and icon
- `Office2PDF.exe` file icon
- `Office2PDFNativeWorker.exe` file icon
- Desktop shortcut icon
- Start Menu shortcut icon

## Build workflow

1. Run `test.bat`.
2. Run `build_exe.bat`.
3. Keep the complete `dist\Office2PDF` folder together.
4. Move that folder to its permanent location.
5. Run `Create Office2PDF Shortcuts.bat` from inside the release folder.

The shortcut script uses Windows' actual Desktop and Programs-folder locations, including redirected Desktop configurations such as OneDrive.
