# MT4 Alert Automation Script

This script automates the process of detecting and handling MetaTrader 4 alerts using audio signals and OCR.

## Prerequisites

- Python 3.7+
- MetaTrader 4
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- [VB-Cable Virtual Audio Device](https://vb-audio.com/Cable/)

## Required Python Packages

```bash
pip install numpy soundfile scipy sounddevice pywin32 pytesseract Pillow psutil
```

## Configuration

1. Install Tesseract OCR and set the path in the script:
   ```python
   pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
   ```

2. Configure your reference audio file:
   ```python
   REFERENCE_FILE = r"PATH\TO\YOUR\REFERENCE_AUDIO.wav"
   ```

3. Adjust window coordinates if needed:
   ```python
   ALERT_WINDOW_COORDS = (10, 45, 380, 65)
   ```

## Usage

1. Start MetaTrader 4
2. Ensure VB-Cable is properly installed and set as the default output device
3. Run the script as Administrator:
   ```bash
   python main.py
   ```

## Features

- Audio signal detection using cross-correlation
- Automated window handling
- OCR text extraction from alerts
- Logging system for detected alerts
- Cooldown period to prevent multiple triggers

## Troubleshooting

- Run as Administrator for proper window handling
- Verify VB-Cable is properly installed and configured
- Check if MetaTrader 4 is running before starting the script
- Verify Tesseract OCR installation path

## Notes

- The script requires admin privileges for proper window handling
- Cooldown period is set to 30 seconds by default
- Minimum time between audio signals is 3.0 seconds
