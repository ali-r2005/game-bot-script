# Dofus Automation Script

This Python script automates interactions with the game "Dofus" on Windows. It uses window detection, screen capture, image processing, and OCR to perform actions such as fighting enemies (e.g., strawberries) and managing inventory.

## Features

- Automatic detection of the "Dofus" game window
- Screen capture and template matching to locate UI elements and enemies
- Simulated mouse clicks for precise interactions
- OCR integration to read in-game text and make decisions
- Automated combat sequences for specific enemies
- Inventory management based on game state

## Requirements

- **Operating System:** Windows
- **Dependencies:**
  - `pywin32`
  - `Pillow`
  - `opencv-python`
  - `pytesseract`
  - `numpy`
  - `requests`
- **Tesseract OCR:** Must be installed and configured
- **Image Templates:** The following PNG files must be in the project directory:
  - `x.png`
  - `firstWeapon.png`
  - `secondWeapon.png`
  - `thirtWeapon.png`
  - `main.png`
  - `trash.png`
  - `redS.png`
  - `yellowS.png`
  - `whitS.png`
  - `greenS.png`

## Installation

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/ali-r2005/game-bot-script.git