# Text tracker (Game Boss Tracker) (Python + PyTesseract)

## General Information
This Python project is designed to **monitor in-game events** by reading text from screenshots using **OCR (PyTesseract)**.  
It automatically updates an **Excel sheet** based on detected patterns, creating a **visual table** of game events.  
The tool was originally used to track bosses in a game, highlighting which player defeated which boss and on which channel.

## Features
- Take screenshots of a specified area on the screen  
- Use **PyTesseract OCR** to extract text from the screenshots  
- Filter extracted text based on a **predefined pattern**  
- Automatically open and update an **Excel spreadsheet**  
- Mark detected events in the Excel sheet with:
  - Player names (nicknames)  
  - Color-coded cells for visual clarity  
- Track multiple bosses or events in a structured table  
- Fully automated: updates happen as events appear in the game chat  

## Software / Tech Stack
- Python 3.x  
- PyTesseract (OCR)  
- Pillow (PIL) for screenshots  
- OpenPyXL for Excel operations  
 

## How It Works
1. Launch the program.  
2. The program takes screenshots of the predefined screen area at regular intervals.  
3. OCR (PyTesseract) reads text from each screenshot.  
4. The extracted text is filtered based on a **pattern** (e.g., "Boss X killed by Player Y on channel Z").  
5. When a match is found:
   - The Excel sheet is automatically updated  
   - The cell corresponding to the event is **filled with the player’s name**  
   - Color coding is applied for easy visualization  
6. The table gives a real-time overview of boss kills and player activity.

## Notes
- The project is intended as a **game monitoring tool** using OCR and Excel automation.  
- The pattern used for detection can be customized depending on the game’s chat format.  
- Works best with **high-contrast text** in the screenshots for accurate OCR.  

## Possible Improvements / Future Work
- Add multi-language support for text recognition  
- Implement automatic Excel formatting for multiple bosses and channels  
- Add GUI for easier configuration of screenshot area and pattern  
- Log historical events for trend analysis or statistics  
