import os
import time
from PIL import ImageGrab, Image, ImageFilter
import pytesseract as tess
import re
import xlwings as xw
from openpyxl.styles import PatternFill

tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class ImageProcessor:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.define_columns()
        self.last_patterns = []

    def define_columns(self):
        # Excel columns definition
        self.boss_columns = ['Boss1', 'Boss2']
        
        # Channels in game definition- 8 channels here
        room_columns = [str(i) for i in range(1, 9)]
        
        if not os.path.isfile(self.excel_file):
            wb = xw.Book()
            wb.sheets.add('Data')
            wb.save(self.excel_file)
        
        # Excel
        wb = xw.Book(self.excel_file)
        sheet = wb.sheets['Data']
        if not sheet.range('A1').value: 
            sheet.range('A1').value = 'Channel'
            for index, room in enumerate(room_columns):
                sheet.range((index + 2, 1)).value = room
            for index, fruit in enumerate(self.boss_columns):
                sheet.range((1, index + 2)).value = fruit

    # Photo process

    def process_image(self, image_path):
        image = Image.open(image_path)
        enlarged_image = image.resize((int(image.width * 3), int(image.height * 3)))
        
        sharpened_image = enlarged_image.filter(ImageFilter.SHARPEN)
        
        text = tess.image_to_string(sharpened_image, lang='eng')
        if self.is_new_pattern(text):
            self.update_excel(text)
            self.last_patterns.append(text)  
            if len(self.last_patterns) > 10:  
                self.last_patterns.pop(0) 
        os.remove(image_path)

    def is_new_pattern(self, text):
        if text in self.last_patterns:
            return False
        return True

    def update_excel(self, text):
        pattern = r"\[Channel (\d+)\].*? - (Boss1|Boss2)\."
        matches = re.findall(pattern, text)

        players = {}
        for match in matches:
            channel, boss = match
            key = (channel, boss)  # key for boss + channel
            if key in players:
                players[key] += 1
            else:
                players[key] = 1

        wb = xw.Book(self.excel_file)
        sheet = wb.sheets['Data']

        # Update Excel data
        for (channel, boss), count in players.items():
            if boss in self.boss_columns:  # Check if the boss is in the list of bosses
                cell = sheet.range((int(channel) + 1, self.boss_columns.index(boss) + 2))
                # Check if the cell is empty before updating
                if cell.value is None:
                    cell.value = 'x' * count
                else:
                    # If cell is not empty, append 'x' to existing value
                    cell.value += 'x' * count

        # Highlight updated cells
        orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
        num_columns = 10
        range_down_right = sheet.range('B2').expand('table').resize(row_size=10, column_size=num_columns)

        for cell in range_down_right:
            if cell.value is not None:
                cell.color = orange_fill.start_color.index
                cell.fill = orange_fill

        for column in sheet.range('A1').expand('right'):
            column.column_width = max(len(str(cell.value)) for cell in column)

        print("Excel updated.")

class MainApp:
    def __init__(self, folder_path, excel_file, update_interval):
        self.folder_path = folder_path
        self.excel_file = excel_file
        self.update_interval = update_interval 
        self.image_processor = ImageProcessor(excel_file)

    def watch_and_process_screen_region(self):
        while True:

            time.sleep(self.update_interval)

            # coordinates
            left = 0    
            top = 0     
            width = 450 
            height = 400

            snapshot = ImageGrab.grab(bbox=(left, top, left+width, top+height))
            file_name = os.path.join(self.folder_path, str(time.time()) + ".png")
            # screen save
            snapshot.save(file_name)
            self.image_processor.process_image(file_name)
            time.sleep(2) 

if __name__ == "__main__":
    folder_path = '' # path
    excel_file = 'namexcel.xlsx'
    update_interval = 1 # time interval
    app = MainApp(folder_path, excel_file, update_interval)
    app.watch_and_process_screen_region()
