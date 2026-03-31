# 리팩토링된 ERP Cost Update Tool
import tkinter as tk
from tkinter import messagebox
import pyautogui
import time
from openpyxl import load_workbook
import configparser
import os
import datetime as dt

# 전역 설정 불러오기
CONFIG_PATH = './data/config.ini'
CONFIG = configparser.ConfigParser()
CONFIG.read(CONFIG_PATH)

IMAGE_PATH = './images/'
EXCEL_PATH = './data/'

# 이미지 경로 및 폴더 유효성 체크 함수
def validate_image_files(image_dict):
    missing = []
    if not os.path.exists(IMAGE_PATH):
        os.makedirs(IMAGE_PATH)
        messagebox.showwarning('경고', f'"images" 폴더가 없어 생성했습니다. 여기에 이미지 파일을 넣으세요.')
    for name, path in image_dict.items():
        if not os.path.isfile(path):
            missing.append(name)
    if missing:
        messagebox.showerror('오류', f'다음 이미지 파일이 존재하지 않습니다: {", ".join(missing)}')
        return False
    return True

# 설정값 불러오기
SCREEN_COORDS = {
    'ct_x1': CONFIG.getint('Screen_3WAY', 'X1'),
    'ct_x2': CONFIG.getint('Screen_3WAY', 'X2'),
    'ct_x3': CONFIG.getint('Screen_3WAY', 'X3'),
    'ct_x4': CONFIG.getint('Screen_3WAY', 'X4'),
    'dx1': CONFIG.getint('Screen_Not_Default', 'DX1'),
    'dy1': CONFIG.getint('Screen_Not_Default', 'DY1'),
    'dx2': CONFIG.getint('Screen_Not_Default', 'DX2'),
    'dy2': CONFIG.getint('Screen_Not_Default', 'DY2')
}

COST_ELEMENTS = {
    'cost_element1': 'Material',
    'cost_element2': 'Material Overhead',
    'sub_element2': 'Freight'
}

IMAGES = {
    'item': os.path.join(IMAGE_PATH, 'item_cost1.png'),
    'cost_3way': os.path.join(IMAGE_PATH, 'basedroll1.png'),
    'not_default': os.path.join(IMAGE_PATH, 'item_cost_not_default.png'),
    'inventory_asset': os.path.join(IMAGE_PATH, 'inv_asset.png'),
    'cost_type_not_found': os.path.join(IMAGE_PATH, 'cost_type_not_found.png'),
    'screen_max': os.path.join(IMAGE_PATH, 'screen_max.png'),
    'find': os.path.join(IMAGE_PATH, 'find.png'),
    'standard': os.path.join(IMAGE_PATH, 'standard.png'),
    'summit_btn': os.path.join(IMAGE_PATH, 'summit_btn.png'),
    'nov_cost_update': os.path.join(IMAGE_PATH, 'nov_cost_update.png'),
    'costs_btn': os.path.join(IMAGE_PATH, 'costs_btn.png')
}

class ERPCostUpdater:
    def __init__(self, gui):
        self.gui = gui

    def update_progress(self, message):
        self.gui.progress_label.config(text=message)
        self.gui.root.update()

    def locate_and_Doubleclick(self, image_key, clicks=1, interval=0.25):
        path = IMAGES.get(image_key)
        location = pyautogui.locateOnScreen(path, confidence=0.8)
        if location:
            center = pyautogui.center(location)
            #pyautogui.click(center.x, center.y, clicks=clicks, interval=interval)
            pyautogui.doubleClick(center)
            return True
        else:
            return False
    
    def locate_and_one_click(self, image_key, clicks=1, interval=0.25):
        path = IMAGES.get(image_key)
        location = pyautogui.locateOnScreen(path, confidence=0.8)
        if location:
            center = pyautogui.center(location)
            #pyautogui.click(center.x, center.y, clicks=clicks, interval=interval)
            pyautogui.click(center)
            #image_center = center
            return True
        else:
            return False    

    #image 가 있으면 center 값 return
    def locate_and_Image(self, image_key, clicks=1, interval=0.25):
        path = IMAGES.get(image_key)
        location = pyautogui.locateOnScreen(path, confidence=0.8)
        if location:
            center = pyautogui.center(location)
            #pyautogui.click(center.x, center.y, clicks=clicks, interval=interval)
            #pyautogui.Click(center)
            image_center = center
            return image_center
        else:
            return False    
    
    def reset_annual_cost(self):
        try:
            self.gui.label_result.config(text='Annual Cost Reset')
            wb = load_workbook(os.path.join(EXCEL_PATH, 'ANNUAL_list.xlsx'))
            ws = wb['Sheet1']
            total = ws.max_row

            if pyautogui.confirm('Annual Cost Reset을 진행하시겠습니까?') != 'OK':
                return

            for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 1):
                item = row[0]
                self.update_progress(f'진행 중: {i}/{total} - {item}')
                if self.locate_and_Doubleclick('item'):
                    time.sleep(0.5)
                    pyautogui.write(str(item), interval=0.1)
                    pyautogui.press('tab')
                    time.sleep(0.5)
                    pyautogui.write('Annual', interval=0.1)
                    pyautogui.press('tab')
                    time.sleep(0.5)
                    self.locate_and_one_click('find')
                    time.sleep(1)
                    ric =  self.locate_and_Image('cost_3way')
                    #3곳이 다 check 되어 있는 경우우
                    if ric !=0:    
                        time.sleep(0.7)
                        pyautogui.click(ric.x-130, ric.y)
                        
                        # based on rollup unckeck
                        time.sleep(0.7)
                        pyautogui.click(ric.x+130, ric.y)
                        time.sleep(1)
                        pyautogui.hotkey('enter')
                        time.sleep(1) 
                        
                        pyautogui.click(ric.x+130, ric.y)                       
                        time.sleep(1)  
                        
                        # user defaul control check
                        pyautogui.click(ric.x-130, ric.y)
                        time.sleep(1)
                        
                        pyautogui.hotkey('ctrl','s')
                        time.sleep(1)
                        pyautogui.hotkey('ctrl','f4')
                        # close_form()    
                        time.sleep(1)
                    elif ric== 0:
                        time.sleep(0.7)
                        ric_not_default = self.locate_and_Image('not_default')
                        if ric_not_default != 0:
                            pyautogui.click(ric.x+115, ric.y+10)
                            time.sleep(0.9)
                            
                            # cost 지울지 메세지 나타나남, enter key 입력
                            pyautogui.hotkey('enter')
                            time.sleep(0.9)
                            
                            pyautogui.click(ric.x+118, ric.y+10)
                            time.sleep(0.9)
                            pyautogui.hotkey('ctrl','s')  
                        elif ric_not_default == 0:
                            ric_inventory = self.locate_and_Image('inventory_asset')
                            if ric_inventory !=0:
                                time.sleep(0.5)
                                self.locate_and_one_click('costs_btn')    
                                time.sleep(0.9)                               
                                pyautogui.write(COST_ELEMENTS.get('cost_element1'), interval=0.1)
                                time.sleep(0.9)
                                pyautogui.hotkey('tab')
                                pyautogui.write(COST_ELEMENTS.get('cost_element1'), interval=0.1)
                                time.sleep(0.9)                        
                                pyautogui.press('tab', presses=3, interval=0.5) 
                                time.sleep(0.9)
                                # Material cost 입력
                                pyautogui.write(str(0), interval=0.1)
                                time.sleep(0.9)
                                pyautogui.hotkey('tab')
                                time.sleep(0.9)
                                pyautogui.hotkey('tab')
                                # material overhead
                                pyautogui.write(COST_ELEMENTS.get('cost_element2'), interval=0.1)
                                time.sleep(0.5)
                                pyautogui.hotkey('tab')
                                # sub element
                                time.sleep(0.5)
                                pyautogui.write(COST_ELEMENTS.get('sub_element2'), interval=0.1)
                                time.sleep(0.5)
                                pyautogui.hotkey('tab')
                                time.sleep(0.5)
                                pyautogui.press('tab', presses=2, interval=0.3) 
                                # Material Overhead rate 입력
                                time.sleep(0.9)                        
                                pyautogui.write(str(0), interval=0.1)
                                time.sleep(0.9)
                                pyautogui.hotkey('tab')                        
                                time.sleep(0.9)
                                pyautogui.hotkey('ctrl','s')
                                # close_form()   
                                time.sleep(0.9)
                                pyautogui.hotkey('ctrl','f4')
                                time.sleep(0.9)
                            else:
                                messagebox.showwarning('경고', f'Item 이미지를 찾을 수 없습니다: {item}')       
                                  
                    else:
                        messagebox.showwarning('경고', f'Item 이미지를 찾을 수 없습니다: {item}')   
                    
                else:
                    messagebox.showwarning('경고', f'Item 이미지를 찾을 수 없습니다: {item}')
            messagebox.showinfo('완료', 'Annual Cost Reset 완료')

        except Exception as e:
            messagebox.showerror('오류', f'Annual Cost Reset 중 오류 발생: {e}')

    def update_annual_cost(self):
        try:
            self.gui.label_result.config(text='Annual Cost Update')
            wb = load_workbook(os.path.join(EXCEL_PATH, 'ANNUAL_UPDATE.xlsx'))
            ws = wb['Sheet1']
            total = ws.max_row

            if pyautogui.confirm('Annual Cost Update를 진행하시겠습니까?') != 'OK':
                return

            for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 1):
                item = row[0]
                self.update_progress(f'진행 중: {i}/{total} - {item}')
                if self.locate_and_click('item'):
                    time.sleep(0.5)
                    pyautogui.write(str(item), interval=0.1)
                    pyautogui.press('tab')
                    pyautogui.write('Annual', interval=0.1)
                    pyautogui.press('tab')
                    self.locate_and_click('find')
                    time.sleep(1)
            messagebox.showinfo('완료', 'Annual Cost Update 완료')

        except Exception as e:
            messagebox.showerror('오류', f'Annual Cost Update 중 오류 발생: {e}')

    def update_pending_frozen_cost(self):
        try:
            self.gui.label_result.config(text='Pending/Frozen Cost Update')
            wb = load_workbook(os.path.join(EXCEL_PATH, 'pending.xlsx'))
            ws = wb['Sheet1']
            total = ws.max_row

            if pyautogui.confirm('Pending/Frozen Cost Update를 진행하시겠습니까?') != 'OK':
                return

            for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 1):
                item = row[0]
                self.update_progress(f'진행 중: {i}/{total} - {item}')
                if self.locate_and_click('item'):
                    time.sleep(0.5)
                    pyautogui.write(str(item), interval=0.1)
                    pyautogui.press('tab')
                    pyautogui.write('Pending', interval=0.1)
                    pyautogui.press('tab')
                    self.locate_and_click('find')
                    time.sleep(1)
            messagebox.showinfo('완료', 'Pending/Frozen Cost Update 완료')

        except Exception as e:
            messagebox.showerror('오류', f'Pending/Frozen Cost Update 중 오류 발생: {e}')


class CostUpdaterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('ERP 원가 자동화 도구')
        self.root.geometry('800x400')
        self.updater = ERPCostUpdater(self)

        if not validate_image_files(IMAGES):
            self.root.destroy()
            return

        self._create_widgets()

    def _create_widgets(self):
        tk.Label(self.root, text='ERP 원가 업데이트 프로그램입니다.', fg='red').pack()
        tk.Label(self.root, text='프로그램 창은 ERP 화면과 겹치지 않도록 해주세요.', fg='red').pack()

        tk.Button(self.root, text='Annual Cost Reset', width=30, height=2,
                  command=self.updater.reset_annual_cost).pack(pady=5)

        tk.Button(self.root, text='Annual Cost Update', width=30, height=2,
                  command=self.updater.update_annual_cost).pack(pady=5)

        tk.Button(self.root, text='Pending/Frozen Cost Update', width=30, height=2,
                  command=self.updater.update_pending_frozen_cost).pack(pady=5)

        self.progress_label = tk.Label(self.root, text='진행 상황', fg='blue')
        self.progress_label.pack()

        self.label_result = tk.Label(self.root, text='', fg='green')
        self.label_result.pack()

        tk.Label(self.root, text='Annual 초기화 파일: ANNUAL_list.xlsx', fg='blue').pack()
        tk.Label(self.root, text='Annual 업데이트 파일: ANNUAL_UPDATE.xlsx', fg='blue').pack()
        tk.Label(self.root, text='Pending 관련 파일: pending.xlsx', fg='blue').pack()

        self.root.bind('<Escape>', lambda e: self.root.destroy())

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    gui = CostUpdaterGUI()
    gui.run()
