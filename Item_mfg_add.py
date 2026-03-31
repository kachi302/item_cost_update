#import sys
import os

import pyautogui as pag
import time
import datetime as dt
from openpyxl import load_workbook
import pymsgbox as pmg
from pyscreeze import locateAllOnScreen

IMG_DIR = r"D:\python_dev\item_cost_update\images"
TOOLS_IMG = os.path.join(IMG_DIR, "ITEM_TOOLS.png")

def wait_locate(img_path, timeout=15, confidence=0.75, grayscale=True, interval=0.5, region=None):
    end = time.time() + timeout
    while time.time() < end:
        loc = pag.locateOnScreen(img_path, confidence=confidence, grayscale=grayscale, region=region)
        if loc:
            return loc
        time.sleep(interval)
    return None

def safe_tab():
    for k in ("ctrl","alt","shift"):
        pag.keyUp(k)
    pag.keyDown("tab")
    time.sleep(0.5)
    pag.keyUp("tab")    

running = True
print('Item Mfg Group 추가...3초후 시작됩니다. 해당 화면을 다른 모니터로 이동 하세요.')
pag.countdown(3)
# item_code = ''
# planner_code = ''18669305-004
try:
    print('Start')
    # pag.click(1000, 800)
    wb = load_workbook(r'D:\NOV_Process\RS_SUPPORT_TEAM\100.NewProject\Automatic creation of item cost\Item_cost_category_01.xlsx', data_only= True)
    ws = wb['Sheet1']
    start1 = time.time()
    rec = 0
    time.sleep(1)
    dest = "D:\\python_dev\\item_cost_update\\images\\"
    # job = pag..confirm('Item Planner 변경 작업을 진행 하시겠습니끼?', buttons= ['OK','Cancel']) # type: ignore    
    job = pmg.confirm('Item MFG 추가 작업을 진행 하시겠습니끼?', buttons= ['OK','Cancel']) 
    if job == 'OK':
        print(time.strftime('%Y.%m%d - %H:%M:%S'))
        time.sleep(3)
        for row in ws.iter_rows():
            rec += 1
            item = str(row[0].value)
            time.sleep(0.3)
            # item_len = len(item)
            #mfg_group = str(row[1].value)
            mfg_group = "Item Cost Category"
            #mfg_value = str(row[2].value)
            mfg_value = str(row[1].value)
           # planning = str(row[3].value)
           # planning_value = str(row[4].value)
            print('Rec : %d , item = %s' % (rec,item))
            # if item =='E':
            #     print('End....')
            #     break
            # if item_len <= 11:
            #     print('Rec : %d , item = %s, Categories : %s' % (rec,item, 'item error', mfg_value))
            #     break
            # item 차기
            pag.hotkey('f11')
            time.sleep(1.5)
            pag.write(item, interval=0.08)
            time.sleep(1)
            pag.hotkey('ctrl','f11')
            time.sleep(2)
        #    tools_png = wait_locate(TOOLS_IMG, timeout=15, confidence=0.75, grayscale=True)
        #    if not tools_png:
        #        pag.screenshot(dest + f"tools_not_found_{rec}_{item}.png")
        #        print(f"Rec : {rec}, item = {item}, tools_png not found (load delay / popup / scaling 가능)")
        #        continue   # break 하지 말고 다음 행으로 넘기는 것을 권장
        #    pag.click(pag.center(tools_png))
            pag.hotkey('alt','t')
            time.sleep(1)
            # tool button 찾기
            #tools_png = pag.locateOnScreen(r'D:\python_dev\item_cost_update\images\ITEM_TOOLS.png', confidence=0.75)
            #if tools_png is not None:
            #    pag.click(pag.center(tools_png))
            #    time.sleep(0.5)
            #pag.press('down')   
            #pag.hotkey('enter')     
            pag.press('enter')
            time.sleep(3)  
                # tools -- categories
                # print('down after')
            category_png = pag.locateOnScreen(r'D:\python_dev\item_cost_update\images\ITEM_COST_CATEGORY.png', confidence=0.8)     
            if category_png is None:
            #    pag.click(pag.center(category_png))
                
            #    time.sleep(1)
                pag.hotkey('ctrl','down')
                time.sleep(2)
                pag.write(mfg_group, interval=0.08)
                time.sleep(2)
                # pag.press('enter')
                #pag.press('tab')
                safe_tab()
                time.sleep(2)
                pag.write(mfg_value, interval=0.08)               
                time.sleep(2)
                safe_tab()
                time.sleep(1)
                pag.hotkey('ctrl','s')
                time.sleep(2)
                        # planning group 추가
                    # pag.press('tab')
                    # time.sleep(2)
                    # pag.hotkey('f11')
                    # time.sleep(0.2)
                   # pag.write(planning,interval=0.03)
                   # pag.press('tab', presses=3)
                   # pag.press('enter')
                   # time.sleep(0.5)
                   # pag.press('tab')
                   # pag.write(planning_value, interval=0.03)
                   # pag.hotkey('ctrl','s')
                   # time.sleep(7)
                    
                    
                pag.hotkey('ctrl','f4')
                    #close_png = pag.locateOnScreen(r'D:\python_dev\item_cost_update\images\Close_page.png', confidence=0.8)
                    #if close_png is not None:
                    #    pag.click(pag.center(close_png))
                time.sleep(1.5)
            else:
                pag.hotkey('ctrl','f4')
                time.sleep(1.5)
                #print('Rec : %d , item = %s, %s' % (rec,item, 'tools_png error'))
        #    print(f"Rec : {rec} , item = {item}, item error, value : {mfg_value}")
        #    pag.screenshot(dest+"tools_png.jpg")
        #    time.sleep(1)
        #    break   
            
    print(time.strftime('%Y.%m%d - %H:%M:%S'))        
    pmg.alert('작업을 종료합니다.')
    print('작업을 종료합니다.')
except Exception as e:
    # wb.active = ws # type: ignore
    # wb.save(r'.\Planner_job.xlsx')    # type: ignore
    print(e)    
    time.sleep(10)       