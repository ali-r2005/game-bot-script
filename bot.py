import time
import random
import win32gui
import win32ui
import win32con
import win32api
import win32print
from ctypes import windll
from PIL import Image
import cv2
import pytesseract
import numpy as np
import sys
import ctypes
import requests

sys.setrecursionlimit(10**6) 


def listWindows():
    def enum_windows_callback(hwnd, lParam):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            if window_text and win32gui.IsWindowEnabled(hwnd):
                windows_on_taskbar.append(window_text)
        return True

    windows_on_taskbar = []
    win32gui.EnumWindows(enum_windows_callback, None)
    return windows_on_taskbar

def getNames():
    taskbar_window_names = listWindows()
    namesList = []
    for idx, window_name in enumerate(taskbar_window_names, start=1):
        try:
            window_bytes = window_name.encode('utf-8', 'ignore')
            decoded_name = window_bytes.decode('utf-8')
            namesList.append(decoded_name)
        except UnicodeEncodeError:
            c = 1
    return namesList

def detect():
    themLists = getNames()
    selected = None
    for thisWindow in themLists:
        if "Dofus" in  thisWindow:
            selected = thisWindow
            break
    return selected

hwnd = win32gui.FindWindow(None, detect())

def clickIT(handle, x, y, isLeft):
  # Get window information
        window_rect = win32gui.GetWindowRect(handle)
        client_rect = win32gui.GetClientRect(handle)
        
        # Calculate border offsets
        border_width = ((window_rect[2] - window_rect[0]) - (client_rect[2] - client_rect[0])) // 2
        title_height = ((window_rect[3] - window_rect[1]) - (client_rect[3] - client_rect[1])) - border_width
        
        # Get DPI scaling
        dc = win32gui.GetDC(0)
        dpi_x = win32print.GetDeviceCaps(dc, win32con.LOGPIXELSX)
        scale_factor = dpi_x / 96.0
        win32gui.ReleaseDC(0, dc)
        
        # Adjust coordinates for DPI and window position
        adjusted_x = int((x - window_rect[0]) / scale_factor)
        adjusted_y = int((y - window_rect[1]) / scale_factor)
        
        # Create LPARAM for mouse position
        lparam = win32api.MAKELONG(adjusted_x, adjusted_y)
        if isLeft:
            time.sleep(0.1)
            win32gui.PostMessage(handle, win32con.WM_MOUSEMOVE, 3, lparam)
            time.sleep(0.1)
            win32gui.PostMessage(handle, win32con.WM_MOUSEMOVE, 3, lparam)
            win32gui.PostMessage(handle, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.05)
            win32gui.PostMessage(handle, win32con.WM_LBUTTONUP, 0, lparam)
        else:
            win32gui.PostMessage(handle, win32con.WM_MOUSEMOVE, 3, lparam)
            time.sleep(0.05)
            win32gui.PostMessage(handle, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lparam)
            win32gui.PostMessage(handle, win32con.WM_RBUTTONUP, 0, lparam)

    
def doubleClickIT(handle, x, y):
    # Get window information
    window_rect = win32gui.GetWindowRect(handle)
    client_rect = win32gui.GetClientRect(handle)
        
    # Calculate border offsets
    border_width = ((window_rect[2] - window_rect[0]) - (client_rect[2] - client_rect[0])) // 2
    title_height = ((window_rect[3] - window_rect[1]) - (client_rect[3] - client_rect[1])) - border_width
        
    # Get DPI scaling
    dc = win32gui.GetDC(0)
    dpi_x = win32print.GetDeviceCaps(dc, win32con.LOGPIXELSX)
    scale_factor = dpi_x / 96.0
    win32gui.ReleaseDC(0, dc)
        
    # Adjust coordinates for DPI and window position
    adjusted_x = int((x - window_rect[0]) / scale_factor)
    adjusted_y = int((y - window_rect[1]) / scale_factor)
        
    # Create LPARAM for mouse position
    lparam = win32api.MAKELONG(adjusted_x, adjusted_y)
    time.sleep(0.1)
    win32gui.PostMessage(handle, win32con.WM_MOUSEMOVE, 3, lparam)
    time.sleep(0.1)
    win32gui.PostMessage(handle, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    win32gui.PostMessage(handle, win32con.WM_LBUTTONUP, 0, lparam)
    win32gui.PostMessage(handle, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    win32gui.PostMessage(handle, win32con.WM_LBUTTONUP, 0, lparam)

def defaultPixelCheck(handle, thisPixelX, thisPixelY, thisHex):
    hwnd = handle

    # Change the line below depending on whether you want the whole window
    # or just the client area. 
    #left, top, right, bot = win32gui.GetClientRect(hwnd)
    #left, top, right, bot = win32gui.GetWindowRect(hwnd)
    # Get window dimensions
    client_rect = win32gui.GetClientRect(handle)
        
     # Calculate actual client area dimensions
    w = client_rect[2] - client_rect[0]
    h = client_rect[3] - client_rect[1]

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

    saveDC.SelectObject(saveBitMap)

    # Change the line below depending on whether you want the whole window
    # or just the client area. 
    #result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)

    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    if result == 1:
        #PrintWindow Succeeded
        #im.save("test.png")

        # Open the image
        #image = Image.open("test.png")
        image = im

        # Get the current width and height of the image
        width, height = image.size

        # Specify the height of the black lines to add
        top_line_height = 23
        bottom_line_height = 40

        # Calculate the new height of the image
        new_height = height + top_line_height + bottom_line_height

        # Create a new image with the new dimensions
        new_image = Image.new("RGB", (width, new_height), color=(0, 0, 0))

        # Paste the original image into the new image, shifted down by the top line height
        new_image.paste(image, (0, top_line_height))

        # Save the modified image
        #new_image.save("modified.png")
        im = new_image

        # Convert Image to NumPy array
        im_array = np.array(im)

        # Convert BGR to RGB
        im_array_rgb = cv2.cvtColor(im_array, cv2.COLOR_BGR2RGB)

        isFound = False
        
        pixel_value = im_array_rgb[thisPixelY, thisPixelX]
        rgb_value = cv2.cvtColor(np.uint8([[pixel_value]]), cv2.COLOR_BGR2RGB)
        hex_value = '#{:02x}{:02x}{:02x}'.format(rgb_value[0][0][0], rgb_value[0][0][1], rgb_value[0][0][2])
        if hex_value == thisHex:
            isFound = True
        
        return isFound

def realityCheck(handle, path, conf, withPos):
    """
    Enhanced accuracy version of realityCheck
    """
    try:
        # Get window dimensions
        window_rect = win32gui.GetWindowRect(handle)
        client_rect = win32gui.GetClientRect(handle)
        
        # Calculate actual client area dimensions
        w = client_rect[2] - client_rect[0]
        h = client_rect[3] - client_rect[1]

        # Get window capture
        hwndDC = win32gui.GetWindowDC(handle)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)

        # Capture window (PW_RENDERFULLCONTENT = 3)
        result = windll.user32.PrintWindow(handle, saveDC.GetSafeHdc(), 3)

        if result == 1:
            # Convert to numpy array directly
            bmpstr = saveBitMap.GetBitmapBits(True)
            img_array = np.frombuffer(bmpstr, dtype=np.uint8)
            img_array = img_array.reshape(h, w, 4)  # RGBA format
            
            # Convert RGBA to BGR (OpenCV format)
            screenshot = cv2.cvtColor(img_array, cv2.COLOR_RGBA2BGR)

            # Load template
            template = cv2.imread(path)
            if template is None:
                raise FileNotFoundError(f"Template not found: {path}")

            # Convert both images to grayscale
            screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
            template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

            # Get template dimensions
            template_h, template_w = template_gray.shape

            # Perform template matching with multiple scales
            scales = [0.8, 0.9, 1.0, 1.1, 1.2]
            max_val = -1
            max_loc = None
            best_scale = 1.0

            for scale in scales:
                # Resize screenshot (instead of template to maintain quality)
                scaled_screenshot = cv2.resize(screenshot_gray, 
                                            (int(screenshot_gray.shape[1] * scale),
                                             int(screenshot_gray.shape[0] * scale)))

                if scaled_screenshot.shape[0] < template_h or scaled_screenshot.shape[1] < template_w:
                    continue

                # Perform template matching
                result = cv2.matchTemplate(scaled_screenshot, template_gray, cv2.TM_CCOEFF_NORMED)
                min_val, curr_max_val, min_loc, curr_max_loc = cv2.minMaxLoc(result)

                if curr_max_val > max_val:
                    max_val = curr_max_val
                    max_loc = (int(curr_max_loc[0] / scale), int(curr_max_loc[1] / scale))
                    best_scale = scale

            # Check if match found
            if max_val >= conf:
                if withPos:
                    # Calculate center of template
                    center_x = max_loc[0] + (template_w // 2)
                    center_y = max_loc[1] + (template_h // 2)

                    # Add window offset
                    screen_x = window_rect[0] + center_x
                    screen_y = window_rect[1] + center_y

                    # Debug info
                    print(f"Match found at ({screen_x}, {screen_y}) with confidence {max_val:.2f}")
                    return screen_x, screen_y
                return True
            return None

    except Exception as e:
        print(f"Error in realityCheck: {e}")
        return None

    finally:
        # Cleanup
        try:
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(handle, hwndDC)
        except:
            pass


def gText(handle, Fx, Fy, Sx, Sy):
    hwnd = handle

    # Change the line below depending on whether you want the whole window
    # or just the client area. 
    #left, top, right, bot = win32gui.GetClientRect(hwnd)
    #left, top, right, bot = win32gui.GetWindowRect(hwnd)
    client_rect = win32gui.GetClientRect(handle)
        
    # Calculate actual client area dimensions
    w = client_rect[2] - client_rect[0]
    h = client_rect[3] - client_rect[1]

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

    saveDC.SelectObject(saveBitMap)

    # Change the line below depending on whether you want the whole window
    # or just the client area. 
    #result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)

    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    if result == 1:
        #PrintWindow Succeeded
        #im.save("test.png")

        # Open the image
        #image = Image.open("test.png")
        image = im

        # Get the current width and height of the image
        width, height = image.size

        # Specify the height of the black lines to add
        top_line_height = 23
        bottom_line_height = 40

        # Calculate the new height of the image
        new_height = height + top_line_height + bottom_line_height

        # Create a new image with the new dimensions
        new_image = Image.new("RGB", (width, new_height), color=(0, 0, 0))

        # Paste the original image into the new image, shifted down by the top line height
        new_image.paste(image, (0, top_line_height))

        # Save the modified image
        #new_image.save("modified.png")
        im = new_image

        # Convert Image to NumPy array
        im_array = np.array(im)

        # Convert BGR to RGB
        im_array_rgb = cv2.cvtColor(im_array, cv2.COLOR_BGR2RGB)

        processed_image = cv2.cvtColor(im_array_rgb, cv2.COLOR_RGB2GRAY)

        x1, y1 = Fx, Fy  # Top-left corner coordinates
        x2, y2 = Sx, Sy  # Bottom-right corner coordinates
        cropped_image = processed_image[y1:y2, x1:x2]
        oc = pytesseract.image_to_string(cropped_image)
        #for character in oc:
        #    if character == " " or character == "\n":
        #        oc = oc.replace(character, "")
        if oc is not None:
            return oc
        else:
            return None


def realityCheckS(handle, path, conf, withPos):
    """
    Enhanced realityCheck function with RGB channel-based matching
    """
    try:
        # Get window dimensions
        window_rect = win32gui.GetWindowRect(handle)
        client_rect = win32gui.GetClientRect(handle)
        
        # Calculate actual client area dimensions
        w = client_rect[2] - client_rect[0]
        h = client_rect[3] - client_rect[1]

        # Get window capture
        hwndDC = win32gui.GetWindowDC(handle)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)

        # Capture window (PW_RENDERFULLCONTENT = 3)
        result = windll.user32.PrintWindow(handle, saveDC.GetSafeHdc(), 3)

        if result == 1:
            # Convert to numpy array directly
            bmpstr = saveBitMap.GetBitmapBits(True)
            img_array = np.frombuffer(bmpstr, dtype=np.uint8)
            img_array = img_array.reshape(h, w, 4)  # RGBA format
            
            # Convert RGBA to BGR
            screenshot = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB )

            # Load template
            template = cv2.imread(path)
            if template is None:
                raise FileNotFoundError(f"Template not found: {path}")

            # Split screenshot and template into R, G, B channels
            screenshotR, screenshotG, screenshotB = cv2.split(screenshot)
            templateR, templateG, templateB = cv2.split(template)

            # Perform template matching for each channel
            resultB = cv2.matchTemplate(screenshotB, templateB, cv2.TM_SQDIFF_NORMED)
            resultG = cv2.matchTemplate(screenshotG, templateG, cv2.TM_SQDIFF_NORMED)
            resultR = cv2.matchTemplate(screenshotR, templateR, cv2.TM_SQDIFF_NORMED)

            # Combine results (normalized for fairness)
            result = (resultB + resultG + resultR) / 3.0

            # Find matching locations
            min_val, _, min_loc, _ = cv2.minMaxLoc(result)
            if min_val <= conf:
                top_left = min_loc
                match_x, match_y = top_left

                # Map to screen coordinates if withPos is True
                if withPos:
                    screen_x = window_rect[0] + match_x
                    screen_y = window_rect[1] + match_y
                    center_x = screen_x + template.shape[1] // 2
                    center_y = screen_y + template.shape[0] // 2

                    print(f"Match found at ({center_x}, {center_y}) with confidence {min_val:.2f}")
                    return center_x, center_y
                return True

        # No match found
        return None

    except Exception as e:
        print(f"Error in realityCheck: {e}")
        return None

    finally:
        # Cleanup
        try:
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(handle, hwndDC)
        except:
            pass

def check_wait_area():
    if defaultPixelCheck(hwnd,614, 843,'#e0cdb8'):
        first_area()

    
def check_turn(clickTime):
    if clickTime == 2:
        time.sleep(1)
        #pass turn 
        clickIT(hwnd,1461, 978, True)
        time.sleep(3)
        while defaultPixelCheck(hwnd,1359, 855,'#ff6600') is not True:
            check_wait_area()
            time.sleep(1) 
        clickTime = 0
    return clickTime

def check_S_hidden():
    if defaultPixelCheck(hwnd,1756, 420,'#514a3c'):
        dX, dY = realityCheck(hwnd,"x.png",0.7, True)
        clickIT(hwnd, dX, dY, True)
        time.sleep(0.05)
        X, Y = realityCheck(hwnd,"secondWeapon.png",0.7, True)
        clickIT(hwnd, X, Y, True)
        clickIT(hwnd, 1508, 886 , True)
        time.sleep(1)
        clickIT(hwnd,1715, 503, True)



def fight_strawberries():
    #first place to stand
    clickIT(hwnd, 1191, 671, True)
    time.sleep(1)
    #ready
    clickIT(hwnd, 1822, 753, True)
    time.sleep(3)
    #second place to stand
    clickIT(hwnd,1341, 646, True)
    try:
        red_strawberrie = realityCheckS(hwnd,"redS.png",0.01, False)    
        if red_strawberrie:
            print("red is visible")
            handle_strawberry_fight("redS.png")        
    except Exception as e:
        print(f"Error in fight_strawberries: {e}")
        
def handle_strawberry_fight(nameS):
    enemy_types = ['yellowS.png','redS.png','whitS.png','greenS.png']
    clickTime = 0
    dX, dY = realityCheck(hwnd,"firstWeapon.png",0.7, True)
    print("first weapon cliked")
    clickIT(hwnd, dX, dY, True)
    time.sleep(0.5)
    dX, dY = realityCheck(hwnd,"main.png",0.7, True)
    clickIT(hwnd, dX, dY, True)
    clickIT(hwnd, 1508, 886 , True)
    time.sleep(1)
    X, Y = realityCheck(hwnd,"secondWeapon.png",0.7, True)
    clickIT(hwnd, X, Y, True)
    clickIT(hwnd, 1508, 886 , True)
    time.sleep(0.5)
    dX, dY = realityCheckS(hwnd,nameS,0.01, True)
    clickIT(hwnd, dX, dY, True)
    clickIT(hwnd, 1508, 886 , True)
    time.sleep(1)
    X, Y = realityCheck(hwnd,"thirtWeapon.png",0.7, True)
    clickIT(hwnd, X, Y, True)
    clickIT(hwnd, 1508, 886 , True)
    time.sleep(1)
    if realityCheckS(hwnd,nameS,0.01, False):
        dX, dY = realityCheckS(hwnd,nameS,0.01, True)
        clickIT(hwnd, dX, dY, True) 
        clickIT(hwnd, 1508, 886 , True)
    clickTime = 2
    clickTime = check_turn(clickTime)  
    for strawberrie in enemy_types*2:
                check_wait_area()
                print("currebt strawbery is",strawberrie)
                while realityCheckS(hwnd,strawberrie,0.01, False):
                    X, Y = realityCheck(hwnd,"secondWeapon.png",0.7, True)
                    print("click in the weapon")
                    clickIT(hwnd, X, Y, True)
                    clickIT(hwnd, 1508, 886 , True)
                    time.sleep(2)
                    check_wait_area()
                    print("should click in the monster")
                    if realityCheckS(hwnd,strawberrie,0.01, False):
                        dX, dY = realityCheckS(hwnd,strawberrie,0.01, True)
                        clickIT(hwnd, dX, dY, True)
                    print("ooooooooooooooooooooohhhhhhhhhhhhhhhhhhhhhh")
                    clickIT(hwnd, 1508, 886 , True)
                    clickTime += 1
                    clickTime = check_turn(clickTime)  


def check_message():   
    ha = gText(hwnd,0,60,555, 991)
    text ="You can't free anymore monsters"
    if text in ha :
        return True
    else:
        return False

finish_4 = False
turn = 0

def first_area():
    global finish_4
    global turn 
    while realityCheck(hwnd,"x.png",0.8, False):
        print("found x")
        dX, dY = realityCheck(hwnd,"x.png",0.7, True)
        clickIT(hwnd, dX, dY, True)
    click_strawberrie = 950, 746 #here the position of the strawberrie
    dX, dY = realityCheck(hwnd,"trash.png",0.7, True)
    clickIT(hwnd, dX, dY, True)
    if finish_4:
        x,y = click_strawberrie
        time.sleep(0.5)
        clickIT(hwnd, x,y, True)
        turn =+ 1
        if turn == 4:
            turn = 0
            finish_4 = False
        time.sleep(1)
        fight_strawberries()
    else:
        #open inventory
        clickIT(hwnd, 1561, 842, True)
        time.sleep(0.5)
        while check_message() is not True:
            doubleClickIT(hwnd,1650, 311)
            time.sleep(0.05)
        finish_4 = True
        dX, dY = realityCheck(hwnd,"x.png",0.7, True)
        clickIT(hwnd, dX, dY, True)
        x,y = click_strawberrie
        time.sleep(0.5)
        clickIT(hwnd, x,y, True)
        time.sleep(1)
        fight_strawberries()

def main():
    while True:
        first_area()
main()