import win32api, pyautogui, PIL, win32gui, win32ui, win32con, os
from PIL import Image

def get_area():
    mouse_state = win32api.GetKeyState(0x01)
    clicked = False
    button_up = False
    
    while True:
        mouse_state = win32api.GetKeyState(0x01)
        
        if clicked is False:
            if mouse_state == -127 or mouse_state == -128:
                first = win32api.GetCursorPos()
                clicked = True
                

        
        if clicked is True:
            if mouse_state == 0 or mouse_state == 1:
                button_up = True
            if button_up is True:
                if mouse_state == -127 or mouse_state == -128:
                    second = win32api.GetCursorPos()
                    break

    top_left = []
    bottom_right = []        

    if first[0] > second[0]:
        top_left.append(second[0])
        bottom_right.append(first[0])
    else:
        top_left.append(first[0])
        bottom_right.append(second[0])

    if first[1] > second[1]:
        top_left.append(second[1])
        bottom_right.append(first[1])
    else:
        top_left.append(first[1])
        bottom_right.append(second[1])


    return (top_left[0], top_left[1], bottom_right[0], bottom_right[1])

def get_screenshot():
    hwin = win32gui.GetDesktopWindow()
    width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
    hwindc = win32gui.GetWindowDC(hwin)
    srcdc = win32ui.CreateDCFromHandle(hwindc)
    memdc = srcdc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(srcdc, width, height)
    memdc.SelectObject(bmp)
    memdc.BitBlt((0, 0), (width, height), srcdc, (left, top), win32con.SRCCOPY)

    bmpinfo = bmp.GetInfo()
    bmpstr = bmp.GetBitmapBits(True)
    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    return im


if __name__ == "__main__":
    im = get_screenshot()
    box = get_area()
    image = im.crop(box)
    image.show()

    print(os.getcwd())

    print("Type file name:")
    name = input()
    name = name.replace(' ', '')
    if name == '':
        name = 'screenshot'
        
    name = os.path.join(os.environ["HOMEPATH"], "Desktop", name + ".png")
    image.save(name)
