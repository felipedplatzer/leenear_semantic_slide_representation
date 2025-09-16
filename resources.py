# resources.py
import win32com.client
import pythoncom

def get_powerpoint_app():
    print(" Launching PowerPoint...")
    _powerpoint_app = win32com.client.Dispatch("PowerPoint.Application")
    _powerpoint_app.Visible = True

    return _powerpoint_app

def rgb_to_int(rgb_tuple):
    red, green, blue = rgb_tuple
    return red + (green * 256) + (blue * 256 * 256)



