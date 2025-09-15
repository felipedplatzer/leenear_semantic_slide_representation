# resources.py
import win32com.client


def get_powerpoint_app():
    print(" Launching PowerPoint...")
    _powerpoint_app = win32com.client.Dispatch("PowerPoint.Application")
    return _powerpoint_app

def rgb_to_int(rgb_tuple):
    red, green, blue = rgb_tuple
    return red + (green * 256) + (blue * 256 * 256)



