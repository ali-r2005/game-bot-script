import win32gui
import win32api
import win32con
import time
import ctypes
from ctypes import windll, Structure, c_long, byref
from ctypes.wintypes import POINT

class CURSORINFO(Structure):
    _fields_ = [
        ("cbSize", c_long),
        ("flags", c_long),
        ("hCursor", c_long),
        ("ptScreenPos", POINT)
    ]

class CursorTracker:
    def __init__(self):
        self.user32 = ctypes.windll.user32

    def get_cursor_pos_multiple_methods(self):
        """
        Get cursor position using multiple methods for cross-verification
        """
        methods = {
            'win32api': self.get_cursor_pos_win32api,
            'user32': self.get_cursor_pos_user32,
            'ctypes_point': self.get_cursor_pos_ctypes
        }
        
        results = {}
        for name, method in methods.items():
            try:
                results[name] = method()
            except Exception as e:
                results[name] = f"Error: {e}"
        
        return results

    def get_cursor_pos_win32api(self):
        """
        Get cursor position using win32api
        """
        return win32api.GetCursorPos()

    def get_cursor_pos_user32(self):
        """
        Get cursor position using user32.GetCursorPos
        """
        pt = POINT()
        self.user32.GetCursorPos(byref(pt))
        return (pt.x, pt.y)

    def get_cursor_pos_ctypes(self):
        """
        Get cursor position using ctypes
        """
        cursor_info = CURSORINFO()
        cursor_info.cbSize = ctypes.sizeof(cursor_info)
        self.user32.GetCursorInfo(byref(cursor_info))
        return (cursor_info.ptScreenPos.x, cursor_info.ptScreenPos.y)

    def get_window_under_cursor(self):
        """
        Get the window handle under the current cursor position
        """
        point = POINT()
        self.user32.GetCursorPos(byref(point))
        return win32gui.WindowFromPoint((point.x, point.y))

    def get_detailed_cursor_info(self):
        """
        Get detailed cursor information
        """
        cursor_info = CURSORINFO()
        cursor_info.cbSize = ctypes.sizeof(cursor_info)
        self.user32.GetCursorInfo(byref(cursor_info))
        
        return {
            'position': (cursor_info.ptScreenPos.x, cursor_info.ptScreenPos.y),
            'flags': cursor_info.flags,  # 0 = hidden, 1 = showing
            'cursor_handle': cursor_info.hCursor
        }

    def track_cursor_movement(self, duration=10, interval=0.5):
        """
        Track cursor movement over a specified duration
        """
        print("Starting cursor tracking...")
        start_time = time.time()
        movements = []

        while time.time() - start_time < duration:
            pos = self.get_cursor_pos_win32api()
            movements.append({
                'time': time.time() - start_time,
                'position': pos
            })
            time.sleep(interval)

        return movements

    def get_window_relative_coordinates(self, handle=None):
        """
        Get cursor coordinates relative to a specific window
        """
        if handle is None:
            handle = self.get_window_under_cursor()

        # Get screen and window coordinates
        screen_pos = self.get_cursor_pos_win32api()
        window_rect = win32gui.GetWindowRect(handle)

        # Calculate relative coordinates
        relative_x = screen_pos[0] - window_rect[0]
        relative_y = screen_pos[1] - window_rect[1]

        return {
            'screen_pos': screen_pos,
            'window_handle': handle,
            'window_rect': window_rect,
            'relative_pos': (relative_x, relative_y)
        }

    def get_cursor_pixel_color(self):
        """
        Get the color of the pixel under the cursor position
        """
        # Get the device context of the entire screen
        dc = win32gui.GetDC(0)
        
        # Get the cursor position
        x, y = self.get_cursor_pos_win32api()
        
        # Get the color of the pixel
        color_ref = win32gui.GetPixel(dc, x, y)
        
        # Release the device context
        win32gui.ReleaseDC(0, dc)
        
        # Convert color_ref to RGB
        red = color_ref & 0xFF
        green = (color_ref >> 8) & 0xFF
        blue = (color_ref >> 16) & 0xFF
        
        # Convert to hex color code
        hex_color = f'#{red:02x}{green:02x}{blue:02x}'
        
        return {
            'rgb': (red, green, blue),
            'hex': hex_color
        }

def main():
    tracker = CursorTracker()

    while True:
        print("\n--- Cursor Detection Menu ---")
        print("1. Get Current Cursor Position")
        print("2. Get Window Under Cursor")
        print("3. Get Detailed Cursor Info")
        print("4. Track Cursor Movement")
        print("5. Get Window Relative Coordinates")
        print("6. Get Pixel Color Under Cursor")
        print("7. Exit")

        choice = input("Enter your choice (1-7): ")

        if choice == '1':
            print("\nCursor Positions:")
            positions = tracker.get_cursor_pos_multiple_methods()
            for method, pos in positions.items():
                print(f"{method}: {pos}")

        elif choice == '2':
            window_handle = tracker.get_window_under_cursor()
            window_title = win32gui.GetWindowText(window_handle)
            print(f"\nWindow Handle: {window_handle}")
            print(f"Window Title: {window_title}")

        elif choice == '3':
            cursor_info = tracker.get_detailed_cursor_info()
            print("\nCursor Information:")
            for key, value in cursor_info.items():
                print(f"{key}: {value}")

        elif choice == '4':
            duration = int(input("Enter tracking duration (seconds): "))
            interval = float(input("Enter tracking interval (seconds): "))
            movements = tracker.track_cursor_movement(duration, interval)
            
            print("\nCursor Movements:")
            for move in movements:
                print(f"Time: {move['time']:.2f}s, Position: {move['position']}")

        elif choice == '5':
            handle = win32gui.FindWindow(None, input("Enter window title (or press enter for current window): "))
            if handle == 0:
                handle = None
            
            relative_info = tracker.get_window_relative_coordinates(handle)
            print("\nWindow Relative Coordinates:")
            for key, value in relative_info.items():
                print(f"{key}: {value}")

        elif choice == '6':
            color_info = tracker.get_cursor_pixel_color()
            print("\nPixel Color Under Cursor:")
            print(f"RGB: {color_info['rgb']}")
            print(f"Hex: {color_info['hex']}")

        elif choice == '7':
            break

        else:
            print("Invalid choice. Please try again.")

        input("\nPress Enter to continue...")

if __name__ == "__main__":
    main()