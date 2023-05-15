# take screenshots of the powerpoint slides in ms teams
import tkinter as tk
import pyautogui
from PIL import Image

from io import BytesIO
import win32clipboard
from PIL import Image, ImageTk
import os

class GUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("300x300")
        self.root.resizable(False, False)
        self.root.title("Screenshot Slides")
        
        self.img_path = r"path"

        self.button = tk.Button(self.root, text="Click to Screenshot", font=('Arial', 18), command=lambda: [self.screenshot(), self.send_to_clipboard(win32clipboard.CF_DIB, self.img_path), self.show_image()])
        self.button.pack(side=tk.BOTTOM, pady = 30)

        self.picture_label = tk.Label(self.root, bg="#25265E")
        self.picture_label.pack(expand=True, fill="both")

        self.root.mainloop()

    def screenshot(self):
        self.screenshot_img = pyautogui.screenshot(region=(100 , 130, 850, 570))
        self.screenshot_img.save(self.img_path)

        self.send_to_clipboard(win32clipboard.CF_DIB, self.img_path)
        
    
    def send_to_clipboard(self, clip_type, filepath):
        image = Image.open(filepath)

        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()
    
    def show_image(self):
        self.open_image = Image.open(self.img_path)

        self.image = ImageTk.PhotoImage(self.open_image.resize((255, 166)))

        self.picture_label.config(image=self.image)
        


if __name__ == "__main__":
    GUI()