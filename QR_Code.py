import cv2
import tkinter as tk
from PIL import Image, ImageTk
import qrcode

# Function to display the camera feed in the Tkinter window
def update_frame():
    if show_camera:
        ret, frame = cap.read()
        if ret:
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame)
            imgtk = ImageTk.PhotoImage(image=img)
            lbl_video.imgtk = imgtk
            lbl_video.configure(image=imgtk)
    lbl_video.after(10, update_frame)

# Function to toggle between camera feed and QR code
def toggle_display():
    global show_camera
    if show_camera:
        qr_img = qrcode.make("https://example.com")
        img = ImageTk.PhotoImage(qr_img)
        lbl_video.imgtk = img
        lbl_video.configure(image=img)
        btn_toggle.config(text="Show Camera Feed")
    else:
        btn_toggle.config(text="Show QR Code")
    show_camera = not show_camera

# Function to save the current camera frame as an image
def save_frame():
    filename = entry_filename.get()
    file_format = format_var.get()
    if show_camera:
        ret, frame = cap.read()
        if ret:
            path = f"{filename}.{file_format}"
            cv2.imwrite(path, frame)
            print(f"Camera frame saved as '{path}'")
    else:
        print("Camera is not active. Switch to camera feed to save a frame.")

# Function to save the QR code as an image
def save_qr_code():
    filename = entry_filename.get()
    file_format = format_var.get()
    qr_img = qrcode.make("https://example.com")
    path = f"{filename}.{file_format}"
    qr_img.save(path)
    print(f"QR code saved as '{path}'")

# Initialize the main window
root = tk.Tk()
root.title("Camera Feed & QR Code")
root.geometry("640x480")

# Initialize OpenCV video capture (0 is usually the default camera)
cap = cv2.VideoCapture(0)

# Create a label in Tkinter to hold the video frames
lbl_video = tk.Label(root)
lbl_video.pack()

# Create a button to toggle display
btn_toggle = tk.Button(root, text="Show QR Code", command=toggle_display)
btn_toggle.pack()

# Create input for filename
lbl_filename = tk.Label(root, text="Filename:")
lbl_filename.pack()
entry_filename = tk.Entry(root)
entry_filename.insert(0, "output")  # Default filename
entry_filename.pack()

# Create dropdown for file format selection
format_var = tk.StringVar(value="png")  # Default format
lbl_format = tk.Label(root, text="Format:")
lbl_format.pack()
dropdown_format = tk.OptionMenu(root, format_var, "png", "jpg", "bmp")
dropdown_format.pack()

# Create buttons to save the frame and QR code
btn_save_frame = tk.Button(root, text="Save Camera Frame", command=save_frame)
btn_save_frame.pack()
btn_save_qr = tk.Button(root, text="Save QR Code", command=save_qr_code)
btn_save_qr.pack()

# Initialize display mode
show_camera = True

# Start updating frames
update_frame()

# Run the Tkinter main loop
root.mainloop()

# Release the camera when done
cap.release()
cv2.destroyAllWindows()

import os

def save_frame():
    filename = entry_filename.get()
    file_format = format_var.get()
    if show_camera:
        ret, frame = cap.read()
        if ret:
            path = f"{filename}.{file_format}"
            cv2.imwrite(path, frame)
            print(f"Camera frame saved as '{path}'")
            os.startfile(os.path.dirname(path))  # Opens the folder in File Explorer
