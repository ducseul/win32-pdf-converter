import tkinter as tk
from tkinter import ttk
import threading
import subprocess

class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Application Runner")
        
        self.entries = []
        self.threads = []

        self.create_widgets()

    def create_widgets(self):
        # Calculate desired window width
        screen_width = self.root.winfo_screenwidth()
        window_width = int(screen_width * 0.5)

        # Set window size and position
        window_height = 400
        x_position = (screen_width - window_width) // 2
        y_position = (self.root.winfo_screenheight() - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # Create frame
        self.frame = ttk.Frame(self.root)
        self.frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.frame.columnconfigure(0, weight=1)
        self.frame.columnconfigure(1, weight=1)

        # Create buttons
        self.add_button = ttk.Button(self.frame, text="Add Row", command=self.add_row)
        self.add_button.grid(row=0, column=0, sticky="e")

        self.start_button = ttk.Button(self.frame, text="Start", command=self.start_threads)
        self.start_button.grid(row=0, column=1, sticky="e")

        self.stop_button = ttk.Button(self.frame, text="Stop", command=self.stop_threads, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=2, sticky="e")

    def add_row(self):
        row = len(self.entries)
        entry1 = ttk.Entry(self.frame, width=50)
        entry2 = ttk.Entry(self.frame, width=60)
        entry1.grid(row=row+1, column=0, padx=(0, 5), pady=5, sticky="ew")
        entry2.grid(row=row+1, column=1, padx=(0, 5), pady=5, sticky="ew")
        entry1.insert(0, "Application Path")
        entry2.insert(0, "Parameters")
        self.entries.append((entry1, entry2))

    def start_threads(self):
        for entry1, entry2 in self.entries:
            app_path = entry1.get()
            parameters = entry2.get()
            if app_path:
                thread = threading.Thread(target=self.run_application, args=(app_path, parameters))
                thread.start()
                self.threads.append(thread)
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

    # def run_application(self, app_path, parameters):
    #     params_list = parameters.split(',')
    #     print([app_path] + params_list)
    #     subprocess.run([app_path] + params_list)
        
    def run_application(self, app_path, parameters):
        params_list = parameters.split()
        subprocess.run([app_path] + params_list)

    def stop_threads(self):
        for thread in self.threads:
            thread.join(timeout=1)
        self.threads.clear()
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

def main():
    root = tk.Tk()
    app = Application(root)
    root.mainloop()

if __name__ == "__main__":
    main()
