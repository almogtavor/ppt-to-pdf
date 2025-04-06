import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
from main import process_file, process_directory

class PDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT/PDF to Multi-Slide PDF Converter")
        self.root.geometry("600x500")
        
        # Configure style
        self.style = ttk.Style()
        self.style.configure('TButton', padding=5)
        self.style.configure('TLabel', padding=5)
        self.style.configure('TEntry', padding=5)
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create drag and drop area
        self.create_drag_drop_area()
        
        # Create settings frame
        self.create_settings_frame()
        
        # Create output directory frame
        self.create_output_frame()
        
        # Create progress bar
        self.create_progress_bar()
        
        # Create convert button
        self.create_convert_button()
        
        # Initialize variables
        self.input_paths = []
        self.output_dir = ""
        self.processing = False

    def create_drag_drop_area(self):
        frame = ttk.LabelFrame(self.main_frame, text="Drag and Drop Files Here", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.drop_area = tk.Text(frame, height=5, wrap=tk.WORD)
        self.drop_area.pack(fill=tk.BOTH, expand=True)
        self.drop_area.insert(tk.END, "Drop files or folders here...")
        self.drop_area.config(state=tk.DISABLED)
        
        # Configure drag and drop
        self.drop_area.drop_target_register(tk.DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.handle_drop)
        
        # Add browse button
        browse_frame = ttk.Frame(frame)
        browse_frame.pack(fill=tk.X, pady=5)
        ttk.Button(browse_frame, text="Browse Files", command=self.browse_files).pack(side=tk.LEFT)
        ttk.Button(browse_frame, text="Browse Folder", command=self.browse_folder).pack(side=tk.LEFT, padx=5)

    def create_settings_frame(self):
        frame = ttk.LabelFrame(self.main_frame, text="Layout Settings", padding="10")
        frame.pack(fill=tk.X, pady=5)
        
        # Slides per row
        ttk.Label(frame, text="Slides per Row:").grid(row=0, column=0, sticky=tk.W)
        self.slides_per_row = ttk.Spinbox(frame, from_=1, to=6, width=5)
        self.slides_per_row.set(3)
        self.slides_per_row.grid(row=0, column=1, sticky=tk.W)
        
        # Gap
        ttk.Label(frame, text="Gap (points):").grid(row=1, column=0, sticky=tk.W)
        self.gap = ttk.Spinbox(frame, from_=0, to=50, width=5)
        self.gap.set(10)
        self.gap.grid(row=1, column=1, sticky=tk.W)
        
        # Margin
        ttk.Label(frame, text="Margin (points):").grid(row=2, column=0, sticky=tk.W)
        self.margin = ttk.Spinbox(frame, from_=0, to=100, width=5)
        self.margin.set(20)
        self.margin.grid(row=2, column=1, sticky=tk.W)
        
        # Top margin
        ttk.Label(frame, text="Top Margin (points):").grid(row=3, column=0, sticky=tk.W)
        self.top_margin = ttk.Spinbox(frame, from_=0, to=100, width=5)
        self.top_margin.set(0)
        self.top_margin.grid(row=3, column=1, sticky=tk.W)

    def create_output_frame(self):
        frame = ttk.LabelFrame(self.main_frame, text="Output Directory", padding="10")
        frame.pack(fill=tk.X, pady=5)
        
        self.output_path = ttk.Entry(frame)
        self.output_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(frame, text="Browse", command=self.browse_output).pack(side=tk.RIGHT)

    def create_progress_bar(self):
        self.progress = ttk.Progressbar(self.main_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)
        self.status_label = ttk.Label(self.main_frame, text="Ready")
        self.status_label.pack(fill=tk.X)

    def create_convert_button(self):
        ttk.Button(self.main_frame, text="Convert", command=self.start_conversion).pack(pady=10)

    def handle_drop(self, event):
        paths = event.data.split()
        self.input_paths = []
        for path in paths:
            # Remove {} if present (Windows adds them)
            path = path.strip('{}')
            if os.path.exists(path):
                self.input_paths.append(path)
        self.update_drop_area()

    def browse_files(self):
        files = filedialog.askopenfilenames(
            title="Select Files",
            filetypes=[("Supported Files", "*.pdf *.ppt *.pptx")]
        )
        if files:
            self.input_paths.extend(files)
            self.update_drop_area()

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            self.input_paths.append(folder)
            self.update_drop_area()

    def browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_path.delete(0, tk.END)
            self.output_path.insert(0, folder)

    def update_drop_area(self):
        self.drop_area.config(state=tk.NORMAL)
        self.drop_area.delete(1.0, tk.END)
        if self.input_paths:
            self.drop_area.insert(tk.END, "\n".join(self.input_paths))
        else:
            self.drop_area.insert(tk.END, "Drop files or folders here...")
        self.drop_area.config(state=tk.DISABLED)

    def start_conversion(self):
        if not self.input_paths:
            messagebox.showerror("Error", "Please select input files or folders")
            return
        
        if not self.output_path.get():
            messagebox.showerror("Error", "Please select an output directory")
            return
        
        if self.processing:
            messagebox.showerror("Error", "Conversion already in progress")
            return
        
        self.processing = True
        self.progress['value'] = 0
        self.status_label.config(text="Starting conversion...")
        
        # Start conversion in a separate thread
        threading.Thread(target=self.convert_files, daemon=True).start()

    def convert_files(self):
        try:
            total_files = len(self.input_paths)
            for i, path in enumerate(self.input_paths):
                self.status_label.config(text=f"Processing {os.path.basename(path)}...")
                
                if os.path.isfile(path):
                    output_file = os.path.join(
                        self.output_path.get(),
                        os.path.splitext(os.path.basename(path))[0] + '.pdf'
                    )
                    process_file(
                        path,
                        output_file,
                        int(self.slides_per_row.get()),
                        int(self.gap.get()),
                        int(self.margin.get()),
                        int(self.top_margin.get())
                    )
                else:
                    process_directory(
                        path,
                        self.output_path.get(),
                        int(self.slides_per_row.get()),
                        int(self.gap.get()),
                        int(self.margin.get()),
                        int(self.top_margin.get())
                    )
                
                self.progress['value'] = (i + 1) / total_files * 100
                self.root.update_idletasks()
            
            self.status_label.config(text="Conversion completed successfully!")
            messagebox.showinfo("Success", "All files have been converted successfully!")
        
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", str(e))
        
        finally:
            self.processing = False

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterGUI(root)
    root.mainloop() 