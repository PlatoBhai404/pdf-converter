import os
import threading
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Core Engine Pipelines
from converters.pdf_to_all import extract_text_layout, pdf_to_images, pdf_to_docx, pdf_to_html, pdf_to_excel
from converters.all_to_pdf import images_to_pdf, docx_to_pdf, txt_to_pdf, html_to_pdf, pptx_to_pdf

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PDFConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Personal Ultimate Document Converter")
        self.geometry("600x520")
        self.resizable(False, False)
        
        self.selected_files = []
        self.latest_output_path = ""
        
        # --- UI Components ---
        self.title_label = ctk.CTkLabel(self, text="Universal PDF & Document Engine", font=ctk.CTkFont(size=22, weight="bold"))
        self.title_label.pack(pady=20)
        
        self.select_btn = ctk.CTkButton(self, text="Select File(s)", command=self.select_files)
        self.select_btn.pack(pady=10)
        
        self.file_label = ctk.CTkLabel(self, text="No files selected", font=ctk.CTkFont(size=12), text_color="gray", wraplength=500)
        self.file_label.pack(pady=10)
        
        self.format_label = ctk.CTkLabel(self, text="Convert to Target Format:")
        self.format_label.pack(pady=5)
        self.format_dropdown = ctk.CTkOptionMenu(self, values=["Select a file first"])
        self.format_dropdown.pack(pady=5)
        
        self.convert_btn = ctk.CTkButton(self, text="Convert Now", fg_color="green", hover_color="darkgreen", command=self.start_conversion_thread)
        self.convert_btn.pack(pady=20)

        self.status_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12, weight="bold"))
        self.status_label.pack(pady=5)

        self.open_file_btn = ctk.CTkButton(self, text="📂 Open Destination Location", fg_color="#1f538d", command=self.open_file_location)

    def select_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Select document(s)",
            filetypes=[("All Supported Files", "*.pdf;*.docx;*.png;*.jpg;*.jpeg;*.txt;*.html;*.pptx")]
        )
        if file_paths:
            self.selected_files = list(file_paths)
            self.open_file_btn.pack_forget()
            self.status_label.configure(text="")
            
            if len(self.selected_files) == 1:
                self.file_label.configure(text=os.path.basename(self.selected_files[0]), text_color="white")
            else:
                self.file_label.configure(text=f"Selected {len(self.selected_files)} files loaded for compilation.", text_color="cyan")
            
            first_ext = os.path.splitext(self.selected_files[0])[1].lower()
            
            # Smart context configuration based on extension mapping
            if first_ext == ".pdf":
                self.format_dropdown.configure(values=["Word (.docx)", "Excel Tables (.xlsx)", "Web Page (.html)", "Images (.png)", "Plain Text (.txt)"])
                self.format_dropdown.set("Word (.docx)")
            elif first_ext in [".png", ".jpg", ".jpeg"]:
                self.format_dropdown.configure(values=["Compiled PDF (.pdf)"])
                self.format_dropdown.set("Compiled PDF (.pdf)")
            elif first_ext in [".docx", ".txt", ".html", ".pptx"]:
                self.format_dropdown.configure(values=["PDF (.pdf)"])
                self.format_dropdown.set("PDF (.pdf)")

    def start_conversion_thread(self):
        if not self.selected_files:
            messagebox.showwarning("Warning", "Please choose your source files first!")
            return
        self.convert_btn.configure(state="disabled")
        self.open_file_btn.pack_forget()
        self.status_label.configure(text="Processing files via Hub conversion engine...", text_color="orange")
        threading.Thread(target=self.run_conversion, daemon=True).start()

    def run_conversion(self):
        target_format = self.format_dropdown.get()
        primary_file = self.selected_files[0]
        base_path, ext = os.path.splitext(primary_file)
        success = False

        try:
            # 1. Pipeline: PDF -> X
            if ext.lower() == ".pdf":
                if "Word" in target_format:
                    self.latest_output_path = base_path + ".docx"
                    success = pdf_to_docx(primary_file, self.latest_output_path)
                elif "Excel" in target_format:
                    self.latest_output_path = base_path + ".xlsx"
                    success = pdf_to_excel(primary_file, self.latest_output_path)
                elif "Web" in target_format:
                    self.latest_output_path = base_path + ".html"
                    success = pdf_to_html(primary_file, self.latest_output_path)
                elif "Images" in target_format:
                    self.latest_output_path = base_path + "_images"
                    success = pdf_to_images(primary_file, self.latest_output_path)
                elif "Plain Text" in target_format:
                    self.latest_output_path = base_path + ".txt"
                    success = extract_text_layout(primary_file, self.latest_output_path)

            # 2. Pipeline: X -> PDF
            elif ext.lower() in [".png", ".jpg", ".jpeg"]:
                self.latest_output_path = os.path.join(os.path.dirname(primary_file), "compiled_images_output.pdf")
                success = images_to_pdf(self.selected_files, self.latest_output_path)
            elif ext.lower() == ".docx":
                self.latest_output_path = base_path + ".pdf"
                success = docx_to_pdf(primary_file, self.latest_output_path)
            elif ext.lower() == ".txt":
                self.latest_output_path = base_path + ".pdf"
                success = txt_to_pdf(primary_file, self.latest_output_path)
            elif ext.lower() == ".html":
                self.latest_output_path = base_path + ".pdf"
                success = html_to_pdf(primary_file, self.latest_output_path)
            elif ext.lower() == ".pptx":
                self.latest_output_path = base_path + ".pdf"
                success = pptx_to_pdf(primary_file, self.latest_output_path)

            if success:
                self.status_label.configure(text="Success! Conversion complete.", text_color="green")
                self.open_file_btn.pack(pady=10)
                messagebox.showinfo("Success", "Process completed successfully!")
            else:
                self.status_label.configure(text="Conversion engine dropped with errors.", text_color="red")
                messagebox.showerror("Error", "The specific format bridge failed.")
        except Exception as e:
            self.status_label.configure(text="Fatal processing exception.", text_color="red")
            messagebox.showerror("Error", f"Error stack: {str(e)}")
        finally:
            self.convert_btn.configure(state="normal")

    def open_file_location(self):
        if self.latest_output_path and os.path.exists(self.latest_output_path):
            clean_path = os.path.normpath(self.latest_output_path)
            if os.path.isdir(clean_path): os.startfile(clean_path)
            else: subprocess.run(['explorer', '/select,', clean_path])

if __name__ == "__main__":
    app = PDFConverterApp()
    app.mainloop()