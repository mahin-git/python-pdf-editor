"""
PDF Editor — Refactored
Originally: Maahin Deen 
Refactor: Assistant

Features:
- Clean, modular structure (PDFEditorApp class + helper functions)
- Robust TTS worker that keeps one pyttsx3 engine instance
- Image selection & replacement on pages
- Draw-to-read region with TTS
- Theme support (light/dark)
- Safer resource cleanup
- Optional automatic package-check (non-blocking, prints instructions)

Requirements:
  pip install PyMuPDF Pillow pyttsx3

Tested with: Python 3.8+ (works on 3.12)

To run: python pdf_editor_refactor.py
"""

import os
import sys
import threading
import queue
import traceback
import subprocess
import time
from dataclasses import dataclass
from typing import Optional, Tuple, List

import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz  # PyMuPDF
import pyttsx3


# ----------------------------- Utilities ---------------------------------

def ensure_packages_installed():
    """Non-blocking check for required packages; prints instructions if missing."""
    missing = []
    try:
        import fitz  # noqa: F401
    except Exception:
        missing.append("PyMuPDF")
    try:
        from PIL import Image  # noqa: F401
    except Exception:
        missing.append("Pillow")
    try:
        import pyttsx3  # noqa: F401
    except Exception:
        missing.append("pyttsx3")

    if missing:
        print("Missing packages detected:", missing)
        print("Install them with:")
        print(f"    {sys.executable} -m pip install {' '.join(missing)}")


# ----------------------------- TTS Worker --------------------------------

class TTSWorker(threading.Thread):
    """Background worker that owns a single pyttsx3 engine instance.

    It reads messages from a queue and speaks them. A stop_event is used to
    request immediate termination; a sentinel None will also stop the loop.
    """

    def __init__(self, text_queue: queue.Queue, stop_event: threading.Event):
        super().__init__(daemon=True)
        self.queue = text_queue
        self.stop_event = stop_event
        self.engine = None

    def run(self):
        try:
            self.engine = pyttsx3.init()
            self.engine.setProperty("rate", 150)
        except Exception as e:
            print("TTS initialization failed:", e)
            return

        print("TTS worker started")
        while not self.stop_event.is_set():
            try:
                text = self.queue.get(timeout=0.2)
            except queue.Empty:
                continue

            if text is None:
                # sentinel to stop immediately
                break

            try:
                # engine.say + runAndWait is blocking — ok inside worker thread
                self.engine.say(text)
                self.engine.runAndWait()
            except Exception as e:
                print("TTS speak error:", e)
                traceback.print_exc()

        # cleanup
        try:
            if self.engine:
                self.engine.stop()
        except Exception:
            pass
        print("TTS worker exiting")

    def stop(self):
        self.stop_event.set()
        # push sentinel to unblock queue
        try:
            self.queue.put_nowait(None)
        except Exception:
            pass


# ----------------------------- Data classes -------------------------------

@dataclass
class ImageModification:
    page_num: int
    old_rect: fitz.Rect
    new_image_path: str


# ----------------------------- Main App ----------------------------------

class PDFEditorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF Editor — Refactor")
        self.root.geometry("1000x720")

        # PDF state
        self.doc: Optional[fitz.Document] = None
        self.doc_path: Optional[str] = None
        self.current_page_idx: int = 0
        self.zoom = 1.5

        # UI images must be kept referenced to avoid GC
        self._canvas_images = []

        # edit tracking
        self.modifications: List[ImageModification] = []

        # selection state
        self.selected_rect_pdf: Optional[fitz.Rect] = None
        self.selected_rect_canvas: Optional[Tuple[int, int, int, int]] = None
        self.selected_highlight_id: Optional[int] = None

        # drawing (read) state
        self.draw_start: Optional[Tuple[float, float]] = None
        self.draw_rect_id: Optional[int] = None

        # TTS
        self.tts_queue = queue.Queue()
        self.tts_stop_event = threading.Event()
        self.tts_worker: Optional[TTSWorker] = None
        self.tts_ready = False

        # Theme
        self.themes = self._default_themes()
        self.theme = "light"

        # Build UI
        self._build_ui()
        self._apply_theme()

        # start TTS worker in background (non-fatal)
        try:
            self.tts_worker = TTSWorker(self.tts_queue, self.tts_stop_event)
            self.tts_worker.start()
            self.tts_ready = True
        except Exception as e:
            print("Failed to start TTS worker:", e)
            self.tts_ready = False

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------- UI construction ----------------

    def _default_themes(self):
        return {
            "light": {
                "bg": "#f6f6f6",
                "fg": "#222",
                "button_bg": "#e0e0e0",
                "canvas_bg": "#ddd"
            },
            "dark": {
                "bg": "#2b2b2b",
                "fg": "#f1f1f1",
                "button_bg": "#3a3a3a",
                "canvas_bg": "#444"
            }
        }

    def _build_ui(self):
        top = tk.Frame(self.root)
        top.pack(side="top", fill="x", padx=8, pady=6)

        tk.Button(top, text="Open PDF", command=self.open_pdf).pack(side="left")
        tk.Button(top, text="Prev", command=self.prev_page).pack(side="left", padx=4)
        tk.Button(top, text="Next", command=self.next_page).pack(side="left")

        self.page_label = tk.Label(top, text="Page: 0/0")
        self.page_label.pack(side="left", padx=8)

        tk.Button(top, text="Light", command=lambda: self.set_theme("light")).pack(side="right")
        tk.Button(top, text="Dark", command=lambda: self.set_theme("dark")).pack(side="right", padx=4)

        mid = tk.Frame(self.root)
        mid.pack(side="top", fill="x", padx=8, pady=6)

        self.mode_var = tk.StringVar(value="select")
        tk.Radiobutton(mid, text="Select Image", variable=self.mode_var, value="select", command=self._bind_canvas).pack(side="left")
        tk.Radiobutton(mid, text="Draw & Read", variable=self.mode_var, value="draw", command=self._bind_canvas, state=("normal" if self.tts_ready else "disabled")).pack(side="left")

        tk.Button(mid, text="Replace Image", command=self.replace_image, bg="#2e7d32", fg="white").pack(side="right")
        tk.Button(mid, text="Save As...", command=self.save_pdf, bg="#1565c0", fg="white").pack(side="right", padx=6)
        tk.Button(mid, text="Clear Page Edits", command=self.clear_page_edits, bg="#f57c00", fg="white").pack(side="right", padx=6)

        self.status_label = tk.Label(self.root, text="Status: Ready", anchor="w")
        self.status_label.pack(side="bottom", fill="x", padx=6, pady=4)

        # Canvas area with scrollbars
        canvas_frame = tk.Frame(self.root)
        canvas_frame.pack(side="top", fill="both", expand=True, padx=6, pady=6)

        self.canvas = tk.Canvas(canvas_frame, bg=self.themes[self.theme]["canvas_bg"])
        self.canvas.pack(side="left", fill="both", expand=True)

        vbar = tk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        vbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=vbar.set)

        hbar = tk.Scrollbar(self.root, orient="horizontal", command=self.canvas.xview)
        hbar.pack(side="bottom", fill="x")
        self.canvas.configure(xscrollcommand=hbar.set)

        # bind canvas actions (mode-dependent)
        self._bind_canvas()

    def _apply_theme(self):
        t = self.themes[self.theme]
        self.root.config(bg=t["bg"]) 
        self.canvas.config(bg=t["canvas_bg"]) 
        self.status_label.config(bg=t["bg"], fg=t["fg"]) 
        self.page_label.config(bg=t["bg"], fg=t["fg"]) 

    def set_theme(self, name: str):
        if name in self.themes:
            self.theme = name
            self._apply_theme()

    # ---------------- Canvas bindings ----------------

    def _bind_canvas(self):
        # Clear any temp state
        self.canvas.unbind("<Button-1>")
        self.canvas.unbind("<B1-Motion>")
        self.canvas.unbind("<ButtonRelease-1>")

        mode = self.mode_var.get()
        if mode == "select":
            self.canvas.bind("<Button-1>", self.on_canvas_click_select)
        else:
            self.canvas.bind("<Button-1>", self.on_draw_start)
            self.canvas.bind("<B1-Motion>", self.on_draw_motion)
            self.canvas.bind("<ButtonRelease-1>", self.on_draw_end)

    # ---------------- PDF actions ----------------

    def open_pdf(self):
        path = filedialog.askopenfilename(title="Open PDF", filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        try:
            if self.doc:
                self.doc.close()
            self.doc = fitz.open(path)
            self.doc_path = path
            self.current_page_idx = 0
            self.modifications.clear()
            self._render_page()
            self._set_status(f"Opened: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Open PDF", f"Failed to open: {e}")
            self._set_status("Failed to open PDF")

    def _render_page(self):
        self.canvas.delete("all")
        self._canvas_images.clear()
        if not self.doc:
            self.page_label.config(text="Page: 0/0")
            return

        if not (0 <= self.current_page_idx < len(self.doc)):
            self._set_status("Invalid page index")
            return

        page = self.doc[self.current_page_idx]
        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        tk_img = ImageTk.PhotoImage(img)
        self._canvas_images.append(tk_img)
        self.canvas.config(width=img.width, height=img.height)
        self.canvas.create_image(0, 0, image=tk_img, anchor="nw")
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

        # re-draw any replacements applied in memory
        for mod in self.modifications:
            if mod.page_num != self.current_page_idx:
                continue
            # convert PDF rect to canvas coords
            x0 = int(mod.old_rect.x0 * self.zoom)
            y0 = int(mod.old_rect.y0 * self.zoom)
            x1 = int(mod.old_rect.x1 * self.zoom)
            y1 = int(mod.old_rect.y1 * self.zoom)
            try:
                new_pil = Image.open(mod.new_image_path)
                new_pil = new_pil.resize((x1 - x0, y1 - y0), Image.LANCZOS)
                new_tk = ImageTk.PhotoImage(new_pil)
                self._canvas_images.append(new_tk)
                self.canvas.create_image(x0, y0, image=new_tk, anchor="nw")
            except Exception as e:
                print("Could not render replacement image:", e)
                self.canvas.create_rectangle(x0, y0, x1, y1, fill="white")

        self.page_label.config(text=f"Page: {self.current_page_idx + 1}/{len(self.doc)}")
        self._set_status(f"Displayed page {self.current_page_idx + 1}")

    def prev_page(self):
        if not self.doc:
            return
        if self.current_page_idx > 0:
            self.current_page_idx -= 1
            self._clear_selection()
            self._render_page()

    def next_page(self):
        if not self.doc:
            return
        if self.current_page_idx < len(self.doc) - 1:
            self.current_page_idx += 1
            self._clear_selection()
            self._render_page()

    def _set_status(self, text: str):
        self.status_label.config(text=f"Status: {text}")

    # ---------------- Selection & Replace ----------------

    def on_canvas_click_select(self, event):
        if not self.doc:
            self._set_status("Open a PDF first")
            return

        # clear previous highlight
        self._clear_selection()

        # convert scrolled coords
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        pdf_x = cx / self.zoom
        pdf_y = cy / self.zoom

        page = self.doc[self.current_page_idx]
        images = page.get_image_info()
        if not images:
            self._set_status("No images detected on this page")
            return

        # small buffer for ease of click
        buffer_pdf = 8 / self.zoom

        for info in images:
            bbox = info.get("bbox")
            if not bbox:
                continue
            rect = fitz.Rect(bbox)
            rect_buffered = fitz.Rect(rect.x0 - buffer_pdf, rect.y0 - buffer_pdf, rect.x1 + buffer_pdf, rect.y1 + buffer_pdf)
            if rect_buffered.contains((pdf_x, pdf_y)):
                self.selected_rect_pdf = rect
                x0 = int(rect.x0 * self.zoom)
                y0 = int(rect.y0 * self.zoom)
                x1 = int(rect.x1 * self.zoom)
                y1 = int(rect.y1 * self.zoom)
                self.selected_rect_canvas = (x0, y0, x1, y1)
                self.selected_highlight_id = self.canvas.create_rectangle(x0, y0, x1, y1, outline="blue", width=3)
                self._set_status(f"Selected image at {rect}")
                return

        self._set_status("No image at clicked location")

    def replace_image(self):
        if self.mode_var.get() != "select":
            messagebox.showwarning("Mode", "Switch to 'Select Image' mode to replace images.")
            return
        if not self.selected_rect_pdf:
            messagebox.showwarning("No selection", "Click an image on the page to select it first.")
            return

        path = filedialog.askopenfilename(title="Choose replacement image", filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.gif")])
        if not path:
            self._set_status("Image replacement cancelled")
            return

        mod = ImageModification(self.current_page_idx, self.selected_rect_pdf, path)
        # replace if same rect & page
        replaced = False
        for i, existing in enumerate(self.modifications):
            if existing.page_num == mod.page_num and existing.old_rect == mod.old_rect:
                self.modifications[i] = mod
                replaced = True
                break
        if not replaced:
            self.modifications.append(mod)

        self._set_status(f"Scheduled replacement: {os.path.basename(path)}")
        self._render_page()

    def clear_page_edits(self):
        before = len(self.modifications)
        self.modifications = [m for m in self.modifications if m.page_num != self.current_page_idx]
        after = len(self.modifications)
        self._set_status(f"Cleared {before - after} edits on page {self.current_page_idx + 1}")
        self._render_page()

    def save_pdf(self):
        if not self.doc:
            messagebox.showwarning("No PDF", "Open a PDF first")
            return
        if not self.modifications:
            messagebox.showinfo("No changes", "There are no image replacements to save.")
            return

        out_path = filedialog.asksaveasfilename(title="Save modified PDF", defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not out_path:
            return

        self._set_status("Saving PDF...")
        try:
            # work on a copy to avoid modifying original in memory
            new_doc = fitz.open(self.doc_path)
            for mod in self.modifications:
                if not (0 <= mod.page_num < len(new_doc)):
                    continue
                p = new_doc[mod.page_num]
                # white-out old area
                p.draw_rect(mod.old_rect, color=(1, 1, 1), fill=(1, 1, 1))
                # insert new image
                p.insert_image(mod.old_rect, filename=mod.new_image_path)

            new_doc.save(out_path)
            new_doc.close()
            messagebox.showinfo("Saved", f"Modified PDF saved to:\n{out_path}")
            self._set_status(f"Saved to {os.path.basename(out_path)}")
            # reopen saved doc
            if self.doc:
                self.doc.close()
            self.doc = fitz.open(out_path)
            self.doc_path = out_path
            self.current_page_idx = 0
            self.modifications.clear()
            self._render_page()
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save PDF: {e}\n{traceback.format_exc()}")
            self._set_status("Save failed")

    # ---------------- Draw-to-read (TTS) ----------------

    def on_draw_start(self, event):
        if not self.doc:
            self._set_status("Open a PDF first")
            return
        if not self.tts_ready:
            self._set_status("TTS not available")
            return

        self.draw_start = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if self.draw_rect_id:
            self.canvas.delete(self.draw_rect_id)
        x, y = self.draw_start
        self.draw_rect_id = self.canvas.create_rectangle(x, y, x + 1, y + 1, outline="red", width=2)
        self._set_status("Drawing region to read")

    def on_draw_motion(self, event):
        if not self.draw_start:
            return
        x0, y0 = self.draw_start
        x1, y1 = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if self.draw_rect_id:
            self.canvas.coords(self.draw_rect_id, x0, y0, x1, y1)

    def on_draw_end(self, event):
        if not self.draw_start:
            return
        x0, y0 = self.draw_start
        x1, y1 = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.draw_start = None
        if self.draw_rect_id:
            self.canvas.delete(self.draw_rect_id)
            self.draw_rect_id = None

        rect_pdf = fitz.Rect(min(x0, x1) / self.zoom, min(y0, y1) / self.zoom, max(x0, x1) / self.zoom, max(y0, y1) / self.zoom)
        self._read_text_region(rect_pdf)

    def _read_text_region(self, rect: fitz.Rect):
        if not self.doc:
            return
        try:
            page = self.doc[self.current_page_idx]
            txt = page.get_text(clip=rect).strip()
            if not txt:
                self._set_status("No text found in selection")
                return
            self._set_status("Queuing text for speech")
            self.tts_queue.put(txt)
        except Exception as e:
            print("Read error:", e)
            self._set_status("Could not extract text")

    def stop_speaking(self):
        if not self.tts_ready:
            return
        # clear existing queue
        while not self.tts_queue.empty():
            try:
                self.tts_queue.get_nowait()
            except queue.Empty:
                break
        # stop engine via worker stop (engine.stop() will be called inside worker cleanup)
        self.tts_stop_event.set()
        # restart worker so user can resume TTS later
        try:
            if self.tts_worker and self.tts_worker.is_alive():
                self.tts_worker.join(timeout=0.5)
        except Exception:
            pass
        # create a fresh worker
        self.tts_stop_event.clear()
        self.tts_worker = TTSWorker(self.tts_queue, self.tts_stop_event)
        self.tts_worker.start()
        self._set_status("Speech stopped")

    # ---------------- Helpers ----------------

    def _clear_selection(self):
        if self.selected_highlight_id:
            self.canvas.delete(self.selected_highlight_id)
        self.selected_rect_pdf = None
        self.selected_rect_canvas = None
        self.selected_highlight_id = None

    def on_close(self):
        # stop TTS worker
        try:
            if self.tts_worker:
                self.tts_worker.stop()
                self.tts_worker.join(timeout=1.0)
        except Exception:
            pass
        try:
            if self.doc:
                self.doc.close()
        except Exception:
            pass
        self.root.destroy()


# ----------------------------- Entry point --------------------------------

if __name__ == "__main__":
    print("PDF Editor (refactor) — Python", sys.version.split()[0])
    ensure_packages_installed()

    root = tk.Tk()
    app = PDFEditorApp(root)
    root.mainloop()
    print("Exiting...")