import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from pathlib import Path

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}

class Cropper:
    def __init__(self, root):
        self.root = root
        self.root.title("Fast Photo Cropper")
        self.root.geometry("1200x850")
        self.root.minsize(800, 600)

        self.files = []
        self.index = 0
        self.img = None
        self.tk_img = None
        self.scale = 1
        self.crop_w = 3
        self.crop_h = 4
        self.zoom = 1.0
        self.offset_x = 0
        self.offset_y = 0
        self.drag_start = None
        self.output_dir = None

        self.ratio_var = tk.StringVar(value="3:4")
        self.size_var = tk.StringVar(value="600x800")
        self.status_var = tk.StringVar(value="Open a folder to begin")

        self.build_ui()
        self.root.bind("<Left>", lambda e: self.prev())
        self.root.bind("<Right>", lambda e: self.next())
        self.root.bind("<Return>", lambda e: self.save_next())
        self.root.bind("<Escape>", lambda e: self.reset_position())

    def build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")

        ttk.Button(top, text="Open folder (batch)", command=self.open_folder).pack(side="left", padx=3)
        ttk.Button(top, text="Open single photo", command=self.open_single).pack(side="left", padx=3)
        ttk.Label(top, text="Ratio:").pack(side="left", padx=(15,3))
        ratio = ttk.Combobox(top, textvariable=self.ratio_var,
                             values=["1:1","3:4","4:3","2:3","3:2","4:5","5:4","9:16","16:9"],
                             width=7, state="readonly")
        ratio.pack(side="left")
        ratio.bind("<<ComboboxSelected>>", lambda e: self.load_current())

        ttk.Label(top, text="Output size:").pack(side="left", padx=(15,3))
        ttk.Entry(top, textvariable=self.size_var, width=10).pack(side="left")
        ttk.Button(top, text="Output folder", command=self.choose_output).pack(side="left", padx=(15,3))

        self.canvas = tk.Canvas(self.root, bg="#202020", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<ButtonPress-1>", self.drag_begin)
        self.canvas.bind("<B1-Motion>", self.drag_move)
        self.canvas.bind("<ButtonRelease-1>", self.drag_end)
        self.canvas.bind("<MouseWheel>", self.wheel_zoom)
        self.canvas.bind("<Button-4>", lambda e: self.zoom_by(1.1))
        self.canvas.bind("<Button-5>", lambda e: self.zoom_by(0.9))
        self.canvas.bind("<Configure>", lambda e: self.draw())

        bottom = ttk.Frame(self.root, padding=8)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="← Previous", command=self.prev).pack(side="left")
        ttk.Button(bottom, text="Reset (Esc)", command=self.reset_position).pack(side="left", padx=6)
        ttk.Button(bottom, text="Save & Next  (Enter)", command=self.save_next).pack(side="right")
        ttk.Button(bottom, text="Next →", command=self.next).pack(side="right", padx=6)
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left", padx=20)

        help_text = "Drag = position   •   Wheel = zoom   •   Enter = save + next   •   ←/→ = previous/next   •   Esc = reset"
        ttk.Label(self.root, text=help_text, padding=(8,0,8,8)).pack(fill="x")

    def parse_ratio(self):
        a, b = self.ratio_var.get().split(":")
        return float(a), float(b)

    def parse_size(self):
        try:
            a, b = self.size_var.get().lower().replace(" ", "").split("x")
            return int(a), int(b)
        except Exception:
            return (600, 800)

    def open_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return
        self.files = sorted([p for p in Path(folder).iterdir()
                             if p.is_file() and p.suffix.lower() in IMAGE_EXTS],
                            key=lambda p: p.name.lower())
        if not self.files:
            messagebox.showinfo("No images", "No supported images were found.")
            return
        self.index = 0
        if self.output_dir is None:
            self.output_dir = Path(folder) / "cropped"
        self.load_current()

    def open_single(self):
        path = filedialog.askopenfilename(
            title="Choose photo",
            filetypes=[
                ("Images", "*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff"),
                ("All files", "*.*")
            ])
        if not path:
            return
        self.files = [Path(path)]
        self.index = 0
        self.output_dir = Path(path).parent / "cropped"
        self.load_current()

    def choose_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir = Path(folder)

    def load_current(self):
        if not self.files:
            return
        try:
            self.img = Image.open(self.files[self.index]).convert("RGB")
            self.zoom = 1.0
            self.offset_x = 0
            self.offset_y = 0
            self.status_var.set(f"{self.index+1} / {len(self.files)}   {self.files[self.index].name}")
            self.draw()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def crop_aspect(self):
        rw, rh = self.parse_ratio()
        iw, ih = self.img.size
        if iw / ih > rw / rh:
            h = ih
            w = int(h * rw / rh)
        else:
            w = iw
            h = int(w * rh / rw)
        return w, h

    def draw(self):
        self.canvas.delete("all")
        if self.img is None:
            self.canvas.create_text(self.canvas.winfo_width()/2,
                                    self.canvas.winfo_height()/2,
                                    text="Open a folder with photos",
                                    fill="white", font=("Segoe UI", 18))
            return

        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        iw, ih = self.img.size
        base = min(cw / iw, ch / ih) * 0.88
        display_scale = base * self.zoom

        nw, nh = max(1, int(iw*display_scale)), max(1, int(ih*display_scale))
        preview = self.img.resize((nw, nh), Image.Resampling.LANCZOS)
        self.tk_img = ImageTk.PhotoImage(preview)

        cx = cw/2 + self.offset_x
        cy = ch/2 + self.offset_y
        x = cx - nw/2
        y = cy - nh/2
        self.canvas.create_image(x, y, image=self.tk_img, anchor="nw")

        rw, rh = self.parse_ratio()
        max_crop_w = min(cw*0.82, ch*0.82*rw/rh)
        max_crop_h = max_crop_w * rh/rw
        if max_crop_h > ch*0.82:
            max_crop_h = ch*0.82
            max_crop_w = max_crop_h * rw/rh

        left = cw/2 - max_crop_w/2
        top = ch/2 - max_crop_h/2
        right = left + max_crop_w
        bottom = top + max_crop_h

        # Dim outside crop area
        self.canvas.create_rectangle(0,0,cw,top,fill="#000000",stipple="gray50",outline="")
        self.canvas.create_rectangle(0,bottom,cw,ch,fill="#000000",stipple="gray50",outline="")
        self.canvas.create_rectangle(0,top,left,bottom,fill="#000000",stipple="gray50",outline="")
        self.canvas.create_rectangle(right,top,cw,bottom,fill="#000000",stipple="gray50",outline="")
        self.canvas.create_rectangle(left,top,right,bottom,outline="white",width=2)

        # Rule-of-thirds guides
        for f in (1/3, 2/3):
            self.canvas.create_line(left+(right-left)*f, top, left+(right-left)*f, bottom,
                                    fill="white", dash=(3,5))
            self.canvas.create_line(left, top+(bottom-top)*f, right, top+(bottom-top)*f,
                                    fill="white", dash=(3,5))

    def drag_begin(self, event):
        self.drag_start = (event.x, event.y)

    def drag_move(self, event):
        if self.drag_start is None:
            return
        dx = event.x - self.drag_start[0]
        dy = event.y - self.drag_start[1]
        self.offset_x += dx
        self.offset_y += dy
        self.drag_start = (event.x, event.y)
        self.draw()

    def drag_end(self, event):
        self.drag_start = None

    def wheel_zoom(self, event):
        self.zoom_by(1.1 if event.delta > 0 else 0.9)

    def zoom_by(self, factor):
        self.zoom = max(0.2, min(5.0, self.zoom * factor))
        self.draw()

    def reset_position(self):
        self.zoom = 1.0
        self.offset_x = 0
        self.offset_y = 0
        self.draw()

    def save_current(self):
        if self.img is None:
            return
        if self.output_dir is None:
            self.output_dir = self.files[0].parent / "cropped"
        self.output_dir.mkdir(parents=True, exist_ok=True)

        out_w, out_h = self.parse_size()
        rw, rh = self.parse_ratio()
        iw, ih = self.img.size

        # Crop dimensions at the current zoom.
        # The crop frame corresponds to a fixed region of the original image.
        base = min(self.canvas.winfo_width()/iw, self.canvas.winfo_height()/ih) * 0.88
        scale = base * self.zoom

        crop_w, crop_h = self.crop_aspect()
        crop_w = min(iw, int(crop_w / self.zoom))
        crop_h = min(ih, int(crop_h / self.zoom))

        center_x = iw/2 - self.offset_x/scale
        center_y = ih/2 - self.offset_y/scale

        left = int(round(center_x - crop_w/2))
        top = int(round(center_y - crop_h/2))
        left = max(0, min(iw-crop_w, left))
        top = max(0, min(ih-crop_h, top))

        cropped = self.img.crop((left, top, left+crop_w, top+crop_h))
        cropped = cropped.resize((out_w, out_h), Image.Resampling.LANCZOS)

        out = self.output_dir / (self.files[self.index].stem + ".jpg")
        cropped.save(out, "JPEG", quality=95, subsampling=0)

    def save_next(self):
        if not self.files:
            return
        try:
            self.save_current()
            if self.index < len(self.files)-1:
                self.index += 1
                self.load_current()
            elif len(self.files) == 1:
                self.status_var.set(f"Saved: {self.output_dir / (self.files[0].stem + ".jpg")}")
            else:
                messagebox.showinfo("Done", f"All {len(self.files)} photos processed.")
        except Exception as e:
            messagebox.showerror("Save error", str(e))

    def next(self):
        if self.files and self.index < len(self.files)-1:
            self.index += 1
            self.load_current()

    def prev(self):
        if self.files and self.index > 0:
            self.index -= 1
            self.load_current()

if __name__ == "__main__":
    root = tk.Tk()
    Cropper(root)
    root.mainloop()
