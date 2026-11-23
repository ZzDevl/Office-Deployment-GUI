import os
import sys
import subprocess
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from xml.dom import minidom

OFFICE_EDITIONS = {
    "Office LTSC Professional Plus 2024": ("ProPlus2024Volume", "PerpetualVL2024"),
    "Office LTSC Standard 2024": ("Standard2024Volume", "PerpetualVL2024"),
    "Project Professional 2024": ("ProjectPro2024Volume", "PerpetualVL2024"),
    "Visio Professional 2024": ("VisioPro2024Volume", "PerpetualVL2024"),
    "Microsoft 365 Apps for Enterprise": ("O365ProPlusRetail", "Current"),
}

OFFICE_APPS = [
    ("Word", True),
    ("Excel", True),
    ("PowerPoint", True),
    ("Outlook", False),
    ("OneNote", False),
    ("Access", False),
    ("Publisher", False),
    ("Teams", False),
    ("Lync", False),
    ("Groove", False),
]

LANGUAGES = {
    "English (en-us)": "en-us",
    "Bahasa Indonesia (id-id)": "id-id",
}

class OfficeInstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Z Dev | Office Installer v1.0")
        self.root.geometry("560x620")
        self.root.resizable(False, False)

        # Pasang ikon jendela dengan aman
        try:
            if getattr(sys, "frozen", False):
                self.root.iconbitmap(default=sys.executable)
            elif os.path.exists("icon.ico"):
                self.root.iconbitmap("icon.ico")
        except Exception:
            pass

        self.odt_path = tk.StringVar()
        self.edition_var = tk.StringVar(value=list(OFFICE_EDITIONS.keys())[0])
        self.arch_var = tk.StringVar(value="64")
        self.lang_var = tk.StringVar(value="English (en-us)")
        self.remove_msi_var = tk.BooleanVar(value=True)
        self.silent_install_var = tk.BooleanVar(value=False)
        self.app_vars = {}

        # Cari setup.exe di folder yang sama dengan aplikasi/.exe
        app_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd()
        default_setup = os.path.join(app_dir, "setup.exe")
        if os.path.exists(default_setup):
            self.odt_path.set(default_setup)

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        frame_odt = ttk.LabelFrame(self.root, text="1. Office Deployment Tool (setup.exe)")
        frame_odt.pack(fill="x", **pad)

        entry_odt = ttk.Entry(frame_odt, textvariable=self.odt_path, width=48)
        entry_odt.pack(side="left", padx=(8, 4), pady=8)

        ttk.Button(frame_odt, text="Pilih File", command=self.browse_odt).pack(side="left", padx=2)
        ttk.Button(frame_odt, text="Unduh ODT", command=self.open_odt_download).pack(side="left", padx=(2, 8))

        frame_cfg = ttk.LabelFrame(self.root, text="2. Konfigurasi Versi & Arsitektur")
        frame_cfg.pack(fill="x", **pad)

        ttk.Label(frame_cfg, text="Edisi Office:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Combobox(frame_cfg, textvariable=self.edition_var, values=list(OFFICE_EDITIONS.keys()), state="readonly", width=38).grid(row=0, column=1, columnspan=2, sticky="w", padx=8, pady=6)

        ttk.Label(frame_cfg, text="Bahasa:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ttk.Combobox(frame_cfg, textvariable=self.lang_var, values=list(LANGUAGES.keys()), state="readonly", width=25).grid(row=1, column=1, sticky="w", padx=8, pady=6)

        ttk.Label(frame_cfg, text="Arsitektur:").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        frame_arch = ttk.Frame(frame_cfg)
        frame_arch.grid(row=2, column=1, sticky="w", padx=8, pady=6)
        ttk.Radiobutton(frame_arch, text="64-bit (Direkomendasikan)", variable=self.arch_var, value="64").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(frame_arch, text="32-bit", variable=self.arch_var, value="32").pack(side="left")

        frame_apps = ttk.LabelFrame(self.root, text="3. Pilih Aplikasi yang Akan Diinstal")
        frame_apps.pack(fill="x", **pad)

        for idx, (app_name, default_state) in enumerate(OFFICE_APPS):
            var = tk.BooleanVar(value=default_state)
            self.app_vars[app_name] = var
            label_text = "Skype for Business" if app_name == "Lync" else ("OneDrive Legacy" if app_name == "Groove" else app_name)
            cb = ttk.Checkbutton(frame_apps, text=label_text, variable=var)
            cb.grid(row=idx // 3, column=idx % 3, sticky="w", padx=16, pady=4)

        frame_opts = ttk.LabelFrame(self.root, text="4. Opsi Instalasi")
        frame_opts.pack(fill="x", **pad)

        ttk.Checkbutton(frame_opts, text="Hapus versi Office MSI lama otomatis (<RemoveMSI />)", variable=self.remove_msi_var).pack(anchor="w", padx=10, pady=3)
        ttk.Checkbutton(frame_opts, text="Silent Install (Sembunyikan jendela progres instalasi Office)", variable=self.silent_install_var).pack(anchor="w", padx=10, pady=3)

        frame_actions = ttk.Frame(self.root)
        frame_actions.pack(fill="x", padx=12, pady=12)

        ttk.Button(frame_actions, text="Preview / Simpan XML", command=self.save_and_preview_xml).pack(side="left", expand=True, fill="x", padx=4, ipady=6)
        ttk.Button(frame_actions, text="Download Offline", command=lambda: self.run_odt("/download")).pack(side="left", expand=True, fill="x", padx=4, ipady=6)
        ttk.Button(frame_actions, text="Install Office", command=lambda: self.run_odt("/configure")).pack(side="left", expand=True, fill="x", padx=4, ipady=6)

        self.status_var = tk.StringVar(value="Siap. Pastikan setup.exe dari Microsoft ODT sudah dipilih.")
        ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w").pack(side="bottom", fill="x")

    def browse_odt(self):
        path = filedialog.askopenfilename(title="Pilih setup.exe (ODT)", filetypes=[("Executable", "setup.exe"), ("All Files", "*.*")])
        if path:
            self.odt_path.set(path)

    def open_odt_download(self):
        webbrowser.open("https://www.microsoft.com/en-us/download/details.aspx?id=49117")

    def generate_xml_string(self):
        product_id, channel = OFFICE_EDITIONS[self.edition_var.get()]
        arch = self.arch_var.get()
        lang_id = LANGUAGES[self.lang_var.get()]

        config = ET.Element("Configuration")
        add = ET.SubElement(config, "Add", OfficeClientEdition=arch, Channel=channel)
        product = ET.SubElement(add, "Product", ID=product_id)
        ET.SubElement(product, "Language", ID=lang_id)

        for app_id, var in self.app_vars.items():
            if not var.get():
                ET.SubElement(product, "ExcludeApp", ID=app_id)

        if self.remove_msi_var.get():
            ET.SubElement(config, "RemoveMSI")

        display_level = "None" if self.silent_install_var.get() else "Full"
        ET.SubElement(config, "Display", Level=display_level, AcceptEULA="TRUE")

        raw_xml = ET.tostring(config, encoding="utf-8")
        return minidom.parseString(raw_xml).toprettyxml(indent="  ")

    def write_xml_file(self):
        setup_exe = self.odt_path.get().strip()
        work_dir = os.path.dirname(setup_exe) if setup_exe and os.path.exists(setup_exe) else os.getcwd()
        xml_path = os.path.join(work_dir, "configuration.xml")

        xml_content = self.generate_xml_string()
        with open(xml_path, "w", encoding="utf-8") as f:
            f.write(xml_content)
        return xml_path, xml_content

    def save_and_preview_xml(self):
        xml_path, xml_content = self.write_xml_file()
        self.status_var.set(f"Tersimpan: {xml_path}")

        top = tk.Toplevel(self.root)
        top.title("Preview configuration.xml")
        top.geometry("480x360")
        txt = tk.Text(top, wrap="none", font=("Consolas", 10))
        txt.insert("1.0", xml_content)
        txt.pack(expand=True, fill="both", padx=8, pady=8)

    def run_odt(self, mode):
        setup_exe = self.odt_path.get().strip()
        if not setup_exe or not os.path.exists(setup_exe):
            messagebox.showerror("Error", "File setup.exe tidak ditemukan! Unduh dan ekstrak ODT terlebih dahulu.")
            return

        product_id, _ = OFFICE_EDITIONS[self.edition_var.get()]
        if "Project" in product_id or "Visio" in product_id:
            messagebox.showinfo("Info Produk", "Project / Visio akan diinstal sebagai aplikasi tunggal tanpa ExcludeApp paket Office.")

        xml_path, _ = self.write_xml_file()
        work_dir = os.path.dirname(setup_exe)

        action_label = "mengunduh file instalasi" if mode == "/download" else "menginstal Microsoft Office"
        if not messagebox.askyesno("Konfirmasi", f"Mulai {action_label} sekarang?\n\nPerintah:\nsetup.exe {mode} configuration.xml"):
            return

        self.status_var.set(f"Menjalankan {mode}... Jangan tutup aplikasi sampai proses selesai.")
        self.root.update()

        try:
            process = subprocess.Popen([setup_exe, mode, xml_path], cwd=work_dir)
            process.wait()
            self.status_var.set(f"Proses {mode} selesai (Exit Code: {process.returncode}).")
            messagebox.showinfo("Selesai", f"Proses {action_label} telah selesai!")
        except Exception as e:
            self.status_var.set("Terjadi kesalahan saat menjalankan setup.exe.")
            messagebox.showerror("Gagal", f"Gagal menjalankan ODT:\n{e}\n\nPastikan aplikasi dijalankan sebagai Administrator.")

if __name__ == "__main__":
    root = tk.Tk()
    app = OfficeInstallerApp(root)
    root.mainloop()