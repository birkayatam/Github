"""Teknik bilgisi olmayan kullanicilar icin masaustu arayuz: 3 dosya secip
'Calistir' demeleri yeterli. order_pipeline.run() etrafinda ince bir tkinter
sarmalayici -- yeni bir bagimlilik gerektirmez (tkinter standart kutuphanede
gelir). PyInstaller ile tek bir .exe olarak paketlenmek uzere tasarlandi
(bkz. README, '.exe paketleme' bolumu)."""
from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from .order_pipeline import run as run_order_pipeline


class OrderRecommenderApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Bayi Siparis Onerisi")
        self.root.geometry("720x520")

        self.transactions_path = tk.StringVar()
        self.dealer_targets_path = tk.StringVar()
        self.product_quota_path = tk.StringVar()
        self.out_dir = tk.StringVar()
        self.api_key = tk.StringVar()

        self._build_widgets()

    def _build_widgets(self) -> None:
        pad = {"padx": 10, "pady": 6}

        self._file_row("1. Qlik ham veri dosyasi (bayi x urun x ay):", self.transactions_path, 0, pad)
        self._file_row("2. Bayi hedef dosyasi (Kitap1 vb.):", self.dealer_targets_path, 1, pad)
        self._file_row("3. Urun kota / siparis dosyasi:", self.product_quota_path, 2, pad)
        self._folder_row("Cikti klasoru:", self.out_dir, 3, pad)

        tk.Label(self.root, text="Anthropic API Key (opsiyonel, boş bırakılabilir):").grid(
            row=4, column=0, sticky="w", **pad
        )
        tk.Entry(self.root, textvariable=self.api_key, width=50, show="*").grid(
            row=4, column=1, columnspan=2, sticky="we", **pad
        )

        self.run_button = tk.Button(self.root, text="Calistir", command=self._on_run, bg="#2e7d32", fg="white")
        self.run_button.grid(row=5, column=0, columnspan=3, sticky="we", padx=10, pady=12)

        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.grid(row=6, column=0, columnspan=3, sticky="we", padx=10)

        self.log = scrolledtext.ScrolledText(self.root, height=14, state="disabled")
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew", padx=10, pady=10)

        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(7, weight=1)

    def _file_row(self, label: str, var: tk.StringVar, row: int, pad: dict) -> None:
        tk.Label(self.root, text=label).grid(row=row, column=0, sticky="w", **pad)
        tk.Entry(self.root, textvariable=var, width=50).grid(row=row, column=1, sticky="we", **pad)
        tk.Button(self.root, text="Gozat...", command=lambda: self._browse_file(var)).grid(
            row=row, column=2, sticky="w", **pad
        )

    def _folder_row(self, label: str, var: tk.StringVar, row: int, pad: dict) -> None:
        tk.Label(self.root, text=label).grid(row=row, column=0, sticky="w", **pad)
        tk.Entry(self.root, textvariable=var, width=50).grid(row=row, column=1, sticky="we", **pad)
        tk.Button(self.root, text="Gozat...", command=lambda: self._browse_folder(var)).grid(
            row=row, column=2, sticky="w", **pad
        )

    def _browse_file(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Tum dosyalar", "*.*")])
        if path:
            var.set(path)
            if not self.out_dir.get():
                self.out_dir.set(str(Path(path).parent / "raporlar"))

    def _browse_folder(self, var: tk.StringVar) -> None:
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def _log_line(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def _on_run(self) -> None:
        missing = [
            name
            for name, var in [
                ("Qlik veri dosyasi", self.transactions_path),
                ("Bayi hedef dosyasi", self.dealer_targets_path),
                ("Urun kota dosyasi", self.product_quota_path),
                ("Cikti klasoru", self.out_dir),
            ]
            if not var.get().strip()
        ]
        if missing:
            messagebox.showwarning("Eksik bilgi", "Lutfen once sunlari secin: " + ", ".join(missing))
            return

        if self.api_key.get().strip():
            os.environ["ANTHROPIC_API_KEY"] = self.api_key.get().strip()

        self.run_button.configure(state="disabled", text="Calisiyor...")
        self.progress.start(12)
        self._log_line("Isleniyor, bu birkac dakika surebilir (veri boyutuna gore)...")

        # Degerleri ana thread'de (Tk widget'lariyla ayni thread) okuyup arka
        # plan thread'ine parametre olarak geciriyoruz -- Tk nesnelerine
        # baska bir thread'den erisim guvenli degildir.
        thread = threading.Thread(
            target=self._run_pipeline_thread,
            args=(
                self.transactions_path.get(),
                self.dealer_targets_path.get(),
                self.product_quota_path.get(),
                self.out_dir.get(),
            ),
            daemon=True,
        )
        thread.start()

    def _run_pipeline_thread(
        self, transactions_path: str, dealer_targets_path: str, product_quota_path: str, out_dir: str
    ) -> None:
        try:
            excel_path = run_order_pipeline(
                transactions_path=transactions_path,
                dealer_targets_path=dealer_targets_path,
                product_quota_path=product_quota_path,
                out_dir=out_dir,
            )
            self.root.after(0, self._on_success, excel_path)
        except Exception as exc:  # kullaniciya anlasilir hata goster, cokmesin
            self.root.after(0, self._on_error, str(exc))

    def _on_success(self, excel_path: Path) -> None:
        self.progress.stop()
        self.run_button.configure(state="normal", text="Calistir")
        self._log_line(f"Tamamlandi. Ozet rapor: {excel_path}")
        self._log_line(f"Bayi raporlari: {Path(self.out_dir.get()) / 'bayi_raporlari'}")
        if messagebox.askyesno("Tamamlandi", "Islem tamamlandi. Cikti klasorunu acmak ister misiniz?"):
            self._open_folder(Path(self.out_dir.get()))

    def _on_error(self, message: str) -> None:
        self.progress.stop()
        self.run_button.configure(state="normal", text="Calistir")
        self._log_line(f"HATA: {message}")
        messagebox.showerror("Hata olustu", message)

    def _open_folder(self, path: Path) -> None:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')


def main() -> None:
    root = tk.Tk()
    OrderRecommenderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
