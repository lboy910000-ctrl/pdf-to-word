"""Desktop GUI (Tkinter) + CLI.   GUI: python app.py
CLI: python app.py file.pdf [-o out.docx] [--ocr] [--latex] [--pages-per-file 200] [--workers N]"""
import argparse
import multiprocessing
import sys
import threading
import time
from pathlib import Path

from extractor import extract


def cli(args):
    out = Path(args.output) if args.output else Path(args.pdf).with_name(Path(args.pdf).stem + "_extracted.docx")
    start = time.time()
    paths, notes = extract(args.pdf, out, force_ocr=args.ocr, latex=args.latex,
                           pages_per_file=args.pages_per_file, workers=args.workers,
                           progress=lambda a, b, m: print(m, flush=True) if a % 10 == 0 or a == b else None)
    for n in notes:
        print("NOTE:", n)
    print(f"Done in {time.time() - start:.0f}s ->")
    for p in paths:
        print("  ", p)


def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Offline PDF to Word (বাংলা / हिन्दी / English)")
    root.geometry("700x520")
    pdf_var, out_var = tk.StringVar(), tk.StringVar()
    ocr_var, tex_var = tk.BooleanVar(value=False), tk.BooleanVar(value=False)
    ppf_var = tk.IntVar(value=200)
    cancel_ev = threading.Event()

    def pick_pdf():
        f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if f:
            pdf_var.set(f)
            out_var.set(str(Path(f).with_name(Path(f).stem + "_extracted.docx")))
            mb = Path(f).stat().st_size / 1e6
            log.insert("end", f"Selected {Path(f).name} ({mb:.0f} MB)\n")

    def pick_out():
        f = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word document", "*.docx")])
        if f:
            out_var.set(f)

    log = tk.Text(root, height=12)
    bar = ttk.Progressbar(root, maximum=100)
    status = tk.Label(root, text="", anchor="w")

    def run():
        if not pdf_var.get():
            return messagebox.showwarning("Select PDF", "Choose a PDF first.")
        btn.config(state="disabled")
        cancel_btn.config(state="normal")
        cancel_ev.clear()
        log.delete("1.0", "end")
        t0 = time.time()

        def work():
            try:
                def prog(a, b, m):
                    bar["value"] = a * 100 / b
                    el = time.time() - t0
                    eta = el / a * (b - a) if a > 1 else 0
                    status.config(text=f"{m}    elapsed {el / 60:.1f} min    remaining about {eta / 60:.1f} min")
                paths, notes = extract(pdf_var.get(), out_var.get(), force_ocr=ocr_var.get(),
                                       latex=tex_var.get(), pages_per_file=max(10, ppf_var.get()),
                                       progress=prog, cancel=cancel_ev.is_set)
                for n in notes:
                    log.insert("end", "NOTE: " + n + "\n")
                log.insert("end", "\nDone. Word file(s):\n" + "\n".join(str(p) for p in paths) + "\n")
            except Exception as e:
                log.insert("end", f"\nERROR: {e}\n")
            btn.config(state="normal")
            cancel_btn.config(state="disabled")
        threading.Thread(target=work, daemon=True).start()

    for label, var, cmd in (("PDF file", pdf_var, pick_pdf), ("Save Word as", out_var, pick_out)):
        f = tk.Frame(root); f.pack(fill="x", padx=10, pady=4)
        tk.Label(f, text=label, width=12, anchor="w").pack(side="left")
        tk.Entry(f, textvariable=var).pack(side="left", fill="x", expand=True)
        tk.Button(f, text="Browse", command=cmd).pack(side="left", padx=4)
    tk.Checkbutton(root, text="Force OCR on every page (scanned PDFs / broken Bengali-Hindi fonts)", variable=ocr_var).pack(anchor="w", padx=10)
    tk.Checkbutton(root, text="Also add LaTeX under equations (needs pix2tex, slower)", variable=tex_var).pack(anchor="w", padx=10)
    f = tk.Frame(root); f.pack(anchor="w", padx=10, pady=2)
    tk.Label(f, text="Pages per Word file (big PDFs are split):").pack(side="left")
    tk.Spinbox(f, from_=10, to=2000, increment=50, textvariable=ppf_var, width=6).pack(side="left", padx=6)
    b = tk.Frame(root); b.pack(pady=6)
    btn = tk.Button(b, text="Convert to Word", command=run, height=2); btn.pack(side="left", padx=6)
    cancel_btn = tk.Button(b, text="Stop (keep what is done)", command=cancel_ev.set, state="disabled", height=2)
    cancel_btn.pack(side="left", padx=6)
    bar.pack(fill="x", padx=10)
    status.pack(fill="x", padx=10)
    log.pack(fill="both", expand=True, padx=10, pady=8)
    root.mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()  # needed for parallel OCR on Windows
    if len(sys.argv) > 1:
        ap = argparse.ArgumentParser()
        ap.add_argument("pdf"); ap.add_argument("-o", "--output")
        ap.add_argument("--ocr", action="store_true"); ap.add_argument("--latex", action="store_true")
        ap.add_argument("--pages-per-file", type=int, default=200)
        ap.add_argument("--workers", type=int, default=None)
        cli(ap.parse_args())
    else:
        gui()
