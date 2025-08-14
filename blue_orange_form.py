import tkinter as tk
from tkinter import ttk


def create_row(parent, heading, fields):
    row = ttk.Frame(parent)
    row.pack(fill="x", pady=5)

    ttk.Label(row, text=heading).grid(row=0, column=0, padx=5, sticky="w")

    for i, field in enumerate(fields):
        ttk.Label(row, text=field).grid(row=0, column=1 + 2 * i, padx=2, sticky="w")
        ttk.Entry(row, width=10).grid(row=0, column=2 + 2 * i, padx=2)


def build_ui(root):
    root.title("Cuadros Azules y Naranjas")

    main = tk.Frame(root)
    main.pack(fill="both", expand=True)

    canvas = tk.Canvas(main)
    canvas.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(main, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)

    content = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=content, anchor="nw")

    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    content.bind("<Configure>", on_configure)

    # Estilos
    style = ttk.Style()
    style.configure("Blue.TLabelframe", background="#2882c7")
    style.configure("Blue.TLabelframe.Label", background="#2882c7", foreground="white", font=("Arial", 12, "bold"))
    style.configure("Orange.TLabelframe", background="#e98300")
    style.configure("Orange.TLabelframe.Label", background="#e98300", foreground="white", font=("Arial", 12, "bold"))

    # Cuadro Azul
    blue = ttk.Labelframe(content, text="Cuadros AZULES", style="Blue.TLabelframe", padding=10)
    blue.pack(fill="x", pady=10)

    create_row(blue, "Base rating:", ["HP", "kW", "w"])
    create_row(blue, "50 Hz Rating:", ["Required HP", "NEMA HP"])
    create_row(blue, "Ambient Temperature:", ["Valor numérico", "Units (°C)"])
    create_row(blue, "Load %", ["%"])    
    create_row(blue, "New Rating:", ["HP", "kW", "w", "NEMA HP"])
    create_row(blue, "Tolerancia (EC):", ["Required HP", "NEMA HP"])
    create_row(blue, "50 Hz Rating (otra sección):", ["Required HP", "NEMA HP"])
    create_row(blue, "Tolerancia (EC+50):", ["Required HP", "NEMA HP"])

    # Cuadro Naranja
    orange = ttk.Labelframe(content, text="Cuadros NARANJAS", style="Orange.TLabelframe", padding=10)
    orange.pack(fill="x", pady=10)

    create_row(orange, "Base rating:", ["HP", "kW", "w"])
    create_row(orange, "50 Hz Rating:", ["Required HP", "NEMA HP"])
    create_row(orange, "FASL (MASL):", ["Valor numérico", "Units (ft / m)"])
    create_row(orange, "Load %", ["%"])
    create_row(orange, "New Rating:", ["HP", "kW", "w", "NEMA HP"])
    create_row(orange, "Tolerancia (EC):", ["Required HP", "NEMA HP"])
    create_row(orange, "50 Hz Rating:", ["Required HP", "NEMA HP"])
    create_row(orange, "Tolerancia (EC+50):", ["Required HP", "NEMA HP"])

    # Scroll con rueda del ratón
    def on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", on_mousewheel)


def main():
    root = tk.Tk()
    build_ui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
