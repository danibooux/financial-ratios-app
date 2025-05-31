import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

class CalculadoraFinanciera:
    def __init__(self, root):
        self.root = root
        self.root.title("Calculadora de Ratios Financieros")
        self.root.geometry("600x600")
        self.root.configure(bg="#f0f4f8")

        self.campos = {
            "Activo Corriente": None,
            "Pasivo Corriente": None,
            "Inventarios": None,
            "Efectivo y Equivalentes": None,
            "Utilidad Neta": None,
            "Ventas Netas": None,
            "Activos Totales": None,
            "Capital Contable": None,
            "Pasivo Total": None,
            "Costo de Ventas": None,
            "Cuentas por Cobrar": None,
            "Cuentas por Pagar": None
        }

        self.entries = {}
        self.resultados = {}

        self.crear_interfaz()

    def crear_interfaz(self):
        ttk.Label(self.root, text="Ingresa los datos financieros:", background="#f0f4f8", font=("Segoe UI", 12, "bold")).pack(pady=10)

        frame_campos = ttk.Frame(self.root)
        frame_campos.pack(pady=10)

        for i, campo in enumerate(self.campos):
            ttk.Label(frame_campos, text=campo + ":").grid(row=i, column=0, sticky="e", padx=5, pady=5)
            entry = ttk.Entry(frame_campos)
            entry.grid(row=i, column=1, padx=5, pady=5)
            self.entries[campo] = entry

        ttk.Button(self.root, text="Calcular", command=self.calcular_ratios).pack(pady=10)
        self.text_resultado = tk.Text(self.root, height=15, width=70)
        self.text_resultado.pack(pady=10)

        ttk.Button(self.root, text="Exportar a Excel", command=self.exportar_excel).pack(pady=10)

    def calcular_ratios(self):
        try:
            datos = {campo: float(entry.get()) for campo, entry in self.entries.items()}

            # Ratios de Liquidez
            razon_corriente = datos["Activo Corriente"] / datos["Pasivo Corriente"]
            prueba_acida = (datos["Activo Corriente"] - datos["Inventarios"]) / datos["Pasivo Corriente"]

            # Ratios de Rentabilidad
            margen_neto = datos["Utilidad Neta"] / datos["Ventas Netas"]
            roa = datos["Utilidad Neta"] / datos["Activos Totales"]
            roe = datos["Utilidad Neta"] / datos["Capital Contable"]

            # Ratios de Endeudamiento
            razon_deuda = datos["Pasivo Total"] / datos["Activos Totales"]
            capitalizacion = datos["Pasivo Total"] / datos["Capital Contable"]

            # Ratios de Actividad
            rotacion_cartera = datos["Ventas Netas"] / datos["Cuentas por Cobrar"]
            rotacion_inventario = datos["Costo de Ventas"] / datos["Inventarios"]
            rotacion_proveedores = datos["Costo de Ventas"] / datos["Cuentas por Pagar"]

            self.resultados = {
                "Razón Corriente": razon_corriente,
                "Prueba Ácida": prueba_acida,
                "Margen Neto": margen_neto,
                "ROA": roa,
                "ROE": roe,
                "Razón de Deuda": razon_deuda,
                "Capitalización": capitalizacion,
                "Rotación de Cartera": rotacion_cartera,
                "Rotación de Inventario": rotacion_inventario,
                "Rotación de Proveedores": rotacion_proveedores
            }

            self.text_resultado.delete("1.0", tk.END)
            for clave, valor in self.resultados.items():
                self.text_resultado.insert(tk.END, f"{clave}: {valor:.2f}\n")

        except ValueError:
            messagebox.showerror("Error", "Por favor ingresa todos los datos numéricos correctamente.")

    def exportar_excel(self):
        if self.resultados:
            df = pd.DataFrame(self.resultados.items(), columns=["Ratio Financiero", "Valor"])
            df.to_excel("ratios_financieros_completos.xlsx", index=False)
            messagebox.showinfo("Exportación Exitosa", "Archivo guardado como 'ratios_financieros_completos.xlsx'")
        else:
            messagebox.showerror("Error", "Primero debes calcular los ratios.")

# Ejecutar interfaz
root = tk.Tk()
app = CalculadoraFinanciera(root)
root.mainloop()
