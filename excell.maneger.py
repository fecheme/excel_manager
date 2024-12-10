import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk

class ExcelApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Gerenciador de Excel")
        self.master.geometry("800x600")

        self.label = tk.Label(master, text="Selecione uma Ação:")
        self.label.pack(pady=10)

        self.btn_open = tk.Button(master, text="Abrir Arquivo Excel", command=self.load_file)
        self.btn_open.pack(pady=10)

        self.filter_frame = tk.Frame(master)
        self.filter_frame.pack(pady=10)

        self.filter_label = tk.Label(self.filter_frame, text="Filtrar por Coluna:")
        self.filter_label.pack(side=tk.LEFT)

        self.filter_column = ttk.Combobox(self.filter_frame, state='readonly')
        self.filter_column.pack(side=tk.LEFT, padx=5)

        self.filter_value_label = tk.Label(self.filter_frame, text="Valor:")
        self.filter_value_label.pack(side=tk.LEFT)

        self.filter_value = ttk.Combobox(self.filter_frame, state='readonly')
        self.filter_value.pack(side=tk.LEFT, padx=5)

        self.btn_apply_filter = tk.Button(self.filter_frame, text="Aplicar Filtro", command=self.apply_filter, state=tk.DISABLED)
        self.btn_apply_filter.pack(side=tk.LEFT, padx=5)

        self.tree = None

    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.df = pd.read_excel(self.file_path)
            self.filter_column['values'] = list(self.df.columns)
            self.filter_column.bind("<<ComboboxSelected>>", self.update_filter_values)
            self.btn_apply_filter.config(state=tk.NORMAL)
            self.display_data(self.df)
            messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")

    def update_filter_values(self, event):
        coluna = self.filter_column.get()
        if coluna:
            unique_values = self.df[coluna].astype(str).unique().tolist()
            self.filter_value['values'] = unique_values
            self.filter_value.set('')

    def display_data(self, df):
        if self.tree:
            self.tree.destroy()

        self.tree = ttk.Treeview(self.master)
        self.tree.pack(pady=20, fill=tk.BOTH, expand=True)

        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def apply_filter(self):
        if hasattr(self, 'df'):
            coluna = self.filter_column.get()
            valor = self.filter_value.get()

            if coluna and valor:
                filtered_df = self.df[self.df[coluna].astype(str) == valor]

                if not filtered_df.empty:
                    self.display_data(filtered_df)
                    self.save_file(filtered_df, "Filtrado com sucesso!")
                else:
                    messagebox.showwarning("Aviso", "Nenhum dado encontrado com o filtro aplicado.")
            else:
                messagebox.showerror("Erro", "Por favor, selecione uma coluna e um valor.")

    def save_file(self, df, message):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Sucesso", message)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()

