import os
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


def generate_random_list_with_sum(n, target_sum):
    if n <= 0 or target_sum <= 0:
        return [0] * n
    dividers = sorted(random.sample(range(1, target_sum), n - 1))
    return [
        b - a for a, b in zip([0] + dividers, dividers + [target_sum])
    ]


class DiscountApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Excel Discount Generator")
        self.root.geometry("600x550")

        self.df = None
        self.file_path = ""
        self.item_entries = []

        # File Selection Frame
        frame_file = ttk.LabelFrame(root, text=" 1. Select Excel File ")
        frame_file.pack(fill="x", padx=10, pady=5)

        self.btn_file = ttk.Button(
            frame_file, text="Browse Excel", command=self.load_file
        )
        self.btn_file.pack(side="left", padx=5, pady=5)

        self.lbl_file = ttk.Label(
            frame_file, text="No file selected", foreground="gray"
        )
        self.lbl_file.pack(side="left", padx=5)

        # Global Settings Frame
        frame_settings = ttk.LabelFrame(root, text=" 2. Global Parameters ")
        frame_settings.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_settings, text="Max Boundary Discount:").grid(
            row=0, column=0, padx=5, pady=5, sticky="e"
        )
        self.ent_max_boundary = ttk.Entry(frame_settings)
        self.ent_max_boundary.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_settings, text="Number of Items to Discount:").grid(
            row=1, column=0, padx=5, pady=5, sticky="e"
        )
        self.ent_num_items = ttk.Entry(frame_settings)
        self.ent_num_items.grid(row=1, column=1, padx=5, pady=5)

        # Item Limits Table Frame
        frame_items = ttk.LabelFrame(root, text=" 3. Item Discount Limits ")
        frame_items.pack(fill="both", expand=True, padx=10, pady=5)

        canvas = tk.Canvas(frame_items)
        scrollbar = ttk.Scrollbar(
            frame_items, orient="vertical", command=canvas.yview
        )
        self.scroll_frame = ttk.Frame(canvas)

        self.scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Action Frame
        frame_action = ttk.Frame(root)
        frame_action.pack(fill="x", padx=10, pady=10)

        self.btn_run = ttk.Button(
            frame_action, text="Process & Save Excel", command=self.process
        )
        self.btn_run.pack(fill="x")

    def load_file(self):
        file = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file:
            return

        try:
            self.df = pd.read_excel(file)
            self.file_path = file
            self.lbl_file.config(
                text=os.path.basename(file), foreground="black"
            )
            self.populate_items()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file:\n{e}")

    def populate_items(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        self.item_entries.clear()

        ttk.Label(self.scroll_frame, text="Product", font=("sans", 9, "bold")).grid(
            row=0, column=0, padx=5, pady=2
        )
        ttk.Label(self.scroll_frame, text="Min Dis", font=("sans", 9, "bold")).grid(
            row=0, column=1, padx=5, pady=2
        )
        ttk.Label(self.scroll_frame, text="Max Dis", font=("sans", 9, "bold")).grid(
            row=0, column=2, padx=5, pady=2
        )

        for idx, row in self.df.iterrows():
            item_name = str(row.iloc[0])
            ttk.Label(self.scroll_frame, text=item_name).grid(
                row=idx + 1, column=0, padx=5, pady=2, sticky="w"
            )

            ent_min = ttk.Entry(self.scroll_frame, width=10)
            ent_min.insert(0, "0")
            ent_min.grid(row=idx + 1, column=1, padx=5, pady=2)

            ent_max = ttk.Entry(self.scroll_frame, width=10)
            ent_max.insert(0, "100")
            ent_max.grid(row=idx + 1, column=2, padx=5, pady=2)

            self.item_entries.append((ent_min, ent_max))

    def process(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return

        try:
            max_boundary = float(self.ent_max_boundary.get())
            num_discount = int(self.ent_num_items.get())
        except ValueError:
            messagebox.showerror(
                "Error", "Max boundary and Number of items must be valid numbers."
            )
            return

        if num_discount > len(self.df):
            messagebox.showerror(
                "Error",
                "Number of discount items exceeds total products in Excel.",
            )
            return

        min_dis_list, max_dis_list = [], []
        for ent_min, ent_max in self.item_entries:
            try:
                mn = float(ent_min.get())
                mx = float(ent_max.get())
                if mx > max_boundary:
                    messagebox.showerror(
                        "Error",
                        f"Item max discount ({mx}) cannot exceed max boundary ({max_boundary}).",
                    )
                    return
                min_dis_list.append(mn)
                max_dis_list.append(mx)
            except ValueError:
                messagebox.showerror(
                    "Error", "Min/Max discounts must be numeric."
                )
                return

        self.df["min_dis"] = min_dis_list
        self.df["max_dis"] = max_dis_list
        self.df["random_discount"] = 0

        selected_indices = random.sample(range(len(self.df)), num_discount)
        discounts = generate_random_list_with_sum(
            num_discount, int(max_boundary)
        )

        for idx, dis in zip(selected_indices, discounts):
            self.df.loc[idx, "random_discount"] = dis

        base, ext = os.path.splitext(self.file_path)
        output_path = f"{base}_discounted{ext}"

        try:
            self.df.to_excel(output_path, index=False)
            messagebox.showinfo(
                "Success", f"File saved successfully to:\n{output_path}"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save output file:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DiscountApp(root)
    root.mainloop() 