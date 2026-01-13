import tkinter as tk
from tkinter import filedialog, messagebox, Menu
import pandas as pd
import random
from collections import defaultdict

def about():
    tk.messagebox.showinfo(
        title="About",
        message="""This is a program to help with floor monitor scheduling.

Program was made for Rīgas 22. vidusskolas programming lessons, 2025./2026. school year.
Program restributed using GPL-3.0 license.""",
    )


def quitting():
    root.destroy()

def read_names(file_path):
    ext = file_path.split('.')[-1].lower()
    try:
        if ext == 'txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                names = [line.strip() for line in f if line.strip()]
        elif ext == 'csv':
            df = pd.read_csv(file_path)
            names = df.iloc[:, 0].dropna().astype(str).tolist()
        elif ext in ['xlsx', 'xls']:
            df = pd.read_excel(file_path)
            names = df.iloc[:, 0].dropna().astype(str).tolist()
        else:
            raise ValueError("Unsupported file type.")
        # Remove empty + duplicates while preserving relative order
        seen = set()
        unique_names = []
        for n in names:
            if n and n not in seen:
                seen.add(n)
                unique_names.append(n)
        return unique_names
    except Exception as e:
        raise ValueError(f"Error reading file: {e}")

def assign_people(floors, days, ppf, names, allow_reuse):
    if ppf > len(names):
        return None, "People per floor > number of available names."

    assignment = {}
    person_floor_count = defaultdict(lambda: defaultdict(int))

    if not allow_reuse:
        # No reuse, simple random shuffle once
        total_needed = floors * days * ppf
        if len(names) < total_needed:
            return None, f"Not enough unique names ({len(names)} < {total_needed})."
        shuffled = names.copy()
        random.shuffle(shuffled)
        idx = 0
        for d in range(days):
            for f in range(floors):
                assignment[(f, d)] = shuffled[idx:idx + ppf]
                idx += ppf
        return assignment, None

    # Allow reuse, fair + random
    for d in range(days):
        for f in range(floors):
            # Sort by (times on this floor, total assignments), fairest first
            # Then shuffle people with identical scores so it's not always alphabetical
            candidates = sorted(
                names,
                key=lambda n: (person_floor_count[n][f], sum(person_floor_count[n].values()))
            )
            # Group by score to shuffle within same-score groups
            from itertools import groupby
            keyed = lambda n: (person_floor_count[n][f], sum(person_floor_count[n].values()))
            shuffled_candidates = []
            for _, group in groupby(candidates, key=keyed):
                group_list = list(group)
                random.shuffle(group_list)
                shuffled_candidates.extend(group_list)

            selected = shuffled_candidates[:ppf]
            assignment[(f, d)] = selected

            for person in selected:
                person_floor_count[person][f] += 1

    return assignment, None

def create_dataframe(floors, days, assignment):
    data = {}
    for d in range(days):
        column = [', '.join(assignment.get((f, d), [])) for f in range(floors)]
        data[f'Day {d+1}'] = column
    df = pd.DataFrame(data, index=[f'Floor {f+1}' for f in range(floors)])
    return df

# Global so Save As can access the latest result
latest_df = None

def process():
    global latest_df
    try:
        floors = int(entry_floors.get())
        days   = int(entry_days.get())
        ppf    = int(entry_ppf.get())
        path   = file_var.get()

        if not path:
            raise ValueError("Select a names file first.")
        if floors < 1 or days < 1 or ppf < 1:
            raise ValueError("All numbers must be positive.")

        names = read_names(path)
        if not names:
            raise ValueError("No names found in the file.")

        total_needed = floors * days * ppf

        while len(names) < total_needed:
            reuse = messagebox.askyesno(
                "Not enough unique names",
                f"You need {total_needed} slots but only have {len(names)} unique names.\n\n"
                "Allow people to be scheduled multiple times?"
            )
            if reuse:
                break
            else:
                new_path = filedialog.askopenfilename(
                    title="Select file with more names",
                    filetypes=[("Text/CSV/Excel", "*.txt *.csv *.xlsx")]
                )
                if not new_path:
                    return
                file_var.set(new_path)
                names = read_names(new_path)

        allow_reuse = len(names) < total_needed
        assignment, err = assign_people(floors, days, ppf, names, allow_reuse)
        if err:
            raise ValueError(err)

        latest_df = create_dataframe(floors, days, assignment)

        text_output.delete(1.0, tk.END)
        text_output.insert(tk.END, "Random assignment generated!\n\n")
        text_output.insert(tk.END, latest_df.to_string())

    except ValueError as ve:
        messagebox.showerror("Error", str(ve))
    except Exception as e:
        messagebox.showerror("Unexpected error", str(e))

def save_as():
    global latest_df
    if latest_df is None:
        messagebox.showinfo("Nothing to save", "Generate an assignment first.")
        return
    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save assignment"
    )
    if not path:
        return
    try:
        if path.endswith('.xlsx'):
            latest_df.to_excel(path, index_label="Floor")
        else:
            latest_df.to_csv(path, encoding='utf-8')
        messagebox.showinfo("Saved", f"Assignment saved to:\n{path}")
    except Exception as e:
        messagebox.showerror("Save failed", str(e))

# GUI
root = tk.Tk()
root.title("Floor Monitor Scheduler")
root.geometry("900x680")

menubar = Menu(root)
root.config(menu=menubar)

file_menu = Menu(menubar, tearoff=0)
file_menu.add_command(label="Exit", command=quitting)
menubar.add_cascade(label="File", menu=file_menu)
menubar.add_command(label="About", command=about)

# Input row
tk.Label(root, text="Floors:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
entry_floors = tk.Entry(root, width=8); entry_floors.grid(row=0, column=1, sticky="w"); entry_floors.insert(0, "4")

tk.Label(root, text="Days:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
entry_days = tk.Entry(root, width=8); entry_days.grid(row=1, column=1, sticky="w"); entry_days.insert(0, "5")

tk.Label(root, text="People per floor:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
entry_ppf = tk.Entry(root, width=8); entry_ppf.grid(row=2, column=1, sticky="w"); entry_ppf.insert(0, "1")

# File selector
file_var = tk.StringVar()
tk.Label(root, text="Names file:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
tk.Entry(root, textvariable=file_var, width=60, state="readonly").grid(row=3, column=1, padx=5, sticky="w")
tk.Button(root, text="Browse…", command=lambda: file_var.set(filedialog.askopenfilename(
    filetypes=[("Text/CSV/Excel", "*.txt *.csv *.xlsx *.xls")]))).grid(row=3, column=2, sticky="w")

# Buttons
tk.Button(root, text="Generate Assignment", command=process, bg="#1976D2", fg="white", font=("Arial", 10, "bold"))\
    .grid(row=4, column=0, columnspan=2, pady=15, sticky="ew")
tk.Button(root, text="Save As CSV", command=save_as, bg="#388E3C", fg="white")\
    .grid(row=4, column=2, pady=15, sticky="w", padx=10)

# Output
text_output = tk.Text(root, font=("Consolas", 10), wrap="none")
text_output.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=10, pady=10)

scroll_y = tk.Scrollbar(root, command=text_output.yview)
scroll_y.grid(row=5, column=3, sticky="ns")
text_output.configure(yscrollcommand=scroll_y.set)

scroll_x = tk.Scrollbar(root, orient="horizontal", command=text_output.xview)
scroll_x.grid(row=6, column=0, columnspan=3, sticky="ew")
text_output.configure(xscrollcommand=scroll_x.set)

root.columnconfigure(1, weight=1)
root.rowconfigure(5, weight=1)

root.mainloop()
