import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from PIL import Image
import pytesseract
import re
import sys
from collections import Counter

# --- Core Data Extraction Logic ---

def find_and_parse_headers(lines):
    """
    Finds the header line, parses it, and returns a list of pandas Period objects.
    """
    header_line = ""
    header_line_regex = re.compile(r'((?:Q|H)\d|Total)\s?\.?\d{4}|\b\d{4}\b')
    for line in lines:
        if len(header_line_regex.findall(line)) >= 2:
            header_line = line
            break

    if not header_line:
        return []

    header_line = header_line.replace('#', 'H')
    parsed_headers = []
    period_regex = re.compile(r'(Q\d)\s?\.?(\d{4})|(H\d)\s?\.?(\d{4})|(\b\d{4}\b)')
    for match in period_regex.finditer(header_line):
        period_str, freq = "", ""
        if match.group(1) and match.group(2):
            period_str, freq = f"{match.group(2)}{match.group(1)}", 'Q'
        elif match.group(3) and match.group(4):
            quarter = 'Q2' if match.group(3) == 'H1' else 'Q4'
            period_str, freq = f"{match.group(4)}{quarter}", 'Q'
        elif match.group(5):
            period_str, freq = match.group(5), 'Y-DEC'
        if period_str:
            try:
                parsed_headers.append(pd.Period(period_str, freq=freq))
            except ValueError:
                pass
    return parsed_headers

def parse_financial_data(text):
    if not text: return pd.DataFrame()
    lines = text.strip().split('\n')
    all_rows = []
    number_regex = re.compile(r'\(?[\$€]?[\d,]+\.?\d*\)?')
    for line in lines:
        matches = list(number_regex.finditer(line))
        if len(matches) > 1:
            desc = line[:matches[0].start()].strip()
            if not desc or desc.lower() in ["basic", "diluted", "of which:"]: continue
            nums = [m.group(0) for m in matches]
            all_rows.append({'desc': desc, 'nums': nums, 'num_cols': len(nums)})
    if not all_rows: return pd.DataFrame()
    col_counts = Counter(row['num_cols'] for row in all_rows)
    if not col_counts: return pd.DataFrame()
    mode_cols = col_counts.most_common(1)[0][0]
    data = []
    for row in all_rows:
        if row['num_cols'] == mode_cols:
            cleaned_nums = [s.replace('$', '').replace('€', '').replace(',', '').replace('(', '-').replace(')', '') for s in row['nums']]
            data.append([row['desc']] + cleaned_nums)
    if not data: return pd.DataFrame()
    headers = find_and_parse_headers(lines)
    if len(headers) == mode_cols:
        columns = ['Description'] + headers
    else:
        if headers: print(f"Warning: Found {len(headers)} headers but data has {mode_cols} columns. Using generic headers.")
        columns = ['Description'] + [f'Value {i+1}' for i in range(mode_cols)]
    return pd.DataFrame(data, columns=columns)

def extract_text_from_image(filepath):
    try:
        return pytesseract.image_to_string(Image.open(filepath))
    except Exception as e:
        return f"Error extracting text from image: {e}"

def extract_text_from_pdf(filepath):
    try:
        doc = fitz.open(filepath)
        text = "".join(page.get_text() for page in doc)
        return text
    except Exception as e:
        return f"Error extracting text from PDF: {e}"

# --- GUI and App Logic ---
sorted_df_global, unsorted_df_global = pd.DataFrame(), pd.DataFrame()

def combine_and_sort(list_of_dfs):
    sortable = [df for df in list_of_dfs if any(isinstance(c, pd.Period) for c in df.columns)]
    unsortable = [df for df in list_of_dfs if not any(isinstance(c, pd.Period) for c in df.columns)]
    sorted_df = pd.DataFrame()
    if sortable:
        long_dfs = [df.melt(id_vars=['Description'], var_name='Period', value_name='Value') for df in sortable]
        combined = pd.concat(long_dfs, ignore_index=True).dropna(subset=['Value']).drop_duplicates(subset=['Description', 'Period'], keep='first')
        sorted_df = combined.pivot(index='Description', columns='Period', values='Value')
    unsorted_df = pd.concat(unsortable, ignore_index=True) if unsortable else pd.DataFrame()
    return sorted_df, unsorted_df

def load_and_process_files(text_widget, save_button):
    global sorted_df_global, unsorted_df_global
    filepaths = filedialog.askopenfilenames(title="Select Image/PDF File(s)", filetypes=(("All parsable files", "*.jpg *.jpeg *.png *.pdf"), ("All files", "*.*")))
    if not filepaths: return
    text_widget.delete('1.0', tk.END)
    save_button.config(state='disabled')
    all_dfs = []
    for fp in filepaths:
        text_widget.insert(tk.END, f"--- Processing: {fp} ---\n")
        raw_text = extract_text_from_pdf(fp) if fp.lower().endswith('.pdf') else extract_text_from_image(fp)
        if "Error" in raw_text:
            text_widget.insert(tk.END, f"Error: {raw_text}\n\n")
        else:
            df = parse_financial_data(raw_text)
            if not df.empty:
                all_dfs.append(df)
                text_widget.insert(tk.END, "Successfully parsed.\n\n")
            else:
                text_widget.insert(tk.END, "Could not parse data.\n\n")
    if all_dfs:
        sorted_df_global, unsorted_df_global = combine_and_sort(all_dfs)
        if not sorted_df_global.empty:
            text_widget.insert(tk.END, "\n--- Combined & Sorted Data ---\n" + sorted_df_global.to_string())
        if not unsorted_df_global.empty:
            text_widget.insert(tk.END, "\n\n--- Unsortable Data ---\n" + unsorted_df_global.to_string())
        if not sorted_df_global.empty or not unsorted_df_global.empty:
            save_button.config(state='normal')

def save_to_excel():
    global sorted_df_global, unsorted_df_global
    if sorted_df_global.empty and unsorted_df_global.empty:
        messagebox.showinfo("No Data", "There is no data to save.")
        return
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Save Excel File")
    if not filepath: return
    try:
        with pd.ExcelWriter(filepath) as writer:
            if not sorted_df_global.empty: sorted_df_global.to_excel(writer, sheet_name='Chronologically Sorted Data')
            if not unsorted_df_global.empty: unsorted_df_global.to_excel(writer, sheet_name='Unsortable Data')
        messagebox.showinfo("Success", f"Data saved to {filepath}")
    except Exception as e:
        messagebox.showerror("Error", f"Could not save file: {e}")

def create_gui():
    window = tk.Tk()
    window.title("Financial Statement Extractor")
    window.geometry("800x600")
    control_frame = tk.Frame(window)
    control_frame.pack(padx=10, pady=10, fill='x')
    text_widget = tk.Text(window, wrap='none', height=40, width=120)
    save_button = tk.Button(control_frame, text="Save to Excel", command=save_to_excel, state='disabled')
    save_button.pack(side='left', padx=10)
    load_button = tk.Button(control_frame, text="Load File(s)", command=lambda: load_and_process_files(text_widget, save_button))
    load_button.pack(side='left')
    text_widget.pack(padx=10, pady=10, fill='both', expand=True)
    window.mainloop()

def main():
    if len(sys.argv) > 1:
        filepaths = sys.argv[1:]
        all_dfs = []
        for fp in filepaths:
            raw_text = extract_text_from_pdf(fp) if fp.lower().endswith('.pdf') else extract_text_from_image(fp)
            if "Error" not in raw_text:
                df = parse_financial_data(raw_text)
                if not df.empty: all_dfs.append(df)
        if all_dfs:
            sorted_df, unsorted_df = combine_and_sort(all_dfs)
            if not sorted_df.empty: print("\n--- Sorted Data ---\n", sorted_df)
            if not unsorted_df.empty: print("\n--- Unsortable Data ---\n", unsorted_df)
    else:
        create_gui()

if __name__ == "__main__":
    main()
