# -*- coding: utf-8 -*-
# Guardar este archivo como desktop_analyzer.py

import tkinter as tk
from tkinter import ttk  # Themed Tkinter widgets for better look
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import io
import threading # To prevent GUI freezing during long operations
import warnings

# --- Panda/Data Loading Function (reusable, slight modification for error reporting) ---
# NOTE: Using @st.cache_data equivalent is harder in Tkinter without extra libs.
# We'll just load data on demand.

def load_data(file_path):
    """Carga datos desde un path de archivo (csv, xlsx, txt)."""
    df = None
    error_message = None
    try:
        file_extension = file_path.split('.')[-1].lower()

        if file_extension == 'csv':
            try:
                df = pd.read_csv(file_path, sep=None, engine='python')
                if df.shape[1] == 1: # Try semicolon if comma fails
                     df_semi = pd.read_csv(file_path, sep=';')
                     # Heuristic: if semicolon gives more columns, prefer it
                     if df_semi.shape[1] > df.shape[1]:
                         df = df_semi
            except Exception as e:
                error_message = f"Error complejo al leer CSV: {e}"

        elif file_extension == 'xlsx':
            try:
                # Requires openpyxl installed
                df = pd.read_excel(file_path, engine='openpyxl')
            except Exception as e:
                 error_message = f"Error al leer archivo Excel (.xlsx): {e}"

        elif file_extension == 'txt':
            try: # Try TSV first
                df = pd.read_csv(file_path, sep='\t')
            except Exception:
                 try: # Try space-separated
                     df = pd.read_csv(file_path, sep=r'\s+', engine='python')
                 except Exception as e2:
                     error_message = f"Error al leer .txt (TSV o espacios): {e2}"
        else:
            error_message = f"Formato de archivo '{file_extension}' no soportado."

        if df is not None:
            df.columns = df.columns.str.strip()

    except Exception as e:
        error_message = f"Error inesperado durante la carga: {e}"

    return df, error_message

def get_df_info_string(df):
    """Obtiene la salida de df.info() como un string."""
    if df is None:
        return "No hay datos cargados."
    buffer = io.StringIO()
    try:
        df.info(buf=buffer, verbose=True)
        return buffer.getvalue()
    except Exception as e:
        return f"Error obteniendo información del DataFrame: {e}"

# --- Helper to display DataFrame in Treeview ---
def populate_treeview(tree, dataframe):
    """Limpia y rellena un ttk.Treeview con un Pandas DataFrame."""
    if dataframe is None:
        # Clear previous contents if any
        for i in tree.get_children():
            tree.delete(i)
        tree["columns"] = ()
        return

    # Clear previous contents
    for i in tree.get_children():
        tree.delete(i)

    # Define columns
    tree["columns"] = list(dataframe.columns)
    tree["show"] = "headings" # Hide the default first empty column

    # Define headings
    for col in tree["columns"]:
        tree.heading(col, text=col)
        tree.column(col, anchor=tk.W, width=100) # Adjust width as needed

    # Add data rows (limit rows for performance in preview)
    # For full view, consider pagination or virtual lists if data is huge
    rows_to_show = dataframe.head(100).values.tolist() # Show first 100 rows
    for i, row in enumerate(rows_to_show):
        tree.insert("", tk.END, values=row, iid=str(i)) # Use iid to avoid issues

# --- Main Application Class ---
class DataAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📊 Analizador de Datos julian venegas")
        # self.geometry("1000x700") # Initial size

        # Style
        self.style = ttk.Style(self)
        self.style.theme_use('clam') # Or 'alt', 'default', 'classic'

        self.dataframe = None
        self.file_path = None

        # --- Main Layout ---
        # Top Frame for file selection
        self.top_frame = ttk.Frame(self, padding="10")
        self.top_frame.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(self.top_frame, text="📂 Abrir Archivo", command=self.browse_file).pack(side=tk.LEFT)
        self.file_label = ttk.Label(self.top_frame, text="Ningún archivo seleccionado")
        self.file_label.pack(side=tk.LEFT, padx=10)

        # Paned Window for resizable sections: Options | Main Content
        self.paned_window = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned_window.pack(side=tk.TOP, expand=True, fill=tk.BOTH, pady=5, padx=5)

        # Left Frame for Options
        self.options_frame = ttk.Labelframe(self.paned_window, text="⚙️ Opciones de Análisis", padding="10")
        self.paned_window.add(self.options_frame, weight=1) # Add with weight

        # Main Content Area (using Notebook for tabs)
        self.notebook = ttk.Notebook(self.paned_window)
        self.paned_window.add(self.notebook, weight=4) # Add with more weight

        # Status Bar
        self.status_bar = ttk.Label(self, text="Listo.", relief=tk.SUNKEN, anchor=tk.W, padding="2")
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # --- Create Tabs ---
        self.tab_preview = ttk.Frame(self.notebook, padding="5")
        self.tab_info = ttk.Frame(self.notebook, padding="5")
        self.tab_stats = ttk.Frame(self.notebook, padding="5")
        self.tab_missing = ttk.Frame(self.notebook, padding="5")
        self.tab_viz = ttk.Frame(self.notebook, padding="5")

        self.notebook.add(self.tab_preview, text='👀 Vista Previa')
        self.notebook.add(self.tab_info, text='ℹ️ Info General')
        self.notebook.add(self.tab_stats, text='📈 Estadísticas')
        self.notebook.add(self.tab_missing, text='❓ Val. Faltantes')
        self.notebook.add(self.tab_viz, text='📊 Visualización')

        # --- Widgets for Options Frame ---
        self.run_button = ttk.Button(self.options_frame, text="🚀 Ejecutar Análisis", command=self.run_analysis, state=tk.DISABLED)
        self.run_button.pack(pady=10, fill=tk.X)

        ttk.Separator(self.options_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)

        ttk.Label(self.options_frame, text="Visualización Individual:").pack(anchor=tk.W, pady=(5,0))
        self.viz_col_var = tk.StringVar()
        self.viz_col_combo = ttk.Combobox(self.options_frame, textvariable=self.viz_col_var, state=tk.DISABLED, postcommand=self.update_viz_combos)
        self.viz_col_combo.pack(fill=tk.X)
        # Placeholder for plot options (radio/slider) - added dynamically later if needed

        ttk.Separator(self.options_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)

        ttk.Label(self.options_frame, text="Diagrama de Dispersión (X vs Y):").pack(anchor=tk.W, pady=(5,0))
        self.scatter_x_var = tk.StringVar()
        self.scatter_x_combo = ttk.Combobox(self.options_frame, textvariable=self.scatter_x_var, state=tk.DISABLED, postcommand=self.update_viz_combos)
        self.scatter_x_combo.pack(fill=tk.X, pady=2)

        self.scatter_y_var = tk.StringVar()
        self.scatter_y_combo = ttk.Combobox(self.options_frame, textvariable=self.scatter_y_var, state=tk.DISABLED, postcommand=self.update_viz_combos)
        self.scatter_y_combo.pack(fill=tk.X, pady=2)

        self.scatter_color_var = tk.StringVar()
        self.scatter_color_combo = ttk.Combobox(self.options_frame, textvariable=self.scatter_color_var, state=tk.DISABLED, postcommand=self.update_viz_combos)
        self.scatter_color_combo.pack(fill=tk.X, pady=2)
        self.scatter_color_var.set(" (Sin Color) ") # Default


        # --- Widgets for Tabs (Placeholders/Containers) ---
        # Preview Tab
        self.preview_tree_frame = ttk.Frame(self.tab_preview)
        self.preview_tree_frame.pack(expand=True, fill=tk.BOTH)
        self.preview_tree = self._create_treeview_with_scrollbar(self.preview_tree_frame)

        # Info Tab
        self.info_text = scrolledtext.ScrolledText(self.tab_info, wrap=tk.WORD, state=tk.DISABLED, height=10)
        self.info_text.pack(expand=True, fill=tk.BOTH, pady=5)

        # Stats Tab
        self.stats_num_frame = ttk.Labelframe(self.tab_stats, text="Numéricas", padding=5)
        self.stats_num_frame.pack(expand=True, fill=tk.BOTH, pady=2)
        self.stats_num_tree = self._create_treeview_with_scrollbar(self.stats_num_frame)

        self.stats_cat_frame = ttk.Labelframe(self.tab_stats, text="Categóricas/Objeto", padding=5)
        self.stats_cat_frame.pack(expand=True, fill=tk.BOTH, pady=2)
        self.stats_cat_tree = self._create_treeview_with_scrollbar(self.stats_cat_frame)

        # Missing Tab
        self.missing_frame = ttk.Frame(self.tab_missing)
        self.missing_frame.pack(expand=True, fill=tk.BOTH)
        self.missing_tree = self._create_treeview_with_scrollbar(self.missing_frame)
        self.missing_label = ttk.Label(self.missing_frame, text="")
        self.missing_label.pack(pady=5)

        # Visualization Tab
        # This frame will hold the Matplotlib canvas and toolbar
        self.viz_plot_frame = ttk.Frame(self.tab_viz)
        self.viz_plot_frame.pack(expand=True, fill=tk.BOTH, side=tk.TOP)
        self.canvas = None # Placeholder for FigureCanvasTkAgg
        self.toolbar = None # Placeholder for NavigationToolbar2Tk


    def _create_treeview_with_scrollbar(self, parent):
        """Creates a Treeview with vertical and horizontal scrollbars."""
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(expand=True, fill=tk.BOTH)

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        tree = ttk.Treeview(tree_frame,
                            yscrollcommand=tree_scroll_y.set,
                            xscrollcommand=tree_scroll_x.set,
                            height=10) # Adjust height as needed

        tree_scroll_y.config(command=tree.yview)
        tree_scroll_x.config(command=tree.xview)

        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(expand=True, fill=tk.BOTH)
        return tree

    def set_status(self, text):
        self.status_bar.config(text=text)
        self.update_idletasks() # Force GUI update

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo de datos",
            filetypes=[("Todos los archivos compatibles", "*.csv *.xlsx *.txt"),
                       ("CSV (Separado por comas)", "*.csv"),
                       ("Excel", "*.xlsx"),
                       ("Texto (Tab o Espacios)", "*.txt"),
                       ("Todos los archivos", "*.*")]
        )
        if not file_path:
            return

        self.file_path = file_path
        self.file_label.config(text=file_path.split('/')[-1]) # Show filename
        self.set_status(f"Cargando archivo: {self.file_label['text']}...")

        # Load data in a separate thread to avoid freezing GUI
        thread = threading.Thread(target=self._load_data_thread, daemon=True)
        thread.start()

    def _load_data_thread(self):
        df, error = load_data(self.file_path)
        if error:
            self.dataframe = None
            self.after(0, lambda: messagebox.showerror("Error de Carga", error)) # Show error in main thread
            self.set_status("Error al cargar archivo.")
            self.after(0, self._update_ui_after_load, False) # Update UI in main thread
        elif df is None or df.empty:
             self.dataframe = None
             self.after(0, lambda: messagebox.showwarning("Archivo Vacío", "El archivo se cargó pero está vacío o no contiene datos válidos."))
             self.set_status("Archivo cargado pero vacío.")
             self.after(0, self._update_ui_after_load, False)
        else:
            self.dataframe = df
            self.set_status(f"Archivo '{self.file_label['text']}' cargado ({df.shape[0]} filas, {df.shape[1]} columnas).")
            # Schedule UI updates to run in the main Tkinter thread
            self.after(0, self._update_ui_after_load, True)

    def _update_ui_after_load(self, success):
        """Updates UI elements after data loading attempt (runs in main thread)."""
        if success and self.dataframe is not None:
            self.run_button.config(state=tk.NORMAL)
            self.viz_col_combo.config(state='readonly')
            self.scatter_x_combo.config(state='readonly')
            self.scatter_y_combo.config(state='readonly')
            self.scatter_color_combo.config(state='readonly')
            self.update_viz_combos() # Populate comboboxes now
            # Maybe trigger initial analysis automatically? Or wait for button.
            # self.run_analysis() # Optional: run analysis immediately
            self.notebook.select(self.tab_preview) # Switch to preview tab
            self.update_preview_tab() # Show preview immediately
        else:
            self.run_button.config(state=tk.DISABLED)
            self.viz_col_combo.config(state=tk.DISABLED)
            self.scatter_x_combo.config(state=tk.DISABLED)
            self.scatter_y_combo.config(state=tk.DISABLED)
            self.scatter_color_combo.config(state=tk.DISABLED)
            # Clear previous results
            self.clear_all_tabs()

    def update_viz_combos(self):
        """Updates the content of visualization comboboxes based on df columns."""
        if self.dataframe is None:
            self.viz_col_combo['values'] = []
            self.scatter_x_combo['values'] = []
            self.scatter_y_combo['values'] = []
            self.scatter_color_combo['values'] = [" (Sin Color) "]
            return

        all_cols = list(self.dataframe.columns)
        numeric_cols = self.dataframe.select_dtypes(include='number').columns.tolist()
        categorical_cols = self.dataframe.select_dtypes(include=['object', 'category', 'boolean']).columns.tolist()

        self.viz_col_combo['values'] = all_cols
        self.scatter_x_combo['values'] = numeric_cols
        self.scatter_y_combo['values'] = numeric_cols
        self.scatter_color_combo['values'] = [" (Sin Color) "] + categorical_cols

        # Set defaults if empty or invalid
        if not self.viz_col_var.get() or self.viz_col_var.get() not in all_cols:
            self.viz_col_var.set(all_cols[0] if all_cols else "")
        if not self.scatter_x_var.get() or self.scatter_x_var.get() not in numeric_cols:
            self.scatter_x_var.set(numeric_cols[0] if numeric_cols else "")
        if not self.scatter_y_var.get() or self.scatter_y_var.get() not in numeric_cols:
            self.scatter_y_var.set(numeric_cols[1] if len(numeric_cols) > 1 else (numeric_cols[0] if numeric_cols else ""))
        # Keep color selection or default
        if not self.scatter_color_var.get() or (self.scatter_color_var.get() != " (Sin Color) " and self.scatter_color_var.get() not in categorical_cols):
             self.scatter_color_var.set(" (Sin Color) ")


    def run_analysis(self):
        """Runs all selected analyses and updates the tabs."""
        if self.dataframe is None:
            messagebox.showwarning("Sin Datos", "Primero carga un archivo de datos.")
            return

        self.set_status("Ejecutando análisis...")
        # Run updates in background thread? For now, run sequentially
        try:
            self.update_preview_tab() # Already done on load, but maybe refresh
            self.update_info_tab()
            self.update_stats_tab()
            self.update_missing_tab()
            self.update_viz_tab() # Update based on selections
            self.set_status("Análisis completado.")
            # Optionally switch to a specific tab after analysis
            # self.notebook.select(self.tab_viz)
        except Exception as e:
            messagebox.showerror("Error de Análisis", f"Ocurrió un error durante el análisis:\n{e}")
            self.set_status("Error durante el análisis.")

    def clear_all_tabs(self):
         populate_treeview(self.preview_tree, None)
         self.info_text.config(state=tk.NORMAL)
         self.info_text.delete('1.0', tk.END)
         self.info_text.config(state=tk.DISABLED)
         populate_treeview(self.stats_num_tree, None)
         populate_treeview(self.stats_cat_tree, None)
         populate_treeview(self.missing_tree, None)
         self.missing_label.config(text="")
         self.clear_plot()


    def update_preview_tab(self):
        populate_treeview(self.preview_tree, self.dataframe) # Shows head(100)

    def update_info_tab(self):
        info_str = get_df_info_string(self.dataframe)
        self.info_text.config(state=tk.NORMAL) # Enable writing
        self.info_text.delete('1.0', tk.END) # Clear previous text
        self.info_text.insert('1.0', info_str)
        self.info_text.config(state=tk.DISABLED) # Disable editing

    def update_stats_tab(self):
        try:
            numeric_stats = self.dataframe.describe(include='number')
            populate_treeview(self.stats_num_tree, numeric_stats.reset_index())
        except Exception:
             populate_treeview(self.stats_num_tree, pd.DataFrame({'Error':['No se pudieron calcular stats numéricas']}))

        try:
            # Include boolean as well
            cat_df = self.dataframe.select_dtypes(include=['object', 'category', 'boolean'])
            if not cat_df.empty:
                cat_stats = cat_df.describe()
                populate_treeview(self.stats_cat_tree, cat_stats.reset_index())
            else:
                 populate_treeview(self.stats_cat_tree, pd.DataFrame({'Info':['No hay columnas categóricas/objeto/boolean']}))
        except Exception:
            populate_treeview(self.stats_cat_tree, pd.DataFrame({'Error':['No se pudieron calcular stats categóricas']}))

    def update_missing_tab(self):
        missing_data = self.dataframe.isnull().sum()
        missing_data = missing_data[missing_data > 0]
        if not missing_data.empty:
            missing_df = pd.DataFrame({
                'Columna': missing_data.index,
                'Valores Faltantes': missing_data.values,
                '% Faltante': (missing_data.values / len(self.dataframe) * 100).round(2)
            }).sort_values(by='% Faltante', ascending=False)
            populate_treeview(self.missing_tree, missing_df)
            self.missing_label.config(text=f"Se encontraron valores faltantes en {len(missing_data)} columnas.")
        else:
            populate_treeview(self.missing_tree, None) # Clear tree if no missing
            self.missing_label.config(text="¡No se encontraron valores faltantes!")

    def clear_plot(self):
         """Removes the current plot canvas and toolbar."""
         if self.canvas:
             self.canvas.get_tk_widget().destroy()
             self.canvas = None
         if self.toolbar:
             self.toolbar.destroy()
             self.toolbar = None
         # Clear any plot-specific options widgets if they exist
         for widget in self.viz_plot_frame.winfo_children():
             # Keep the main frame itself, destroy others
             if widget not in [self.viz_plot_frame]:
                  widget.destroy()


    def update_viz_tab(self):
        """Generates plots based on combobox selections."""
        self.clear_plot() # Clear previous plot first

        plot_choice = self.viz_col_var.get()
        scatter_x = self.scatter_x_var.get()
        scatter_y = self.scatter_y_var.get()
        scatter_color = self.scatter_color_var.get() if self.scatter_color_var.get() != " (Sin Color) " else None

        # Determine which plot to make based on selections
        # Simple logic: Prioritize scatter if both X and Y are valid and different
        make_scatter = False
        if scatter_x and scatter_y and scatter_x != scatter_y:
             numeric_cols = self.dataframe.select_dtypes(include='number').columns.tolist()
             if scatter_x in numeric_cols and scatter_y in numeric_cols:
                 make_scatter = True


        # --- Create Plot ---
        fig, ax = plt.subplots(figsize=(7, 5)) # Create a new figure and axes
        plot_made = False

        try:
            warnings.filterwarnings("ignore", category=UserWarning, module="seaborn") # Suppress warnings

            if make_scatter:
                sns.scatterplot(data=self.dataframe, x=scatter_x, y=scatter_y, hue=scatter_color, alpha=0.7, ax=ax)
                ax.set_title(f'{scatter_x} vs {scatter_y}' + (f' (Color: {scatter_color})' if scatter_color else ''))
                # Simple legend handling
                if scatter_color and self.dataframe[scatter_color].nunique() < 20:
                    ax.legend(title=scatter_color, bbox_to_anchor=(1.05, 1), loc='upper left')
                elif ax.get_legend() is not None:
                    ax.get_legend().remove()
                plot_made = True
                self.set_status(f"Mostrando diagrama de dispersión: {scatter_x} vs {scatter_y}")

            elif plot_choice:
                # If not making scatter, try individual column plot
                if plot_choice in self.dataframe.select_dtypes(include='number').columns:
                    # Histogram for numeric
                    sns.histplot(self.dataframe[plot_choice], kde=True, ax=ax, bins=30)
                    ax.set_title(f'Histograma de {plot_choice}')
                    plot_made = True
                    self.set_status(f"Mostrando histograma: {plot_choice}")
                elif plot_choice in self.dataframe.select_dtypes(include=['object', 'category', 'boolean']).columns:
                    # Bar plot for categorical (Top N)
                    counts = self.dataframe[plot_choice].value_counts()
                    top_n = min(20, len(counts)) # Limit to top 20
                    counts_top = counts.nlargest(top_n)
                    sns.barplot(x=counts_top.values, y=counts_top.index.astype(str), ax=ax, palette="viridis", orient='h')
                    ax.set_title(f'Top {top_n} Categorías en {plot_choice}')
                    ax.set_xlabel('Frecuencia')
                    plot_made = True
                    self.set_status(f"Mostrando barras de frecuencia: {plot_choice}")

            else:
                 # If no valid selection, show placeholder text
                 ax.text(0.5, 0.5, 'Selecciona columnas en el panel de opciones\npara generar una visualización.',
                         horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)
                 self.set_status("Selecciona opciones para visualizar.")

            # --- Embed Plot in Tkinter ---
            # Ensure figure layout adjusts for titles/labels/legends
            try:
                 fig.tight_layout()
            except Exception: # Can fail sometimes with complex plots/legends
                 pass

            self.canvas = FigureCanvasTkAgg(fig, master=self.viz_plot_frame)
            self.canvas.draw()
            # Place canvas widget
            canvas_widget = self.canvas.get_tk_widget()
            canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            # Add Matplotlib toolbar if a plot was made
            if plot_made:
                self.toolbar = NavigationToolbar2Tk(self.canvas, self.viz_plot_frame)
                self.toolbar.update()
                # Place toolbar widget below plot
                self.toolbar.pack(side=tk.BOTTOM, fill=tk.X)
            else:
                # If no plot was made, explicitly remove potential old toolbar
                if self.toolbar:
                     self.toolbar.destroy()
                     self.toolbar = None


        except Exception as e:
            plt.close(fig) # Close the figure if error occurred
            self.clear_plot() # Ensure canvas/toolbar are removed
            messagebox.showerror("Error de Visualización", f"No se pudo generar el gráfico:\n{e}")
            self.set_status("Error al generar gráfico.")
        finally:
            warnings.filterwarnings("default", category=UserWarning, module="seaborn") # Restore warnings


# --- Main Execution ---
if __name__ == "__main__":
    app = DataAnalyzerApp()
    app.mainloop()