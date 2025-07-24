import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re

# Set up Matplotlib font support
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False  # Solve negative sign display issue


class SPTLiquefactionSoftware:
    def __init__(self, root):
        self.root = root
        self.root.title(
            "A Software for Seismic Sand Liquefaction Index and Classification Based on Standard Penetration Test")
        self.root.geometry("1200x700")
        self.root.resizable(True, True)

        # Configure ttk styles for better button appearance
        self.style = ttk.Style()
        self.style.configure("TButton",
                             padding=6,  # Increase padding for better appearance
                             font=("Arial", 10))  # Slightly larger font
        self.style.map("TButton",
                       background=[("active", "#e1e1e1"), ("pressed", "#c1c1c1")])

        # Multi-point data storage: each point contains basic parameters and soil layer data
        self.points = [self._create_default_point()]  # Initial point
        self.current_point_idx = 0  # Current point index

        # Column name mapping library: standard column names -> possible aliases (supports fuzzy matching)
        self.column_mapping = {
            "Point ID": ["Point ID", "Point No", "Measurement Point ID", "ID", "Point Number"],
            "Saturated soil depth ds(m)": ["Saturated soil depth ds(m)", "Saturated depth(m)", "ds", "Saturated depth",
                                           "Soil depth ds", "Depth ds"],
            "Measured N-value": ["Measured N-value", "N-value", "Measured penetration value", "Standard penetration N",
                                 "N", "Blow count N"],
            "Layer thickness di(m)": ["Layer thickness di(m)", "Thickness di(m)", "Soil layer thickness", "di",
                                      "Thickness(m)", "Layer thickness di"]
        }

        # Create menu bar
        self.create_menu()

        # Create notebook interface
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Manual input tab
        self.tab_manual = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_manual, text="Manual Input")

        # Table import tab
        self.tab_import = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_import, text="Table Import")

        # Results display tab
        self.tab_results = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_results, text="Calculation Results")

        # Initialize tabs
        self.init_manual_input()
        self.init_import_page()
        self.init_results_page()

        # Store calculation results (by point)
        self.results = {}
        # Store current identified column mappings
        self.current_column_mapping = {}

    def _create_default_point(self):
        """Create default point data structure"""
        return {
            "intensity": "7",
            "depth_criteria": "15",
            "dw": 2.0,
            "layers": [
                {"ds": 1.5, "N": 12, "di": 3.0},
                {"ds": 4.5, "N": 14, "di": 3.0},
                {"ds": 7.5, "N": 16, "di": 3.0},
                {"ds": 10.5, "N": 18, "di": 3.0},
                {"ds": 13.5, "N": 20, "di": 3.0}
            ]
        }

    def create_menu(self):
        """Create menu bar and help menu"""
        menubar = tk.Menu(self.root)

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Software Introduction", command=self.show_help)
        help_menu.add_command(label="Parameter Explanation", command=self.show_parameter_help)
        help_menu.add_command(label="Liquefaction Classification Standards", command=self.show_liquefaction_standards)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)

        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)

    def show_help(self):
        """Display software introduction"""
        help_text = """
        A Software for Seismic Sand Liquefaction Index and Classification Based on Standard Penetration Test

        This software calculates the soil liquefaction index (ILE) and evaluates liquefaction classification 
        strictly according to the methods specified in "SPT-Liquefaction Classification.docx".
        The calculation process fully follows the 4 steps in the document:
        1. Calculate the standard penetration critical value Ncr (considering saturated soil depth and groundwater level)
        2. Calculate the safety factor FS
        3. Calculate the liquefaction index ILE
        4. Determine the liquefaction classification (according to Table 1)

        Function description:
        1. Multi-point input: supports management and calculation of multiple site points data
        2. Table import: automatically identifies common column name aliases without strict matching
        3. Result display: includes liquefaction index, classification determination and corresponding 
           anti-liquefaction measures (according to Table 2)

        Usage process:
        1. Fill in parameters on the "Manual Input" or "Table Import" tab
        2. Click the "Calculate" button
        3. View results and anti-liquefaction measures on the "Calculation Results" tab
        """
        self.show_info_dialog("Software Introduction", help_text)

    def show_parameter_help(self):
        """Display parameter explanation"""
        param_text = """
        Parameter Explanation :

        1. General parameters:
        - Seismic Intensity: 7/8/9 degrees 
        - Discrimination Depth: 15m or 20m (used for liquefaction classification, corresponding to Table 1 standards)
        - Groundwater level depth dw(m): depth of groundwater level at the site

        2. Soil layer parameters:
        - Saturated soil standard penetration depth ds(m): depth of standard penetration test point in soil layer
        - Measured standard penetration value N: blow count of standard penetration test
        - Layer thickness di(m): vertical thickness of soil layer

        3. Supported column name formats for table import:
        - Point ID: Point No, Measurement Point ID, ID, Point Number, etc.
        - Saturated soil depth ds(m): Saturated depth(m), ds, Saturated depth, Soil depth ds, etc.
        - Measured N-value: N-value, Measured penetration value, Standard penetration N, N, etc.
        - Layer thickness di(m): Thickness di(m), Soil layer thickness, di, Thickness(m), etc.
        """
        self.show_info_dialog("Parameter Explanation", param_text)

    def show_liquefaction_standards(self):
        """Display liquefaction classification standards (content from Table 1 and Table 2)"""
        standard_text = """
        Core content of "SPT-Liquefaction Classification.docx":

        I. Table 1: Discrimination of liquefaction classification
        +----------------+--------------------------------+--------------------------------+
        | Liquefaction   | Liquefaction index when        | Liquefaction index when        |
        | Classification | discrimination depth is 15m    | discrimination depth is 20m    |
        +----------------+--------------------------------+--------------------------------+
        | Slight         | 0＜ILE≤5                       | 0＜ILE≤6                       |
        +----------------+--------------------------------+--------------------------------+
        | Moderate       | 5＜ILE≤15                      | 6＜ILE≤18                      |
        +----------------+--------------------------------+--------------------------------+
        | Severe         | ILE＞15                        | ILE＞18                        |
        +----------------+--------------------------------+--------------------------------+


        II. Table 2: Anti-liquefaction measures
        +-------------------------------+----------------------------------------------+--------------------------------------------------+--------------------------------------------------------+
        | Seismic Fortification Category| Slight Liquefaction                          | Moderate Liquefaction                            | Severe Liquefaction                                    |
        +-------------------------------+----------------------------------------------+--------------------------------------------------+--------------------------------------------------------+
        | Category B                    | Partially eliminate liquefaction settlement, | Completely eliminate liquefaction settlement, or | Completely eliminate liquefaction settlement           |
        |                               | or treat foundation and superstructure       | partially eliminate and treat foundation and     |                                                        |
        |                               |                                              | superstructure                                   |                                                        |
        +-------------------------------+----------------------------------------------+--------------------------------------------------+--------------------------------------------------------+
        | Category C                    | Treat foundation and superstructure,         | Treat foundation and superstructure, or take     | Completely eliminate liquefaction settlement, or       |
        |                               | or no measures may be taken                  | more stringent measures                          | partially eliminate and treat foundation and           |
        |                               |                                              |                                                  | superstructure                                         |
        +-------------------------------+----------------------------------------------+--------------------------------------------------+--------------------------------------------------------+
        | Category D                    | No measures may be taken                     | No measures may be taken                         | Treat foundation and superstructure, or take other     |
        |                               |                                              |                                                  | economical measures                                    |
        +-------------------------------+----------------------------------------------+--------------------------------------------------+--------------------------------------------------------+
        """
        self.show_info_dialog("Liquefaction Classification Standards", standard_text)

    def show_about(self):
        """Display about information"""
        messagebox.showinfo("About",
                            "A Software for Seismic Sand Liquefaction Index and Classification Based on Standard Penetration Test\nDeveloped according to \"SSSLIaCBSPT.docx\"")

    def show_info_dialog(self, title, content):
        """Display information dialog box"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("800x600")
        dialog.resizable(True, True)

        text_widget = scrolledtext.ScrolledText(dialog, wrap=tk.WORD, font=("Courier New", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)

        # Improved button with style
        ttk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=10)

    def init_manual_input(self):
        """Initialize manual input tab (supports multi-point input)"""
        # Multi-point management frame
        point_frame = ttk.LabelFrame(self.tab_manual, text="Site Point Management")
        point_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(point_frame, text="Current Point:").grid(row=0, column=0, padx=10, pady=5)
        self.point_combobox = ttk.Combobox(point_frame, state="readonly", width=10)
        self.point_combobox.grid(row=0, column=1, padx=5, pady=5)
        self.point_combobox.bind("<<ComboboxSelected>>", self._on_point_changed)

        # Improved buttons with consistent padding
        ttk.Button(point_frame, text="Add Point", command=self._add_point).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(point_frame, text="Delete Point", command=self._remove_point).grid(row=0, column=3, padx=5, pady=5)

        # Basic parameters frame
        basic_frame = ttk.LabelFrame(self.tab_manual,
                                     text="Basic Parameters (according to \"SPT-Liquefaction Classification.docx\")")
        basic_frame.pack(fill=tk.X, padx=10, pady=10)

        # Seismic intensity selection
        ttk.Label(basic_frame, text="Seismic Intensity:").grid(row=0, column=0, padx=10, pady=5)
        self.intensity_var = tk.StringVar(value="7")
        ttk.Combobox(basic_frame, textvariable=self.intensity_var, values=["7", "8", "9"], width=10).grid(row=0,
                                                                                                          column=1,
                                                                                                          padx=10,
                                                                                                          pady=5)

        # Discrimination depth selection
        ttk.Label(basic_frame, text="Discrimination Depth(m):").grid(row=0, column=2, padx=10, pady=5)
        self.depth_var = tk.StringVar(value="15")
        ttk.Combobox(basic_frame, textvariable=self.depth_var, values=["15", "20"], width=10).grid(row=0, column=3,
                                                                                                   padx=10, pady=5)

        # Groundwater level depth
        ttk.Label(basic_frame, text="Groundwater Level Depth dw(m):").grid(row=1, column=0, padx=10, pady=5)
        self.dw_var = tk.DoubleVar(value=2.0)
        ttk.Entry(basic_frame, textvariable=self.dw_var, width=12).grid(row=1, column=1, padx=10, pady=5)

        # Soil layer parameters table
        self.layer_frame = ttk.LabelFrame(self.tab_manual,
                                          text="Soil Layer Parameters (according to \"SPT-Liquefaction Classification.docx\")")
        self.layer_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Table headers
        headers = ["Saturated soil depth ds(m)", "Measured N-value", "Layer thickness di(m)"]
        for i, header in enumerate(headers):
            ttk.Label(self.layer_frame, text=header, font=("Arial", 10, "bold")).grid(row=0, column=i, padx=20, pady=5)

        # Soil layer input rows (store Entry widgets)
        self.layer_entries = []
        self._refresh_layer_entries()

        # Button frame with better spacing
        btn_frame = ttk.Frame(self.tab_manual)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        # Buttons with consistent styling and padding
        ttk.Button(btn_frame, text="Add Soil Layer", command=self.add_layer).pack(side=tk.LEFT, padx=8, pady=5)
        ttk.Button(btn_frame, text="Delete Last Layer", command=self.remove_layer).pack(side=tk.LEFT, padx=8, pady=5)
        ttk.Button(btn_frame, text="Clear Input", command=self.clear_manual_input).pack(side=tk.RIGHT, padx=8, pady=5)
        ttk.Button(btn_frame, text="Calculate", command=self.calculate_manual).pack(side=tk.RIGHT, padx=8, pady=5)

        # Initialize point dropdown menu
        self._refresh_point_combobox()

    def _refresh_point_combobox(self):
        """Refresh point selection dropdown menu"""
        point_ids = [f"Point {i + 1}" for i in range(len(self.points))]
        self.point_combobox['values'] = point_ids
        if point_ids:
            self.point_combobox.current(self.current_point_idx)

    def _add_point(self):
        """Add new site point"""
        new_point = self._create_default_point()
        self.points.append(new_point)
        self.current_point_idx = len(self.points) - 1
        self._refresh_point_combobox()
        self._refresh_layer_entries()
        # Update basic parameters
        self.intensity_var.set(new_point["intensity"])
        self.depth_var.set(new_point["depth_criteria"])
        self.dw_var.set(new_point["dw"])

    def _remove_point(self):
        """Delete current site point"""
        if len(self.points) <= 1:
            messagebox.showwarning("Warning", "At least one site point must be retained")
            return
        del self.points[self.current_point_idx]
        self.current_point_idx = max(0, self.current_point_idx - 1)
        self._refresh_point_combobox()
        self._on_point_changed(None)

    def _on_point_changed(self, event):
        """Update interface when switching site points"""
        self.current_point_idx = self.point_combobox.current()
        current_point = self.points[self.current_point_idx]
        # Update basic parameters
        self.intensity_var.set(current_point["intensity"])
        self.depth_var.set(current_point["depth_criteria"])
        self.dw_var.set(current_point["dw"])
        # Update soil layer inputs
        self._refresh_layer_entries()

    def _refresh_layer_entries(self):
        """Refresh soil layer input boxes"""
        # Clear existing input boxes
        for row_entries in self.layer_entries:
            for entry in row_entries:
                entry.destroy()
        self.layer_entries = []

        # Create new input boxes
        current_point = self.points[self.current_point_idx]
        for i, layer in enumerate(current_point["layers"]):
            row_entries = []
            for j in range(3):  # 3 parameters
                entry = ttk.Entry(self.layer_frame, width=15)
                entry.grid(row=i + 1, column=j, padx=20, pady=3)
                if j == 0:
                    entry.insert(0, f"{layer['ds']}")
                elif j == 1:
                    entry.insert(0, f"{layer['N']}")
                elif j == 2:
                    entry.insert(0, f"{layer['di']}")
                row_entries.append(entry)
            self.layer_entries.append(row_entries)

    def add_layer(self):
        """Add soil layer input row"""
        current_point = self.points[self.current_point_idx]
        # Calculate default ds (last layer ds + 3)
        last_ds = current_point["layers"][-1]["ds"] if current_point["layers"] else 1.5
        new_ds = last_ds + 3.0
        new_N = 12 + len(current_point["layers"]) * 2
        new_layer = {"ds": new_ds, "N": new_N, "di": 3.0}
        current_point["layers"].append(new_layer)
        self._refresh_layer_entries()

    def remove_layer(self):
        """Delete last soil layer"""
        current_point = self.points[self.current_point_idx]
        if len(current_point["layers"]) <= 1:
            messagebox.showwarning("Warning", "At least one soil layer must be retained")
            return
        current_point["layers"].pop()
        self._refresh_layer_entries()

    def init_import_page(self):
        """Initialize table import tab (optimized layout)"""
        # Main container uses grid layout
        self.tab_import.grid_columnconfigure(0, weight=3)  # Left side occupies 3/5 width
        self.tab_import.grid_columnconfigure(1, weight=2)  # Right side occupies 2/5 width
        self.tab_import.grid_rowconfigure(0, weight=1)

        # Left area: file operations and data preview
        left_frame = ttk.Frame(self.tab_import)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=10)
        left_frame.grid_rowconfigure(3, weight=1)  # Preview area auto-expands

        # Left title
        ttk.Label(left_frame, text="Table Data Import", font=("Arial", 12, "bold")).pack(pady=10)
        ttk.Label(left_frame, text="Supports Excel format, automatically identifies common column name formats",
                  font=("Arial", 10)).pack(pady=5)

        # File selection area
        file_frame = ttk.LabelFrame(left_frame, text="File Selection")
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        # Buttons with better spacing and consistent styling
        ttk.Button(file_frame, text="Browse Files", command=self.browse_file).pack(side=tk.LEFT, padx=8, pady=5)
        ttk.Button(file_frame, text="Clear Import", command=self.clear_import).pack(side=tk.LEFT, padx=8, pady=5)
        self.file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path, width=30).pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True,
                                                                          pady=5)

        # Column identification result display
        identify_frame = ttk.LabelFrame(left_frame, text="Column Identification Results (Auto-matching)")
        identify_frame.pack(fill=tk.X, padx=5, pady=5)

        self.identify_labels = {}
        for i, std_col in enumerate(self.column_mapping.keys()):
            ttk.Label(identify_frame, text=f"{std_col}:").grid(row=i, column=0, padx=10, pady=3, sticky=tk.W)
            label = ttk.Label(identify_frame, text="Not identified")
            label.grid(row=i, column=1, padx=5, pady=3, sticky=tk.W)
            self.identify_labels[std_col] = label
            identify_frame.grid_columnconfigure(1, weight=1)  # Let identification result column auto-expand

        # Table preview area
        preview_frame = ttk.LabelFrame(left_frame, text="Data Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.preview_canvas = tk.Canvas(preview_frame)
        preview_scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        preview_scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_canvas.xview)
        self.preview_frame_inner = ttk.Frame(self.preview_canvas)

        self.preview_frame_inner.bind(
            "<Configure>",
            lambda e: self.preview_canvas.configure(
                scrollregion=self.preview_canvas.bbox("all")
            )
        )

        self.preview_canvas.create_window((0, 0), window=self.preview_frame_inner, anchor=tk.NW)
        self.preview_canvas.configure(yscrollcommand=preview_scroll_y.set, xscrollcommand=preview_scroll_x.set)

        # Layout preview area scrollbars
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        preview_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Right area: parameter settings and data statistics
        right_frame = ttk.Frame(self.tab_import)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=10)
        right_frame.grid_rowconfigure(1, weight=1)  # Statistics area auto-expands

        # Right title
        ttk.Label(right_frame, text="Calculation Parameter Settings", font=("Arial", 12, "bold")).pack(pady=10)

        # Import parameter settings (upper part of right side)
        param_frame = ttk.LabelFrame(right_frame, text="Calculation Parameters")
        param_frame.pack(fill=tk.X, padx=5, pady=5)

        # Parameter grid layout
        param_frame.grid_columnconfigure(1, weight=1)
        row_idx = 0

        ttk.Label(param_frame, text="Seismic Intensity:").grid(row=row_idx, column=0, padx=5, pady=8, sticky=tk.W)
        self.import_intensity_var = tk.StringVar(value="7")
        ttk.Combobox(param_frame, textvariable=self.import_intensity_var, values=["7", "8", "9"], width=10).grid(
            row=row_idx, column=1, padx=5, pady=8, sticky=tk.W)
        row_idx += 1

        ttk.Label(param_frame, text="Discrimination Depth(m):").grid(row=row_idx, column=0, padx=5, pady=8, sticky=tk.W)
        self.import_depth_var = tk.StringVar(value="15")
        ttk.Combobox(param_frame, textvariable=self.import_depth_var, values=["15", "20"], width=10).grid(
            row=row_idx, column=1, padx=5, pady=8, sticky=tk.W)
        row_idx += 1

        ttk.Label(param_frame, text="Groundwater Level Depth dw(m):").grid(row=row_idx, column=0, padx=5, pady=8,
                                                                           sticky=tk.W)
        self.import_dw_var = tk.DoubleVar(value=2.0)
        ttk.Entry(param_frame, textvariable=self.import_dw_var, width=12).grid(
            row=row_idx, column=1, padx=5, pady=8, sticky=tk.W)
        row_idx += 1

        # Point selection (multi-point import)
        ttk.Label(param_frame, text="Select Calculation Point:").grid(row=row_idx, column=0, padx=5, pady=8,
                                                                      sticky=tk.W)
        self.import_point_combobox = ttk.Combobox(param_frame, state="readonly", width=10)
        self.import_point_combobox.grid(row=row_idx, column=1, padx=5, pady=8, sticky=tk.W)
        row_idx += 1

        # Data statistics information (middle part of right side)
        self.stats_frame = ttk.LabelFrame(right_frame, text="Data Statistics")
        self.stats_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Initial statistics information
        self.stats_text = scrolledtext.ScrolledText(self.stats_frame, wrap=tk.WORD, height=10)
        self.stats_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.stats_text.insert(tk.END, "Please import data to display statistics...")
        self.stats_text.config(state=tk.DISABLED)

        # Import tips (bottom part of right side)
        tips_frame = ttk.LabelFrame(right_frame, text="Import Tips")
        tips_frame.pack(fill=tk.X, padx=5, pady=5)

        tips_text = """
1. Table should contain: Point ID, Saturated soil depth, Measured N-value, Layer thickness
2. Supports multiple column name formats (e.g., "Point No", "N-value", etc.)
3. Data should be grouped by Point ID; same Point ID represents the same measurement point
4. Please confirm column identification results are correct after import
        """
        ttk.Label(tips_frame, text=tips_text, justify=tk.LEFT).pack(fill=tk.X, padx=5, pady=5)

        # Calculate buttons (fixed at the bottom of import tab) with better styling
        btn_frame = ttk.Frame(self.tab_import)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")

        # Add "Calculate All Points" button
        ttk.Button(btn_frame, text="Calculate All Points", command=self.calculate_all_imported_points).pack(
            side=tk.LEFT, padx=8, pady=8, ipady=2)

        # Larger calculate button for better visibility
        calc_btn = ttk.Button(btn_frame, text="Calculate Selected Point", command=self.calculate_import)
        calc_btn.pack(side=tk.LEFT, padx=8, pady=8, ipady=2)  # Increase internal padding

        # Store imported multi-point data
        self.imported_points = {}

    def calculate_all_imported_points(self):
        """Calculate all imported points at once"""
        # Check if column identification is complete
        if not self.current_column_mapping or any(not val for val in self.current_column_mapping.values()):
            messagebox.showwarning("Warning",
                                   "Please first import data containing complete columns and ensure identification is successful")
            return

        # Check if there are imported points
        if not self.imported_points:
            messagebox.showwarning("Warning", "Please first import data")
            return

        try:
            intensity = int(self.import_intensity_var.get())
            depth_criteria = int(self.import_depth_var.get())
            dw = self.import_dw_var.get()

            # Show progress message
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Calculating")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            progress_window.grab_set()

            ttk.Label(progress_window, text="Calculating all points, please wait...").pack(pady=20)
            self.root.update_idletasks()

            # Calculate each point
            total_points = len(self.imported_points)
            for i, (point_id, point_data) in enumerate(self.imported_points.items()):
                # Update progress
                progress_window.title(f"Calculating ({i + 1}/{total_points})")

                # Construct soil layer data
                layers = []
                for _, row in point_data.iterrows():
                    layers.append({
                        "ds": row[self.current_column_mapping["Saturated soil depth ds(m)"]],
                        "N": row[self.current_column_mapping["Measured N-value"]],
                        "di": row[self.current_column_mapping["Layer thickness di(m)"]]
                    })

                # Calculate and store results
                result = self.calculate_liquefaction_index({
                    "intensity": intensity,
                    "depth_criteria": depth_criteria,
                    "dw": dw,
                    "layers": layers
                })
                self.results[f"Imported Point {point_id}"] = result

                # Update UI
                self.root.update_idletasks()

            # Close progress window
            progress_window.destroy()

            # Update results display
            self._refresh_result_point_combobox()
            if self.results:
                self.result_point_combobox.current(0)
                self._on_result_point_changed(None)

            # Switch to results tab
            self.notebook.select(self.tab_results)

            messagebox.showinfo("Complete", f"Successfully calculated {total_points} points")

        except Exception as e:
            messagebox.showerror("Error", f"Calculation failed: {str(e)}")

    def _normalize_text(self, text):
        """Normalize text for matching: remove special characters, spaces and convert case"""
        if not text:
            return ""
        # Remove special characters and spaces
        normalized = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', text)
        # Convert to lowercase
        return normalized.lower()

    def _auto_identify_columns(self, columns):
        """Automatically identify table column names"""
        normalized_columns = {self._normalize_text(col): col for col in columns}
        mapping = {}

        for std_col, aliases in self.column_mapping.items():
            best_match = None
            # Normalize standard column name
            std_normalized = self._normalize_text(std_col)

            # 1. Prefer exact match
            if std_normalized in normalized_columns:
                best_match = normalized_columns[std_normalized]
            else:
                # 2. Match aliases
                for alias in aliases:
                    alias_normalized = self._normalize_text(alias)
                    if alias_normalized in normalized_columns:
                        best_match = normalized_columns[alias_normalized]
                        break
                # 3. Fuzzy match (contains keywords)
                if not best_match:
                    for norm_col, orig_col in normalized_columns.items():
                        if std_normalized in norm_col or any(self._normalize_text(a) in norm_col for a in aliases):
                            best_match = orig_col
                            break
            mapping[std_col] = best_match

        return mapping

    def _update_data_stats(self, df):
        """Update data statistics information"""
        self.stats_text.config(state=tk.NORMAL)
        self.stats_text.delete(1.0, tk.END)

        # Basic statistics
        self.stats_text.insert(tk.END, f"Total records: {len(df)}\n")

        # Point count statistics
        point_col = self.current_column_mapping.get("Point ID")
        if point_col and point_col in df.columns:
            point_count = df[point_col].nunique()
            self.stats_text.insert(tk.END, f"Number of site points: {point_count}\n\n")

            # Records per point
            self.stats_text.insert(tk.END, "Records per point:\n")
            point_counts = df[point_col].value_counts().to_dict()
            for point, count in sorted(point_counts.items()):
                self.stats_text.insert(tk.END, f"  Point {point}: {count} records\n")

        # N-value statistics
        n_col = self.current_column_mapping.get("Measured N-value")
        if n_col and n_col in df.columns:
            self.stats_text.insert(tk.END, "\nMeasured N-value statistics:\n")
            self.stats_text.insert(tk.END, f"  Minimum: {df[n_col].min()}\n")
            self.stats_text.insert(tk.END, f"  Maximum: {df[n_col].max()}\n")
            self.stats_text.insert(tk.END, f"  Average: {df[n_col].mean():.2f}\n")

        # Depth statistics
        ds_col = self.current_column_mapping.get("Saturated soil depth ds(m)")
        if ds_col and ds_col in df.columns:
            self.stats_text.insert(tk.END, "\nSaturated soil depth statistics:\n")
            self.stats_text.insert(tk.END, f"  Minimum: {df[ds_col].min():.2f}m\n")
            self.stats_text.insert(tk.END, f"  Maximum: {df[ds_col].max():.2f}m\n")

        self.stats_text.config(state=tk.DISABLED)

    def browse_file(self):
        """Browse and preview Excel files (supports multi-point)"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.file_path.set(file_path)
            self.preview_file(file_path)

    def preview_file(self, file_path):
        """Preview imported file data (automatically identify column names)"""
        # Clear previous preview
        for widget in self.preview_frame_inner.winfo_children():
            widget.destroy()
        self.imported_points = {}
        self.import_point_combobox['values'] = []
        self.current_column_mapping = {}

        try:
            df = pd.read_excel(file_path)
            columns = df.columns.tolist()

            # Automatically identify column names
            self.current_column_mapping = self._auto_identify_columns(columns)

            # Update identification result display
            for std_col, label in self.identify_labels.items():
                label.config(
                    text=self.current_column_mapping[std_col] if self.current_column_mapping[
                        std_col] else "Not identified")

            # Update data statistics
            self._update_data_stats(df)

            # Check for missing required columns
            missing = [col for col, val in self.current_column_mapping.items() if not val]
            if missing:
                messagebox.showwarning("Identification Tip",
                                       f"The following required columns were not identified, please check the table:\n{', '.join(missing)}")
                return

            # Group by point ID
            point_col = self.current_column_mapping["Point ID"]
            self.imported_points = {
                str(point_id): group.reset_index(drop=True)
                for point_id, group in df.groupby(point_col)
            }

            # Update point selection dropdown
            self.import_point_combobox['values'] = list(self.imported_points.keys())
            if self.imported_points:
                self.import_point_combobox.current(0)

            # Display all point data preview
            # Display headers
            for j, col in enumerate(columns):
                ttk.Label(self.preview_frame_inner, text=col, font=("Arial", 10, "bold")).grid(row=0, column=j, padx=15,
                                                                                               pady=5)
            # Display data (show up to 20 rows)
            display_rows = min(20, len(df))
            for i in range(display_rows):
                row = df.iloc[i]
                for j, value in enumerate(row):
                    ttk.Label(self.preview_frame_inner, text=str(value)).grid(row=i + 1, column=j, padx=15, pady=2)

        except Exception as e:
            messagebox.showerror("Error", f"File reading failed: {str(e)}")

    def init_results_page(self):
        """Initialize results tab (with clear results option)"""
        results_frame = ttk.LabelFrame(self.tab_results,
                                       text="Calculation Results (according to \"SPT-Liquefaction Classification.docx\")")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Top button frame (contains clear results button)
        top_btn_frame = ttk.Frame(results_frame)
        top_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        # Clear button with better styling
        ttk.Button(top_btn_frame, text="Clear Calculation Results", command=self.clear_results).pack(side=tk.RIGHT,
                                                                                                     padx=8, pady=5)

        # Point selection frame (for switching multi-point results)
        point_select_frame = ttk.Frame(results_frame)
        point_select_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(point_select_frame, text="View Point Results:").pack(side=tk.LEFT, padx=5)
        # 增加宽度从10到18，确保内容完全显示
        self.result_point_combobox = ttk.Combobox(point_select_frame, state="readonly", width=18)
        self.result_point_combobox.pack(side=tk.LEFT, padx=5)
        self.result_point_combobox.bind("<<ComboboxSelected>>", self._on_result_point_changed)

        # Left result summary
        summary_frame = ttk.LabelFrame(results_frame, text="Liquefaction Index and Classification")
        summary_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.ile_var = tk.DoubleVar(value=0.0)
        self.grade_var = tk.StringVar(value="Not calculated")
        self.intensity_var_display = tk.StringVar(value="--")
        self.depth_var_display = tk.StringVar(value="--")
        self.dw_var_display = tk.StringVar(value="--")

        ttk.Label(summary_frame, text="Seismic Intensity:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Label(summary_frame, textvariable=self.intensity_var_display, font=("Arial", 12)).grid(row=0, column=1,
                                                                                                   pady=5)

        ttk.Label(summary_frame, text="Discrimination Depth:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Label(summary_frame, textvariable=self.depth_var_display, font=("Arial", 12)).grid(row=1, column=1, pady=5)

        ttk.Label(summary_frame, text="Groundwater Level Depth:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Label(summary_frame, textvariable=self.dw_var_display, font=("Arial", 12)).grid(row=2, column=1, pady=5)

        ttk.Label(summary_frame, text="Liquefaction Index (ILE):").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Label(summary_frame, textvariable=self.ile_var, font=("Arial", 12, "bold")).grid(row=3, column=1, pady=5)

        ttk.Label(summary_frame, text="Liquefaction Classification:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Label(summary_frame, textvariable=self.grade_var, font=("Arial", 12, "bold")).grid(row=4, column=1, pady=5)

        # Building category selection (for anti-liquefaction measures)
        ttk.Label(summary_frame, text="Seismic Fortification Category:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.building_class_combobox = ttk.Combobox(summary_frame, values=["Category B", "Category C", "Category D"],
                                                    width=10)
        self.building_class_combobox.current(0)
        self.building_class_combobox.grid(row=5, column=1, pady=5)
        self.building_class_combobox.bind("<<ComboboxSelected>>", self._update_anti_measures)

        # Right classification explanation and anti-liquefaction measures
        right_frame = ttk.Frame(results_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Liquefaction classification criteria (Table 2)
        grade_frame = ttk.LabelFrame(right_frame, text="Liquefaction Classification Criteria (Table 2)")
        grade_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        grade_text = """
        For discrimination depth 15m:
        0 < ILE ≤ 5 → Slight
        5 < ILE ≤ 15 → Moderate
        ILE > 15 → Severe

        For discrimination depth 20m:
        0 < ILE ≤ 6 → Slight
        6 < ILE ≤ 18 → Moderate
        ILE > 18 → Severe
        """
        ttk.Label(grade_frame, text=grade_text, justify=tk.LEFT).pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Anti-liquefaction measures (Table 3)
        self.anti_measures_frame = ttk.LabelFrame(right_frame, text="Anti-Liquefaction Measures (Table 3)")
        self.anti_measures_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.anti_measures_var = tk.StringVar(value="Please calculate liquefaction classification first")
        ttk.Label(self.anti_measures_frame, textvariable=self.anti_measures_var, justify=tk.LEFT).pack(
            fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Calculation details
        detail_frame = ttk.LabelFrame(results_frame, text="Calculation Details for Each Soil Layer")
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5, side=tk.BOTTOM)

        self.detail_text = scrolledtext.ScrolledText(detail_frame, height=10)
        self.detail_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Chart area
        chart_frame = ttk.LabelFrame(results_frame, text="Safety Factor Distribution")
        chart_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5, side=tk.BOTTOM)

        self.fig, self.ax = plt.subplots(figsize=(8, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def clear_results(self):
        """Clear all calculation results"""
        self.results = {}
        self.result_point_combobox['values'] = []
        # Reset display
        self.ile_var.set(0.0)
        self.grade_var.set("Not calculated")
        self.intensity_var_display.set("--")
        self.depth_var_display.set("--")
        self.dw_var_display.set("--")
        self.detail_text.delete(1.0, tk.END)
        self.ax.clear()
        self.canvas.draw()
        self.anti_measures_var.set("Please calculate liquefaction classification first")

    def _on_result_point_changed(self, event):
        """Update display when switching result points"""
        if not self.results:
            return
        point_id = self.result_point_combobox.get()
        if point_id in self.results:
            self._update_result_display(self.results[point_id])

    def _update_anti_measures(self, event):
        """Update anti-liquefaction measures display (according to Table 3)"""
        current_grade = self.grade_var.get()
        if current_grade in ["Not calculated", "No liquefaction"]:
            self.anti_measures_var.set("Please calculate a valid liquefaction classification first")
            return

        building_class = self.building_class_combobox.get()
        # Table 3 data mapping
        anti_measures_map = {
            "Category B": {
                "Slight": "Partially eliminate liquefaction settlement, or treat foundation and superstructure",
                "Moderate": "Completely eliminate liquefaction settlement, or partially eliminate and treat foundation and superstructure",
                "Severe": "Completely eliminate liquefaction settlement"
            },
            "Category C": {
                "Slight": "Treat foundation and superstructure, or no measures may be taken",
                "Moderate": "Treat foundation and superstructure, or take more stringent measures",
                "Severe": "Completely eliminate liquefaction settlement, or partially eliminate and treat foundation and superstructure"
            },
            "Category D": {
                "Slight": "No measures may be taken",
                "Moderate": "No measures may be taken",
                "Severe": "Treat foundation and superstructure, or take other economical measures"
            }
        }
        self.anti_measures_var.set(anti_measures_map[building_class].get(current_grade, "No corresponding measures"))

    def get_manual_data(self):
        """Get manually input data for current point"""
        try:
            current_point_id = f"Point {self.current_point_idx + 1}"
            intensity = int(self.intensity_var.get())
            depth_criteria = int(self.depth_var.get())
            dw = self.dw_var.get()

            # Save current point data
            self.points[self.current_point_idx] = {
                "intensity": self.intensity_var.get(),
                "depth_criteria": self.depth_var.get(),
                "dw": dw,
                "layers": [
                    {
                        "ds": float(entries[0].get()),
                        "N": float(entries[1].get()),
                        "di": float(entries[2].get())
                    } for entries in self.layer_entries
                ]
            }

            return current_point_id, {
                "intensity": intensity,
                "depth_criteria": depth_criteria,
                "dw": dw,
                "layers": self.points[self.current_point_idx]["layers"]
            }
        except Exception as e:
            messagebox.showerror("Input Error", f"Please check your input: {str(e)}")
            return None, None

    def calculate_manual(self):
        """Calculate for manual input (store results by point)"""
        point_id, data = self.get_manual_data()
        if data:
            result = self.calculate_liquefaction_index(data)
            self.results[point_id] = result
            self._refresh_result_point_combobox()
            self._on_result_point_changed(None)
            self.notebook.select(self.tab_results)

    def _refresh_result_point_combobox(self):
        """Refresh result point selection dropdown"""
        point_ids = list(self.results.keys())
        self.result_point_combobox['values'] = point_ids
        if point_ids:
            self.result_point_combobox.current(0)

    def _update_result_display(self, result):
        """Update result display"""
        self.intensity_var_display.set(f"{result['intensity']} degrees")
        self.depth_var_display.set(f"{result['depth_criteria']}m")
        self.dw_var_display.set(f"{result['dw']}m")
        self.ile_var.set(round(result["ile"], 2))
        self.grade_var.set(result["grade"])

        # Update detailed results
        self.detail_text.delete(1.0, tk.END)
        header = "{:^12}{:^10}{:^10}{:^10}{:^10}{:^12}\n".format(
            "Saturated Depth", "Measured N", "N0", "Ncr", "FS", "Contribution"
        )
        self.detail_text.insert(tk.END, header)
        self.detail_text.insert(tk.END, "-" * 65 + "\n")

        for layer in result["layers"]:
            row = "{:^12.1f}{:^10.1f}{:^10.0f}{:^10.2f}{:^10.2f}{:^12.2f}\n".format(
                layer["ds"], layer["N"], layer["N0"],
                layer["Ncr"], layer["FS"], layer["contribution"]
            )
            self.detail_text.insert(tk.END, row)

        self.detail_text.insert(tk.END, "\n" + "-" * 65 + "\n")
        self.detail_text.insert(tk.END, f"Total Liquefaction Index (ILE): {result['ile']:.2f}\n")
        self.detail_text.insert(tk.END, f"Liquefaction Classification: {result['grade']}")

        # Draw safety factor distribution chart
        self.ax.clear()
        depths = [l["ds"] for l in result["layers"]]
        fs_values = [l["FS"] for l in result["layers"]]
        self.ax.plot(fs_values, depths, 'o-')
        self.ax.axvline(x=1.0, color='r', linestyle='--', label='Liquefaction threshold (FS=1)')
        self.ax.set_xlabel('Safety Factor FS')
        self.ax.set_ylabel('Saturated Soil Depth (m)')
        self.ax.set_title('Distribution of Safety Factors for Each Soil Layer')
        self.ax.invert_yaxis()  # Depth increases downward
        self.ax.legend()
        self.fig.tight_layout()
        self.canvas.draw()

        # Update anti-liquefaction measures
        self._update_anti_measures(None)

    def calculate_import(self):
        """Calculate for imported data (supports multi-point)"""
        # Check if column identification is complete
        if not self.current_column_mapping or any(not val for val in self.current_column_mapping.values()):
            messagebox.showwarning("Warning",
                                   "Please first import data containing complete columns and ensure identification is successful")
            return

        # Check if calculation point is selected
        if not self.imported_points or self.import_point_combobox.current() == -1:
            messagebox.showwarning("Warning", "Please first import data and select a calculation point")
            return

        try:
            selected_point_id = self.import_point_combobox.get()
            point_data = self.imported_points[selected_point_id]
            intensity = int(self.import_intensity_var.get())
            depth_criteria = int(self.import_depth_var.get())
            dw = self.import_dw_var.get()

            # Construct soil layer data (using automatically identified column names)
            layers = []
            for _, row in point_data.iterrows():
                layers.append({
                    "ds": row[self.current_column_mapping["Saturated soil depth ds(m)"]],
                    "N": row[self.current_column_mapping["Measured N-value"]],
                    "di": row[self.current_column_mapping["Layer thickness di(m)"]]
                })

            # Calculate and store results
            result = self.calculate_liquefaction_index({
                "intensity": intensity,
                "depth_criteria": depth_criteria,
                "dw": dw,
                "layers": layers
            })
            self.results[f"Imported Point {selected_point_id}"] = result
            self._refresh_result_point_combobox()
            self._on_result_point_changed(None)
            self.notebook.select(self.tab_results)

        except Exception as e:
            messagebox.showerror("Error", f"Calculation failed: {str(e)}")

    def calculate_liquefaction_index(self, data):
        """Core calculation (according to \"SPT-Liquefaction Classification.docx\")"""
        # Standard penetration blow count reference values
        N0_map = {7: 13, 8: 15, 9: 19}
        N0 = N0_map[data["intensity"]]
        dw = data["dw"]

        results = []
        for layer in data["layers"]:
            ds = layer["ds"]

            # Calculate standard penetration critical value Ncr
            if ds <= dw:
                depth_factor = 1.0 + 0.05 * (dw - ds)
            else:
                depth_factor = 1.0 + 0.1 * (ds - dw)
            Ncr = N0 * depth_factor

            # Calculate safety factor FS
            FS = layer["N"] / Ncr if Ncr != 0 else float('inf')

            # Calculate weight wi
            if ds <= 5:
                wi = 10.0
            elif ds <= 15:
                wi = 5.0
            elif ds <= 20:
                wi = 2.0
            else:
                wi = 0.0

            # Calculate contribution value
            contribution = max(0, 1 - FS) * layer["di"] * wi if FS <= 1 else 0

            results.append({
                "ds": ds,
                "N": layer["N"],
                "N0": N0,
                "Ncr": Ncr,
                "FS": FS,
                "di": layer["di"],
                "wi": wi,
                "contribution": contribution
            })

        # Total liquefaction index
        ile = sum(r["contribution"] for r in results)

        # Determine liquefaction classification (Table 2)
        depth = data["depth_criteria"]
        if ile <= 0:
            grade = "No liquefaction"
        elif depth == 15:
            if ile <= 5:
                grade = "Slight"
            elif ile <= 15:
                grade = "Moderate"
            else:
                grade = "Severe"
        else:  # 20m
            if ile <= 6:
                grade = "Slight"
            elif ile <= 18:
                grade = "Moderate"
            else:
                grade = "Severe"

        return {
            "intensity": data["intensity"],
            "depth_criteria": depth,
            "dw": dw,
            "ile": ile,
            "grade": grade,
            "layers": results
        }

    def clear_manual_input(self):
        """Clear manual input for current point"""
        self.intensity_var.set("7")
        self.depth_var.set("15")
        self.dw_var.set(2.0)
        # Reset to default 5 layers
        current_point = self.points[self.current_point_idx]
        current_point["layers"] = self._create_default_point()["layers"]
        self._refresh_layer_entries()

    def clear_import(self):
        """Clear imported data"""
        self.file_path.set("")
        for widget in self.preview_frame_inner.winfo_children():
            widget.destroy()
        self.imported_points = {}
        self.import_point_combobox['values'] = []
        self.import_intensity_var.set("7")
        self.import_depth_var.set("15")
        self.import_dw_var.set(2.0)
        # Reset identification results
        for label in self.identify_labels.values():
            label.config(text="Not identified")
        self.current_column_mapping = {}
        # Reset statistics information
        self.stats_text.config(state=tk.NORMAL)
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, "Please import data to display statistics...")
        self.stats_text.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = SPTLiquefactionSoftware(root)
    root.mainloop()
