import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Basic GUI setup
root = tk.Tk()
root.title("Portfolio Allocation")
root.configure(bg="#add8e6")

# Center the window
window_width = 800
window_height = 900
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = int(screen_height/2 - window_height/2)
position_right = int(screen_width/2 - window_width/2)
root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

# Create and place the file selection widgets
frame = tk.Frame(root, bg="#add8e6")
frame.pack(pady=10)

# Label and entry for file path
file_label = tk.Label(frame, text="Excel File:", bg="#add8e6", fg="black")
file_label.grid(row=0, column=0, padx=10, pady=10)
file_entry = tk.Entry(frame, width=50)
file_entry.grid(row=0, column=1, padx=10, pady=10)
file_button = tk.Button(frame, text="Browse", command=lambda: browse_file(file_entry))
file_button.grid(row=0, column=2, padx=10, pady=10)

# Label and entry for portfolio worth
portfolio_label = tk.Label(frame, text="Portfolio Worth:", bg="#add8e6", fg="black")
portfolio_label.grid(row=1, column=0, padx=10, pady=10)
portfolio_entry = tk.Entry(frame, width=20)
portfolio_entry.grid(row=1, column=1, padx=10, pady=10)

# Button to trigger the calculation
calculate_button = tk.Button(frame, text="Calculate Allocation", command=lambda: calculate_allocation(file_entry, portfolio_entry))
calculate_button.grid(row=2, column=0, columnspan=3, pady=20)

# Treeview for displaying results
columns = ("Stock", "Average Close", "Average Daily Return", "Volatility", "Historical VaR", "Monte Carlo VaR", "Allocation")
tree = ttk.Treeview(root, columns=columns, show="headings")
tree.pack(pady=10)

# Set up column headings
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=150)

# Frame for the pie chart
pie_frame = tk.Frame(root, bg="#add8e6")
pie_frame.pack(fill=tk.BOTH, expand=True, pady=10)

def browse_file(entry):
    """
    Open a file dialog to browse for an Excel file and insert the file path into the entry widget.
    """
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def load_data(file_path):
    """
    Load the Excel file and return a pandas ExcelFile object.
    """
    try:
        return pd.ExcelFile(file_path)
    except Exception as e:
        raise ValueError(f"Error loading Excel file: {str(e)}")

def check_required_columns(df):
    """
    Check if the required columns are present in the dataframe.
    """
    required_columns = {'Close'}
    return required_columns.issubset(df.columns)

def calculate_metrics(df):
    """
    Calculate financial metrics for the given dataframe.
    """
    df['Close'] = pd.to_numeric(df['Close'], errors='coerce')
    df.dropna(subset=['Close'], inplace=True)
    df['Return'] = df['Close'].pct_change().dropna()
    return {
        'average_close': df['Close'].mean(),
        'average_daily_return': df['Return'].mean(),
        'volatility': df['Return'].std(ddof=0)
    }

def historical_var(df, confidence_level=0.95):
    """
    Calculate historical Value at Risk (VaR).
    """
    sorted_returns = df['Return'].sort_values()
    index = int((1 - confidence_level) * len(sorted_returns))
    return sorted_returns.iloc[index]

def monte_carlo_var(avg_return, volatility, portfolio_worth, num_simulations=10000, confidence_level=0.95):
    """
    Calculate Monte Carlo Value at Risk (VaR).
    """
    simulated_end_values = np.zeros(num_simulations)
    for i in range(num_simulations):
        simulated_returns = np.random.normal(avg_return, volatility, 252)
        simulated_end_values[i] = portfolio_worth * np.prod(1 + simulated_returns)
    return np.percentile(simulated_end_values, (1 - confidence_level) * 100)

def allocate_portfolio(metrics_list, portfolio_worth):
    """
    Allocate the portfolio based on calculated metrics.
    """
    total_inverse_volatility = sum(1 / metrics['volatility'] for _, metrics in metrics_list)
    total_inverse_hvar = sum(1 / metrics['historical_var'] for _, metrics in metrics_list)
    total_inverse_mcvar = sum(1 / metrics['monte_carlo_var'] for _, metrics in metrics_list)

    allocations = []
    for sheet_name, metrics in metrics_list:
        inverse_volatility = 1 / metrics['volatility']
        inverse_hvar = 1 / metrics['historical_var']
        inverse_mcvar = 1 / metrics['monte_carlo_var']
        normalized_metric = (
            (inverse_volatility / total_inverse_volatility) * 0.33 + 
            (inverse_hvar / total_inverse_hvar) * 0.33 + 
            (inverse_mcvar / total_inverse_mcvar) * 0.33
        )
        allocation = normalized_metric * portfolio_worth
        allocations.append((sheet_name, metrics, allocation))

    total_allocated = sum(allocation for _, _, allocation in allocations)
    adjustment_factor = portfolio_worth / total_allocated

    return [(sheet_name, metrics, allocation * adjustment_factor) for sheet_name, metrics, allocation in allocations]

def run_allocation(file_path, portfolio_worth):
    """
    Run the portfolio allocation calculation.
    """
    excel_file = load_data(file_path)
    results = pd.DataFrame(columns=["Stock", "Average Close", "Average Daily Return", "Volatility", "Historical VaR", "Monte Carlo VaR", "Allocation"])
    metrics_list = []

    # Loop through each sheet in the Excel file
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        if not df.empty and check_required_columns(df):
            metrics = calculate_metrics(df)
            metrics['historical_var'] = historical_var(df)
            metrics['monte_carlo_var'] = monte_carlo_var(metrics['average_daily_return'], metrics['volatility'], portfolio_worth)
            metrics_list.append((sheet_name, metrics))

    if not metrics_list:
        raise ValueError("The Excel file contains no non-empty sheets with the required columns.")

    allocations = allocate_portfolio(metrics_list, portfolio_worth)

    for sheet_name, metrics, allocation in allocations:
        stock_allocation_df = pd.DataFrame({
            "Stock": [sheet_name], 
            "Average Close": [metrics['average_close']], 
            "Average Daily Return": [metrics['average_daily_return']], 
            "Volatility": [metrics['volatility']], 
            "Historical VaR": [metrics['historical_var']],
            "Monte Carlo VaR": [metrics['monte_carlo_var']],
            "Allocation": [allocation]
        })
        results = pd.concat([results, stock_allocation_df], ignore_index=True)

    return results

def calculate_allocation(file_entry, portfolio_entry):
    """
    Calculate the portfolio allocation and update the GUI with the results.
    """
    file_path = file_entry.get()
    try:
        portfolio_worth = float(portfolio_entry.get())
    except ValueError:
        messagebox.showerror("Input Error", "Portfolio worth must be a number.")
        return

    try:
        results = run_allocation(file_path, portfolio_worth)
        # Clear the treeview before inserting new results
        for row in tree.get_children():
            tree.delete(row)
        for _, row in results.iterrows():
            tree.insert("", "end", values=[f"{val:.4f}" if isinstance(val, float) else val for val in row])
        plot_pie_chart(results)
    except Exception as e:
        messagebox.showerror("Error", str(e))

def plot_pie_chart(results):
    """
    Plot the pie chart for portfolio allocation.
    """
    fig, ax = plt.subplots()
    stocks = results['Stock']
    allocations = results['Allocation']
    total_allocation = allocations.sum()
    percentages = (allocations / total_allocation) * 100
    colors = plt.cm.tab20.colors  # Use tab20 colors

    num_colors_needed = len(stocks)
    if num_colors_needed > len(colors):
        colors = colors * (num_colors_needed // len(colors) + 1)
    
    colors = colors[:num_colors_needed]  # Ensure we have enough colors

    ax.pie(percentages, labels=stocks, autopct='%1.2f%%', startangle=140, colors=colors)
    ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    ax.set_title("Portfolio Allocation Percentage")

    # Clear the previous pie chart before adding the new one
    for widget in pie_frame.winfo_children():
        widget.destroy()
    
    canvas = FigureCanvasTkAgg(fig, master=pie_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

root.mainloop()
