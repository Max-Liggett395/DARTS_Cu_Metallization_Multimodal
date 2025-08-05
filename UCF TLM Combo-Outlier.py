#TLM Script
# Import necessary libraries
import os                       # For interacting with the operating system (e.g., file paths)
import re                       # For using regular expressions
import pandas as pd             # For working with dataframes
import numpy as np              # For numerical operations
import matplotlib.pyplot as plt
from scipy.stats import linregress  # For performing linear regression
from openpyxl import Workbook   # For writing Excel files


# Define the hyperbolic cotangent function
def coth(x):
    return 1 / np.tanh(x)

# User-defined variables (in cm)
contact_width = 1.5
contact_spacing = 0.125
contact_length = 0.017

# Define input and output directories and output Excel file path
input_directory = r"C:\Users\quint\Documents\Processing_Scripts\TLM Data"
output_directory = r"C:\Users\quint\Documents\Processing_Scripts"
output_file = os.path.join(output_directory, "TLM_testing.xlsx")


# Calculate a list of distances based on TLM spacing formula
def calculate_distance(S3, S4, max_iter=9):
    return [i * S4 + (i - 1) * S3 for i in range(1, max_iter + 1)]

# Parse a single .txt file to extract relevant resistance measurement data
def parse_txt_file(file_path):
    data = []
    with open(file_path, 'r') as f:
        lines = f.readlines()

    sample_name = os.path.basename(file_path)  # Extract file name
    skip_lines = False

    for line in lines:
        # Skip data blocks with questionable quality
        if "Possible questionable voltage" in line or "Bad Current" in line:
            skip_lines = True
        elif "FingerLow" in line:
            skip_lines = False
        # Parse valid measurement lines
        elif not skip_lines and "RData" in line:
            parts = line.strip().split(',')
            if len(parts) >= 12:
                try:
                    # Extract necessary fields
                    f1 = parts[3]
                    f2 = parts[4]
                    amps = float(parts[8])
                    actual_v = float(parts[9])
                    actual_i = float(parts[10])
                    r = float(parts[11])
                    data.append({
                        "Sample_Name": sample_name,
                        "F1_F2": f"{f1}-{f2}",
                        "Amps": amps,
                        "Actual_V": actual_v,
                        "Actual_I": actual_i,
                        "R": r
                    })
                except ValueError:
                    continue  # Skip lines with invalid numerical data
    return pd.DataFrame(data)

# Calculate sheet resistance (RSH) using linear regression of R vs. distance
def calculate_rsh(resistance_values, distances, contact_width):
    resistance_values = resistance_values.reset_index(drop=True)
    distances = pd.Series(distances).reset_index(drop=True)
    
    # Align lengths of resistance and distance arrays
    min_length = min(len(resistance_values), len(distances))
    resistance_values = resistance_values[:min_length]
    distances = distances[:min_length]

    # Only include non-NaN values for regression
    mask = ~resistance_values.isna()
    if mask.sum() > 1:
        slope, _, _, _, _ = linregress(distances[mask], resistance_values[mask])
        return slope * contact_width
    return np.nan

# Calculate contact resistance Rcf from the intercept of R vs. distance
def calculate_rcf(resistance_values, distances):
    n = len(resistance_values)
    distances_subset = distances[:n]
    slope, intercept, _, _, _ = linregress(distances_subset, resistance_values)
    return intercept / 2

# Calculate transfer length LT using the inverse relationship
def calculate_lt(resistance_values, distances):
    n = len(resistance_values)
    resistance_subset = resistance_values[:n]
    slope, intercept, _, _, _ = linregress(resistance_subset, distances[:n])
    return intercept / -2

# Calculate specific contact resistivity pc
def calculate_pc(Rcf, LT, contact_width, contact_length):
    if LT == 0:
        return np.nan
    return (Rcf * LT * contact_width / coth(contact_length / LT)) * 1000  # Convert to mOhm-cm^2

# Calculate sheet resistivity rho_s
def calculate_rho_s(RSH):
    return RSH * 0.02  # Multiplies by semiconductor thickness

# Calculate eta (efficiency factor)
def calculate_eta(pc, rho_s):
    if rho_s * 0.02 == 0:
        return np.nan
    return (pc / 1000) / (rho_s * 0.02)

# Calculate corrected contact resistivity pc'
def calculate_pc_prime(pc, rho_s):
    return pc + 0.19 * rho_s * 1000

def detect_resistivity_outliers(df, cols):
    
    for col in cols:
        df = df[df[col] >= 0]
        Q1 = df[col].quantile(0.25)
        Q3 = df[col].quantile(0.75)
        IQR = Q3 -Q1
        lower, upper = Q1 - 1.5 * IQR, Q3 + 1.5 * IQR
        df[f"{col}_outlier"] = (df[col] < lower) | (df[col] > upper)
    return df

def show_boxplots_in_spyder(df): #show the boxplots within spyder
    parameters = {
        "pc": {"label": "Specific Contact Resistivity", "ylabel": "pc (mΩ·cm²)"},
        "RSH": {"label": "Sheet Resistance", "ylabel": "RSH (Ω/□)"},
        "LT": {"label": "Transfer Length", "ylabel": "LT (cm)"},
        "Rcf": {"label": "Contact Resistance", "ylabel": "Rcf (Ω)"}
    }

    for resistance_type, plot_info in parameters.items():
        data = df[resistance_type].dropna()
        if data.empty:
            print(f" No data for {resistance_type}, skipping plot.")
            continue

        plt.figure(figsize=(6, 4))
        plt.boxplot(data, labels=[plot_info["label"]])
        plt.ylabel(plot_info["ylabel"]) # Set y-axis label
        plt.title(f"{plot_info['label']}") # Set plot title
        plt.grid(True) #show grid
        plt.tight_layout() #adjust spacing 
        plt.show()
        
# Outlier Plot function
def plot_outliers_vs_all(df, metrics):
    x = np.arange(len(df))
    for col in metrics:
        y = df[col].values
        outliers = df[df[f"{col}_outlier"]]

        fig, axs = plt.subplots(1, 2, figsize=(12, 4))

        # Full data with regression
        mask = ~np.isnan(y)
        if mask.sum() >= 2:
            slope, intercept = np.polyfit(x[mask], y[mask], 1)
            y_fit = slope * x + intercept
        else:
            y_fit = np.full_like(x, np.nan)

        axs[0].scatter(x, y, label="All Data")
        axs[0].plot(x, y_fit, color="green", label="Fit Line")
        axs[0].scatter(df.index[df[f"{col}_outlier"]], df.loc[df[f"{col}_outlier"], col],
                       color="red", marker='x', label="Outliers")
        axs[0].set_title(f"{col} (All)")
        axs[0].set_xlabel("Sample Index")
        axs[0].set_ylabel(col)
        axs[0].legend()
        axs[0].grid(True)

        # Outliers only
        axs[1].scatter(outliers.index, outliers[col], color="red", label="Outliers Only")
        axs[1].set_title(f"{col} (Outliers Only)")
        axs[1].set_xlabel("Sample Index")
        axs[1].set_ylabel(col)
        axs[1].legend()
        axs[1].grid(True)

        plt.tight_layout()
        plt.show()


# Main function to process all files and export results
def process_files(input_directory, output_directory):
    all_data = []         # Store parsed data from all files
    file_averages = []    # Store average resistance per contact pair for each file

    # Loop through all .txt files in the input directory
    for file in os.listdir(input_directory):
        if file.endswith(".txt"):
            file_path = os.path.join(input_directory, file)
            df = parse_txt_file(file_path)
            all_data.append(df)

            # Group by contact pair and average the resistance values
            if not df.empty:
                avg = df.groupby("F1_F2")["R"].mean().reset_index()
                avg["File"] = re.sub(r"_t\d+\.txt$", "", file)  # Clean filename
                file_averages.append(avg)
          


    # Combine all parsed data into a single dataframe
    all_data_df = pd.concat(all_data, ignore_index=True)

    # Generate calculated distances and corresponding contact pairs
    calculated_distances = calculate_distance(contact_length, contact_spacing)
    contact_pairs = [f"1 to {i+2}" for i in range(len(calculated_distances))]

    calculated_data = pd.DataFrame({
        "Contact pair": contact_pairs,
        "Calculated Distance": calculated_distances
    })

    # Combine average resistances for all files
    file_averages_df = pd.concat(file_averages, ignore_index=True)
    final_averages = file_averages_df.pivot_table(index="F1_F2", columns="File", values="R").reset_index()

    # Compute resistivity parameters for each sample (column in final_averages)
    resistivity_data = []
    for file in final_averages.columns[1:]:  # Skip the index column "F1_F2"
        resistances = final_averages[file]
        rsh = calculate_rsh(resistances, pd.Series(calculated_distances), contact_width)
        rcf = calculate_rcf(resistances.dropna(), calculated_distances)
        lt = calculate_lt(resistances.dropna(), calculated_distances)
        pc = calculate_pc(rcf, lt, contact_width, contact_length)
        rho_s = calculate_rho_s(rsh)
        eta = calculate_eta(pc, rho_s)
        pc_prime = calculate_pc_prime(pc, rho_s)

        resistivity_data.append({
            "File": file,
            "RSH": rsh,
            "Rcf": rcf,
            "LT": lt,
            "pc": pc,
            "rho_s": rho_s,
            "eta": eta,
            "pc_prime": pc_prime
        })

    # Compile all resistivity results into a dataframe
    resistivity_df = pd.DataFrame(resistivity_data)
    # Detect outliers in key resistivity metrics
    metrics = ["RSH", "Rcf", "LT", "pc"]
    resistivity_df = detect_resistivity_outliers(resistivity_df, metrics)

    return resistivity_df
resistivity_df = process_files(input_directory, output_directory)

# If you want Plotly interactivity:
import plotly.express as px
resistivity_df = resistivity_df.reset_index(drop=True)
resistivity_df["Sample_Index"] = resistivity_df.index
metrics = ["RSH", "Rcf", "LT","pc"]

for metric in ["RSH", "Rcf", "LT","pc"]:
    outlier_col = f"{metric}_outlier"
    color_col = f"{metric}_Color"
    resistivity_df[color_col] = np.where(resistivity_df[outlier_col], "red", "blue")

for col in metrics:
    fig = px.scatter(
        resistivity_df,
        x="Sample_Index",
        y=col, 
        color=f"{col}_Color",
        trendline="ols",
        hover_data=["File", col],
        title=f"Interactive {col} vs. Sample Index"
    )
    fig.update_layout(width=800, height=450)
    fig.show(renderer="browser")
# Run script
if __name__ == "__main__":
    df = process_files(input_directory, output_directory)
    show_boxplots_in_spyder(df)
    plot_outliers_vs_all(df, ["RSH", "Rcf", "LT", "pc"])