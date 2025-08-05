import os
import glob
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
from scipy.signal import find_peaks
from openpyxl import Workbook
import sys 

input_directory = r"C:\Users\Justin G\Documents\Desktop\DARTS DATA\Input"
output_directory = r"C:\Users\Justin G\Documents\Desktop\DARTS DATA\Output"
output_excel_file = os.path.join(output_directory, "PPF_testing_with_R2.xlsx")

width_column = 'Lateral(µm)'
height_column = 'Total Profile(Å)'

def gaussian(x, A, x0, sigma, y0):
    """
    Gaussian function for curve fitting.
    A: Amplitude
    x0: Center of the peak
    sigma: Standard deviation (related to width)
    y0: Baseline offset
    """
    return A * np.exp(-((x - x0)**2) / (2 * sigma**2)) + y0

plt.figure(figsize=(12, 7)) 
plt.xlabel(width_column + ' (µm)')
plt.ylabel(height_column)
plt.title('Gaussian Curve Fits for Multiple Profiles')
plt.grid(True, linestyle='--', alpha=0.7)

# Prepare an Excel workbook for all metrics
wb = Workbook()
ws = wb.active
ws.title = "All_Metrics"
ws.append(["File Name", "Center (x0) (µm)", "Amplitude (A)", "Sigma (µm)",
           "Baseline (y0)", "FWHM (µm)", "Area", "R-squared"]) 

# LOOP THROUGH EACH EXCEL FILE 
try:
    # Get a list of all CSV files in the directory using the glob module 
    csv_files = glob.glob(input_directory + '/*.csv')

    # In case of failure
    if not csv_files:
        print(f"No CSV files found in '{input_directory}'. Please check the directory path.")
        sys.exit(1)

    for file_path in csv_files:
        file_name = os.path.basename(file_path)
        print(f"\n--- Processing file: {file_name} ---")

    # This reads each excel file starting from 28th row and uses first two columns only 
        try:
            df = pd.read_csv(file_path, skiprows=27, usecols=[width_column, height_column])

            x = df[width_column].values
            y = df[height_column].values

            # Locate peaks to get initial guesses for curve_fit
            # Using a height threshold to find prominent peaks
            peaks, _ = find_peaks(y, height=np.mean(y) + np.std(y) * 0.5) # Adjusted threshold slightly for robustness
            if peaks.size == 0:
                # If no peaks found with the initial threshold, try a lower one
                peaks, _ = find_peaks(y, height=np.mean(y))
                if peaks.size == 0:
                    print(f"Warning: No prominent peaks found for {file_name}; skipping fitting for this file.")
                    continue # Skip to the next file
            
            # Select the most prominent peak for initial guess
            peak_idx = peaks[np.argmax(y[peaks])]
            x0_guess = x[peak_idx]
            A_guess = y[peak_idx] - np.min(y) # Amplitude from peak to minimum
            sigma_guess = (x.max() - x.min()) / 10 # Rough estimate for sigma
            y0_guess = np.min(y) # Baseline guess

            p0 = [A_guess, x0_guess, sigma_guess, y0_guess] # Initial parameters for curve_fit

            try:
                params, covariance = curve_fit(gaussian, x, y, p0=p0)
                A_fit, x0_fit, sigma_fit, y0_fit = params
            except RuntimeError as e:
                print(f"Error: Curve fitting failed for {file_name}. {e}")
                print("This might be due to poor initial guesses or data that doesn't resemble a Gaussian. Skipping this file.")
                continue # Skip to the next file

            # Generate the fitted curve using the optimal parameters
            x_fit = np.linspace(x.min(), x.max(), 500) 
            y_fit = gaussian(x_fit, *params)

            # Goodness of Fit 
            # Predict y values based on the fitted parameters for the original x_data points
            y_predicted = gaussian(x, *params)
            
            # Calculate Sum of Squares Total (SStot)
            ss_tot = np.sum((y - np.mean(y)) ** 2)
            
            # Calculate Sum of Squares Residual (SSres)
            ss_res = np.sum((y - y_predicted) ** 2)
            
            # Calculate R-squared
            r_squared = 1 - (ss_res / ss_tot) if ss_tot > 0 else np.nan # Avoid division by zero


            # Plot raw data points for the current file
            plt.plot(x, y, '.')
            # Plot the fitted curve for the current file
            plt.plot(x_fit, y_fit, '-')
            
            # Write metrics to the excel file
            fwhm = 2 * np.sqrt(2 * np.log(2)) * abs(sigma_fit) # Use absolute value for sigma
            area = A_fit * abs(sigma_fit) * np.sqrt(2 * np.pi) # Use absolute value for sigma

            print(f"  Center: {x0_fit:.4f} µm")
            print(f"  FWHM: {fwhm:.4f} µm")
            print(f"  Area: {area:.4f} (Arbitrary Units)")
            print(f"  R-squared ($R^2$): {r_squared:.4f}")

            # Append metrics for the current file to the Excel worksheet
            ws.append([file_name, f"{x0_fit:.6f}", f"{A_fit:.6f}", f"{sigma_fit:.6f}",
                       f"{y0_fit:.6f}", f"{fwhm:.6f}", f"{area:.6f}", f"{r_squared:.6f}"])

        except KeyError as e:
            print(f"Error: Column '{e}' not found in {file_name}. Please check column names.")
            continue # Skip to the next file
        except Exception as e:
            print(f"An unexpected error occurred while processing {file_name}: {e}")
            continue # Skip to the next file

except FileNotFoundError:
    print(f"Error: Input directory '{input_directory}' not found.")
    sys.exit(1)
except Exception as e:
    print(f"An error occurred during file globbing or initial setup: {e}")
    sys.exit(1)

# FINAL PLOT 
plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0.) 
plt.tight_layout(rect=[0, 0, 0.8, 1]) 
plt.show()


# Ensure output directory is there 
os.makedirs(output_directory, exist_ok=True)
try:
    wb.save(output_excel_file)
    print(f"\nAll metrics written to {output_excel_file}")
except Exception as e:
    print(f"Error saving Excel file: {e}")

print("\n Script Finished ")