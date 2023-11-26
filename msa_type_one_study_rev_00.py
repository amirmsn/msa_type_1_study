import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog
import scipy.stats as stats
from datetime import datetime
import os

# Constant values
TOLERANCE_PERCENTAGE = 20
SIGMA_LEVEL = 6
LEVEL_OF_SIGNIFICANCE = 0.05
CAPABILITY_CRITERION = 1.67
VARIABILITY_CRITERION = 15

# Function to clear the console before running the main parts of the code
os.system('cls')

# Get the current date and time
current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# Function to take the resolution of the measurement system from the user
def get_resolution():
    while True:
        resolution = input("Enter the resolution value (number of decimal points): ")
        if resolution.isdigit():
            return 1 / (10 ** int(resolution))
        print('Invalid input. Please try again!')


# Function to take the reference/target value from the user
def get_ref_val():
    while True:
        user_input = input("Enter the reference value: ")
        try:
            reference_value = float(user_input)
            return reference_value
        except ValueError:
            print('Invalid input. Please try again!')


# Function to take the upper value of the tolerance range
def get_up_tol():
    while True:
        user_input = input("Enter the upper value of the tolerance range: ")
        try:
            upper_tolerance = float(user_input)
            return upper_tolerance
        except ValueError:
            print('Invalid input. Please try again!')


# Function to take the lower value of the tolerance range
def get_lo_tol():
    while True:
        user_input = input("Enter the lower value of the tolerance range: ")
        try:
            lower_tolerance = float(user_input)
            return lower_tolerance
        except ValueError:
            print('Invalid input. Please try again!')


# Prompt the user to enter the resolution
resolution = get_resolution()

# Prompt the user to enter the reference value
reference_value = get_ref_val()

# Prompt the user to enter the upper value of the tolerance range
upper_tolerance = get_up_tol()

# Prompt the user to enter the lower value of the tolerance range
lower_tolerance = get_lo_tol()

# Create Tkinter root window
root = Tk()
root.withdraw()  # Hide the root window

# Prompt the user to select an Excel file
file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

# Read the data from the selected Excel file
data = pd.read_excel(file_path, header=None, names=["Value"])

# Create a range column
data["Range"] = range(1, len(data) + 1)

# Calculate the tolerance
tolerance = upper_tolerance - lower_tolerance

# Calculate the upper and lower limits for the 10% tolerance lines
upper_limit = reference_value + (TOLERANCE_PERCENTAGE / 200 * tolerance)
lower_limit = reference_value - (TOLERANCE_PERCENTAGE / 200 * tolerance)

# The mean of the records
sum_of_data = 0
for data_point in data["Value"]:
    sum_of_data += data_point
mean = sum_of_data / len(data)

# The standard deviation of the records
sum_of_squares = 0
for data_point in data["Value"]:
    sum_of_squares += (data_point - mean) ** 2
std_dev = (sum_of_squares / (len(data) - 1)) ** (1 / 2)

# Standard deviation spread
std_dev_spread = SIGMA_LEVEL * std_dev

# Bias
bias = abs(mean - reference_value)

# t-distribution value
t_value = (mean - reference_value) / (std_dev / (len(data) ** 0.5))

# P-value for two-tailed test
p_value = 2 * stats.t.sf(abs(t_value), 50)

# Cg
c_g = 0.2 * tolerance / (std_dev_spread)

# Cgk
c_g_k = (0.1 * tolerance - bias) / (std_dev_spread / 2)

# Percent variability (repeatability)
percent_var_repeat = 20 / c_g

# Percent variability (repeatability and bias)
percent_var_repeat_bias = 20 / c_g_k

# Create a figure and axes with space for annotations
fig, ax = plt.subplots(figsize=(10, 6))
fig.subplots_adjust(bottom=0.3)  # Adjust bottom margin for annotations

# Plot the run chart
plt.plot(data["Range"], data["Value"], marker="o")
plt.axhline(reference_value, color="green", linestyle="-")
plt.axhline(upper_limit, color="red", linestyle="-", label="Upper Limit")
plt.axhline(lower_limit, color="red", linestyle="-", label="Lower Limit")
plt.xlabel("Observation")
plt.ylabel("Record")
plt.title("Run Chart")

# Add text annotations at the top of the chart
plt.text(
    -0.1,
    1.1,
    "Type 1 Gage Study Report",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=18,
    weight='bold',
)

# Add current date and time at the top of the chart
plt.text(
    1.1,
    1.1,
    f"Date and time of the study: {current_datetime}",
    transform=plt.gca().transAxes,
    ha="right",
    fontsize=10,
)

# Add text annotations at the bottom of the chart
plt.text(
    0,
    -0.2,
    "Basic Statistics",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
    weight='bold'
)
plt.text(
    0,
    -0.25,
    f"Reference: {'{:.3e}'.format(reference_value)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0,
    -0.3,
    f"Mean: {'{:.3e}'.format(mean)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0,
    -0.35,
    f"StDev: {'{:.3e}'.format(std_dev)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0,
    -0.4,
    f"{SIGMA_LEVEL} x StDev (SV): {'{:.3e}'.format(std_dev_spread)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0,
    -0.45,
    f"Tolerance (Tol): {'{:.3e}'.format(tolerance)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0,
    -0.5,
    f"Resolution: {resolution}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)

plt.text(
    0.25,
    -0.2,
    "Bias",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
    weight='bold'
)
plt.text(
    0.25,
    -0.25,
    f"Bias: {'{:.3e}'.format(bias)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0.25,
    -0.3,
    f"T-value: {'{:.3e}'.format(t_value)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0.25,
    -0.35,
    f"P-value: {'{:.3f}'.format(p_value)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0.25,
    -0.4,
    "Null hypothesis: Bias = 0?",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
if p_value < LEVEL_OF_SIGNIFICANCE:
    plt.text(
        0.25,
        -0.45,
        "Statistically significant bias availabe",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
        color='red'
    )
else:
    plt.text(
        0.25,
        -0.45,
        "Bias NOT avaialbe",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
        color='green'
    )

plt.text(
    0.6,
    -0.2,
    "Capability",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
    weight='bold'
)
plt.text(
    0.6,
    -0.25,
    f"Cg: {'{:.3e}'.format(c_g)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0.6,
    -0.3,
    f"Cgk: {'{:.3e}'.format(c_g_k)}",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0.6,
    -0.35,
    f"%Var (Repeatability): {'{:.3f}'.format(percent_var_repeat)}%",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)
plt.text(
    0.6,
    -0.4,
    f"%Var (Repeatability and Bias): {'{:.3f}'.format(percent_var_repeat_bias)}%",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=10,
)

if c_g > CAPABILITY_CRITERION:
    plt.text(
        1,
        -0.25,
        "PASSED",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
        color='green'
    )
else:
    plt.text(
        1,
        -0.25,
        "FAILED",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
        color='red'
    )

if c_g_k > CAPABILITY_CRITERION:
    plt.text(
        1,
        -0.3,
        "PASSED",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
        color='green'
    )
else:
    plt.text(
        1,
        -0.3,
        "FAILED",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
        color='red'
    )

if percent_var_repeat > VARIABILITY_CRITERION:
    plt.text(
        1,
        -0.35,
        "Large var",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
    )
else:
    plt.text(
        1,
        -0.35,
        "Small var",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
    )

if percent_var_repeat_bias > VARIABILITY_CRITERION:
    plt.text(
        1,
        -0.4,
        "Large var",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
    )
else:
    plt.text(
        1,
        -0.4,
        "Small var",
        transform=plt.gca().transAxes,
        ha="left",
        fontsize=10,
    )

plt.text(
    0.75,
    -0.5,
    "Powered by EurDex Lab\u00AE (https://eurdexlab.se)",
    transform=plt.gca().transAxes,
    ha="left",
    fontsize=8,
)

plt.show()
