#Octogenarian Data Analysis
#Barnabas Obeng-Gyasi
#November 3, 2023

pip install pandas
pip install scipy

import pandas as pd
import scipy.stats as stats
import tkinter as tk
from tkinter import filedialog

# Initialize tkinter
root = tk.Tk()
root.withdraw()  # Hide the main tkinter window

# Open a file dialog to select the Excel file
file_path = filedialog.askopenfilename(
    title="Select the Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

# Ensure a file was selected before proceeding
if file_path:
    # Load the data from the selected file
    df_stats = pd.read_excel(file_path, sheet_name='Stats')

    # Assuming the data frame df_stats is now loaded with the correct data
    # Convert the age groups into categories
    age_bins = [79, 84, 89, 94]  # Define age bins as 80-84, 85-89, 90-94
    age_labels = ['80-84', '85-89', '90-94']
    df_stats['Age Group'] = pd.cut(df_stats['Age at Surgery (Calculated: DOB & DOS)'], bins=age_bins, labels=age_labels, right=False)

    # ANOVA for the difference in survival rate among the age groups
    anova_results = stats.f_oneway(
        df_stats['Estimated survival in mo (assuming surviving patients died this month 11/1/2023)'][df_stats['Age Group'] == '80-84'],
        df_stats['Estimated survival in mo (assuming surviving patients died this month 11/1/2023)'][df_stats['Age Group'] == '85-89'],
        df_stats['Estimated survival in mo (assuming surviving patients died this month 11/1/2023)'][df_stats['Age Group'] == '90-94']
    )

    # Chi-Squared test for complication rates among different procedure types
    contingency_procedure = pd.crosstab(df_stats['Post-Surgery Complications (Y/N)'], df_stats['Fusion Status (Y/N)'])
    chi2_procedure, p_procedure, dof_procedure, expected_procedure = stats.chi2_contingency(contingency_procedure)

    # Chi-Squared test for sex difference in complication rates
    contingency_sex = pd.crosstab(df_stats['Post-Surgery Complications (Y/N)'], df_stats['Sex'])
    chi2_sex, p_sex, dof_sex, expected_sex = stats.chi2_contingency(contingency_sex)

    # Chi-Squared test for complications across different age groups
    contingency_age = pd.crosstab(df_stats['Post-Surgery Complications (Y/N)'], df_stats['Age Group'])
    chi2_age, p_age, dof_age, expected_age = stats.chi2_contingency(contingency_age)

    # Output the results
    results = {
        'ANOVA': {
            'F-statistic': anova_results.statistic,
            'p-value': anova_results.pvalue
        },
        'Chi-Squared': {
            'Procedure Type': {
                'Chi2': chi2_procedure,
                'p-value': p_procedure
            },
            'Sex': {
                'Chi2': chi2_sex,
                'p-value': p_sex
            },
            'Age Group': {
                'Chi2': chi2_age,
                'p-value': p_age
            }
        }
    }

    # Print the results
    for test in results:
        print(f"{test} Results:")
        for key, value in results[test].items():
            if isinstance(value, dict):
                print(f"  {key}:")
                for subkey, subvalue in value.items():
                    print(f"    {subkey}: {subvalue}")
            else:
                print(f"  {key}: {value}")
else:
    print("File selection was cancelled.")

# Close tkinter
root.destroy()
