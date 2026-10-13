from pathlib import Path

import numpy as np, pandas as pd, matplotlib.pyplot as plt, torch
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

csv_path = Path(__file__).parent / "telco.csv" / "telco.csv"
dataset = pd.read_csv(csv_path)
""""
print("Missing values before cleaning:")
print(dataset.isna().sum()[dataset.isna().sum() > 0])

# These missing values mean the service/event did not apply to the customer.
dataset["Offer"] = dataset["Offer"].fillna("No Offer")
dataset["Internet Type"] = dataset["Internet Type"].fillna("No Internet")
dataset["Churn Category"] = dataset["Churn Category"].fillna("Not Applicable")
dataset["Churn Reason"] = dataset["Churn Reason"].fillna("Not Applicable")

print("\nMissing values after cleaning:", dataset.isna().sum().sum())
print(dataset.head())
print(dataset.shape)
dataset.info()
"""
missing = pd.DataFrame({
    "count": dataset.isna().sum(),
    "percent": dataset.isna().sum() / len(dataset) * 100
})
print(missing[missing["count"] > 0])

print(dataset.dtypes)

for col in dataset.select_dtypes(include="object"):
    print(col, dataset[col].value_counts(dropna=False).head(10))

print("Duplicates", dataset.duplicated().sum())
print(dataset.describe().T)
