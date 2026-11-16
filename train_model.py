# Trains the loan default model and scores the test set.
# Run this once before starting the API: it writes model.pkl and test_predictions.xlsx.

import pandas as pd
import joblib
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import roc_auc_score

# both files sit next to this script
TRAIN_FILE = "BANK LOAN.csv"
TEST_FILE = "BANK LOAN_TEST.csv"

# load the training data
df = pd.read_csv(TRAIN_FILE)
print("shape:", df.shape)
print("\nmissing values:")
print(df.isnull().sum())
print("\nDEFAULTER counts:")
print(df["DEFAULTER"].value_counts())

# SN is just a row number, and DEFAULTER is the target
y = df["DEFAULTER"]
X = df.drop(columns=["SN", "DEFAULTER"])

# all the columns should be numeric already, so nothing to encode
print("\nfeature dtypes:")
print(X.dtypes)
non_numeric = list(X.select_dtypes(exclude="number").columns)
if non_numeric:
    raise SystemExit(f"unexpected text columns, these need encoding first: {non_numeric}")

# quick hold-out so I know the model is actually sensible before saving it
X_tr, X_val, y_tr, y_val = train_test_split(X, y, test_size=0.25, stratify=y, random_state=42)
check = RandomForestClassifier(n_estimators=500, random_state=42).fit(X_tr, y_tr)
print("\nhold-out AUC:", round(roc_auc_score(y_val, check.predict_proba(X_val)[:, 1]), 4))

# now fit on everything for the model we actually ship
model = RandomForestClassifier(n_estimators=500, random_state=42)
model.fit(X, y)

print("\nfeature importance:")
print(pd.Series(model.feature_importances_, index=X.columns).sort_values(ascending=False).round(4))

# save the model together with the column order, the API has to rebuild rows the same way
joblib.dump({"model": model, "features": list(X.columns)}, "model.pkl")
print("\nsaved model.pkl with features:", list(X.columns))

# score the test set and write it out for Excel
test = pd.read_csv(TEST_FILE)
X_test = test[list(X.columns)]  # same columns, same order
test["Prob_Default"] = model.predict_proba(X_test)[:, 1].round(4)

test.to_excel("test_predictions.xlsx", index=False)
print("\nsaved test_predictions.xlsx")
print(test)
