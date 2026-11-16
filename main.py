# FastAPI app that serves the loan default model.
# Start it with:  uvicorn main:app --reload
# Then open http://127.0.0.1:8000/docs to try it in the browser.
#
# Sample request body for POST /predict:
# {
#   "AGE": 3,
#   "EMPLOY": 16,
#   "ADDRESS": 10,
#   "DEBTINC": 9.1,
#   "CREDDEBT": 11.3,
#   "OTHDEBT": 5.01
# }

import joblib
import pandas as pd
from fastapi import FastAPI
from pydantic import BaseModel

app = FastAPI(title="Bank Loan Default Prediction")

# load the model once at startup, not on every request
bundle = joblib.load("model.pkl")
model = bundle["model"]
features = bundle["features"]


# one field per training feature, same names as the columns
class Customer(BaseModel):
    AGE: int
    EMPLOY: float
    ADDRESS: float
    DEBTINC: float
    CREDDEBT: float
    OTHDEBT: float


@app.get("/")
def home():
    return {"message": "Loan default API is running. POST to /predict, or see /docs.",
            "features": features}


@app.post("/predict")
def predict(customer: Customer):
    # build a one-row frame with the columns in the order the model was trained on
    row = pd.DataFrame([customer.model_dump()])[features]
    prob = model.predict_proba(row)[0, 1]
    return {"probability_of_default": round(float(prob), 4)}
