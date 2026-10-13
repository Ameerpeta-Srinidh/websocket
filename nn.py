import numpy as np, pandas as pd, matplotlib.pyplot as plt,torch
from ucimlrepo import fetch_ucirepo 
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
  
# fetch dataset 
heart_disease = fetch_ucirepo(id=45) 
  
# data (as pandas dataframes) 
X = heart_disease.data.features 
y = heart_disease.data.targets 
  
# metadata 
print(heart_disease.metadata) 
  
# variable information 
print(heart_disease.variables) 

print(X.shape)

print(heart_disease.data.features.head())

X = X.fillna(X.mean(numeric_only=True))
y = (y.squeeze() > 0).astype(np.float32)

X_train, X_test, y_train, y_test = train_test_split(X,y, test_size = 0.2, random_state = 42)
scaler = StandardScaler()
X_train = scaler.fit_transform(X_train)
X_test = scaler.fit_transform(X_test)

X_train = torch.tensor(X_train, dtype=torch.float32)
X_test = torch.tensor(X_test, dtype=torch.float32)

y_train = torch.tensor(y_train.values, dtype = torch.float32)
y_test = torch.tensor(y_test.values, dtype = torch.float32)

import torch.nn as nn

class HeartNet(nn.Module):
    def __init__(self, input_size):
        super().__init__()
        self.network = nn.Sequential(
            nn.Linear(input_size,32),
            nn.ReLU(),

            nn.Linear(32,16),
            nn.ReLU(),

            nn.Linear(16,1),
            nn.Sigmoid(),

        )
    def forward(self,x):
        return self.network(x)
    
input_size = X_train.shape[1]
model = HeartNet(input_size)

criterion = nn.BCELoss()
optimizer = torch.optim.Adam(
    model.parameters(),
    lr = 0.001
)

epochs = 100

for epoch in range(epochs):
    outputs = model(X_train)
    loss = criterion(
        outputs.squeeze(),
        y_train
    )

    optimizer.zero_grad()
    loss.backward()
    optimizer.step()

    if epoch % 10 == 0:
        print(
            f"Epoch {epoch}, Loss: {loss.item():.4f}"
        )
model.eval()
with torch.no_grad():
    predictions = model(X_test)
    predictions = (
        predictions.squeeze() > 0.5
    ).float()

    accuracy = (
        predictions == y_test
    ).float().mean()

    print(f"Accuracy:{accuracy:.4f}"
    )
