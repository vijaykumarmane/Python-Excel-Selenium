import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pmdarima.model_selection import train_test_split as time_train_test_split
from xgboost import XGBRegressor
from sklearn import metrics

### Reads first sheet
data_df = pd.read_excel('Input Test Data ML.xlsx')

#### Significance
# data_df = data_df.fillna(0.0)

def mean_absolute_percentage_error(y_true, y_pred):
  
    y_true, y_pred = np.array(y_true), np.array(y_pred)
    
    ##### Gives RuntimeWarning: divide by zero encountered in true_divide
    return np.mean(np.abs((y_true - y_pred) / y_true)) * 100;

def timeseries_evaluation_metrics_func(y_true, y_pred):
    mse = metrics.mean_squared_error(y_true, y_pred)
    mae = metrics.mean_absolute_error(y_true, y_pred)
    rmse = np.sqrt(metrics.mean_squared_error(y_true, y_pred))
    mape= mean_absolute_percentage_error(y_true, y_pred)
    r2 = metrics.r2_score(y_true, y_pred)

    return {"Mean Square Error":mse, "Mean Absolute Error":mae, "Root Mean Squared Error": rmse, "Mean Absolute Percentage Error":mape, "R squared":r2}


for row in range(data_df.shape[0]):
    
    sku_item = data_df.iloc[row,0]
    
    # To take date as a index and actuals as a column
    actuals = list(data_df.iloc[row, 2:])
    index = [i for i in data_df.columns[2:]]

    df =  pd.DataFrame(data = actuals, index= index)

    # change column name
    df = df.rename(columns={0:"Actuals"})
    df.index.name = "Date"
    
    # To add Features
    df['Date'] = df.index
    df['Quarter'] = df['Date'].dt.quarter
    df['Month'] = df['Date'].dt.month

    # Lag 1 and 12
    df['Lag 1'] = df['Actuals'].shift(1)
    df['Lag 12'] = df['Actuals'].shift(12)

    # Rolling Simple Moving Average 12
    df['SMA(12)'] = df['Actuals'].rolling(window=12).mean()
    df['SMA(12)'] = df['SMA(12)'].shift(1)
    df = df.dropna()
    df = df[['Actuals','Quarter','Month','Lag 1', 'Lag 12', 'SMA(12)']].copy()
    
    data_from_row = df.shape[0]

    # Test Size = 20%
    train_df, test_df = time_train_test_split(df, test_size=int(len(df)*0.2))

    train_df = pd.DataFrame(train_df)
    test_df = pd.DataFrame(test_df)
    train_df_copy = train_df.copy()
    test_df_copy = test_df.copy()

    trainX, trainY = train_df_copy.iloc[:,1:], train_df_copy['Actuals']
    testX, testY = test_df_copy.iloc[:,1:], test_df_copy['Actuals']
    
    xgb = 0
    xgb = XGBRegressor(objective= 'reg:squarederror', n_estimators=1000);

    xgb.fit(trainX, trainY,
            eval_set=[(trainX, trainY), (testX, testY)],
            early_stopping_rounds=50,
            verbose=False); # Change verbose to True if you want to see it train
    
    predicted_results = xgb.predict(testX);
    
    accuracy_metrics = timeseries_evaluation_metrics_func(testY, predicted_results)
    
    feature_importance = xgb.get_booster().get_score(importance_type="weight");
    
    data_dict = {}
    data_dict = accuracy_metrics.copy()
    data_dict.update(feature_importance)
        
    # if data_dict["R squared"] > .2:
    #     print("R squared", data_dict["R_squared"], end="\n")
    for i in range(12):
        # To iterate over one future feature
        next_date = df.index[-1] + relativedelta(months=1)
        predicted = 0
        data = {'Actuals': [predicted], 'Quarter':[next_date.quarter], 'Month':[next_date.month],
        'Lag 1':[df['Actuals'][-1]], 'Lag 12':[df['Actuals'][-12]],
                'SMA(12)':[(df['Actuals'][-12:].sum()) / 12]}

        # Input dataframe
        input_df = pd.DataFrame(data=data, index=[next_date]).iloc[:,1:]

        predicted = xgb.predict(input_df);

        # Append with predicted
        df.loc[next_date] = predicted[0], next_date.quarter, next_date.month, df['Actuals'][-1], df['Actuals'][-12], (df['Actuals'][-12:].sum()) / 12

    df_forecast_transpose = df.iloc[data_from_row:][["Actuals"]].transpose()
    df_forecast_transpose = df_forecast_transpose.rename(index={'Actuals': sku_item})

    if row > 0:
        df_forecast = df_forecast.append(df_forecast_transpose)
        df_metrics = df_metrics.append(pd.DataFrame(data=data_dict, index= [sku_item]))
    else:
        df_forecast = df_forecast_transpose.copy()
        df_metrics = pd.DataFrame(data=data_dict, index= [sku_item])
        
df_forecast.join(df_metrics).to_excel("Output Forecast with Metrics.xlsx")
