**Short-Term Load Forecasting for 11kV Distribution Feeder Using XGBoost: A Case Study of Khaboorah Substation in Oman**

Abdul Saleem Shaik

*Department of Electrical and Electronics Engineering
University of Technology and Applied Sciences (UTAS)
Sohar, Oman
abdulsaleem.shaik@utas.edu.om*

***Abstract***

*Accurate short-term load forecasting (STLF) is essential for efficient operation and planning of electrical distribution systems. This paper presents a comprehensive machine learning approach for predicting hourly load on the 11kV Khaboorah (KHBR) distribution feeder in Oman. A real-world dataset spanning March 2025 to August 2025 was collected from SCADA systems, comprising hourly current readings from eight feeder lines along with engineered temporal, lag, and weather features. Four machine learning models were evaluated: XGBoost, LightGBM, Random Forest, and Ridge Regression. Feature engineering incorporated rolling statistics, multiple lag features (1-168 hours), cyclical time encodings, and weather parameters including temperature, humidity, and Ramadan period indicators. Results demonstrate that XGBoost achieved superior performance with a test R2 of 0.8094, RMSE of 6.827A, and MAPE of 3.25%, outperforming other models by 1.3-20.7% in R2 score. Feature importance analysis revealed that rolling mean statistics and 24-hour lag features are the most significant predictors. The study also identified substantial load reduction during Ramadan periods (median ~75A versus ~160A during non-Ramadan), providing valuable insights for demand-side management in regions with similar cultural patterns.*

***Keywords: ****Short-term load forecasting, XGBoost, machine learning, distribution feeder, feature engineering, Ramadan load patterns*

**I. INTRODUCTION**

**A. Background and Motivation**

The increasing complexity of modern electrical distribution networks demands accurate load forecasting for effective grid management and operational planning. Short-term load forecasting (STLF) plays a crucial role in various aspects of power system operations, including generation scheduling, demand-side management, equipment maintenance planning, and grid stability enhancement [1]. At the distribution level, particularly for 11kV feeders serving residential and commercial loads, accurate predictions are essential for preventing equipment overloading, reducing operational costs, and maintaining voltage stability [2].

The Sultanate of Oman has experienced rapid growth in electricity demand, driven by economic development, population growth, and increasing urbanization. The distribution network faces unique challenges including extreme summer temperatures exceeding 45C, cultural factors such as Ramadan affecting consumption patterns, and seasonal variations in demand [3]. Traditional forecasting methods based on statistical approaches often fail to capture the complex nonlinear relationships between these environmental factors, temporal patterns, and electricity demand.

**B. Research Objectives and Contributions**

This study addresses the challenge of accurate short-term load forecasting for an 11kV distribution feeder by leveraging advanced machine learning techniques. The main contributions include: (1) Development of a comprehensive ML framework for hourly load forecasting using real operational data from March to August 2025; (2) Systematic comparison of four ML algorithms; (3) Extensive feature engineering incorporating temporal patterns, lag features, rolling statistics, and weather variables; (4) Analysis of Ramadan period impact on load consumption patterns; (5) Identification of the most influential features through comprehensive feature importance analysis.

**II. LITERATURE REVIEW**

**A. Machine Learning in Load Forecasting**

Machine learning techniques have emerged as powerful tools for electrical load forecasting, demonstrating superior performance compared to traditional statistical methods. Hong and Fan [1] provided a comprehensive review of probabilistic load forecasting methods. Recent studies have shown that gradient boosting methods, particularly XGBoost, achieve excellent results in structured data problems including load forecasting [4]. Ke et al. [6] introduced LightGBM, a gradient boosting framework using histogram-based algorithms for efficient training.

**B. Feature Engineering for Load Forecasting**

The importance of feature engineering in load forecasting has been extensively documented. Haben et al. [7] reviewed short-term load forecasting methods, emphasizing the significance of incorporating calendar variables, weather data, and customer behavior patterns. Studies by Wang et al. [8] demonstrated that lag features and rolling statistics significantly improve forecast accuracy.

**C. Distribution-Level Load Forecasting**

Load forecasting at the distribution level presents unique challenges compared to system-level forecasting. The higher variability at lower aggregation levels and limited historical data availability require specialized approaches [11]. Studies specific to the GCC region have highlighted the importance of considering local factors such as extreme temperatures and religious observances [13].

**III. METHODOLOGY**

**A. Data Collection and Description**

The dataset was collected from the Khaboorah (KHBR) 33/11kV distribution substation located in the North Batinah governorate of Oman. Hourly load current readings were obtained from the SCADA system for March 2025 to August 2025. The dataset includes measurements from eight 11kV feeder lines designated as KHBR01_K_LN01 through KHBR01_K_LN08. The raw dataset contains 4,392 hourly observations (183 days x 24 hours).

**B. Feature Engineering**

Temporal Features: Hour of day (0-23), day of week (0-6), day of month, week of year, and month. Sine and cosine encodings capture cyclical patterns.

Lag Features: Historical load values at 1, 2, 3, 6, 12, 24, 48, and 168 hours prior to the forecast time.

Rolling Statistics: Moving window mean and standard deviation for 6, 12, 24, and 48 hour windows.

Weather Features: Temperature (C) and relative humidity (%). A binary indicator identifies the Ramadan period (March 1-30, 2025).

**C. Machine Learning Models**

Four algorithms were evaluated: XGBoost - optimized gradient boosting with regularization [4]; LightGBM - histogram-based gradient boosting [6]; Random Forest - ensemble of decision trees [15]; Ridge Regression - regularized linear baseline.

**D. Model Evaluation Metrics**

Performance metrics include: Mean Absolute Error (MAE), Root Mean Square Error (RMSE), Coefficient of Determination (R2), and Mean Absolute Percentage Error (MAPE). The dataset was split 80/20 chronologically.

**IV. RESULTS**

**A. Data Exploration**

The load distribution exhibits a bimodal pattern with peaks at 70-80A and 160-170A. Clear diurnal patterns show peak loads at 14:00 hours when air conditioning demand is highest. Temperature ranged from 25C in March to 44C in August.

**B. Model Performance Comparison**

| Model | Train MAE | Test MAE | Train RMSE | Test RMSE |
| --- | --- | --- | --- | --- |
| XGBoost | 1.600 | 4.923 | 2.156 | 6.827 |
| LightGBM | 3.048 | 5.061 | 4.232 | 7.015 |
| Random Forest | 3.910 | 6.149 | 6.087 | 8.565 |
| Ridge Regression | 10.598 | 7.153 | 14.476 | 8.978 |

**TABLE I: MODEL PERFORMANCE COMPARISON (MAE AND RMSE)**

| Model | Train R2 | Test R2 | Test MAPE |
| --- | --- | --- | --- |
| XGBoost | 0.9965 | 0.8094 | 3.25% |
| LightGBM | 0.9867 | 0.7988 | 3.35% |
| Random Forest | 0.9725 | 0.7000 | 4.04% |
| Ridge Regression | 0.8443 | 0.6704 | 4.59% |

**TABLE II: R2 SCORES AND MAPE COMPARISON**

XGBoost achieved the best overall performance with test R2 of 0.8094 and RMSE of 6.827A. LightGBM demonstrated comparable performance with test R2 of 0.7988.

**C. Feature Importance Analysis**

| Rank | Feature | Importance | Category |
| --- | --- | --- | --- |
| 1 | rolling_mean_6 | 0.284 | Rolling Statistics |
| 2 | lag_24 | 0.198 | Lag Feature |
| 3 | rolling_mean_12 | 0.142 | Rolling Statistics |
| 4 | rolling_mean_24 | 0.089 | Rolling Statistics |
| 5 | lag_168 | 0.072 | Lag Feature |
| 6 | lag_48 | 0.058 | Lag Feature |
| 7 | rolling_std_24 | 0.041 | Rolling Statistics |
| 8 | hour_sin | 0.034 | Temporal |
| 9 | temperature | 0.028 | Weather |
| 10 | lag_12 | 0.024 | Lag Feature |

**TABLE III: TOP 10 FEATURE IMPORTANCE RANKINGS (XGBOOST)**

Rolling statistics (6-hour rolling mean) emerged as most influential (28.4%). The 24-hour lag feature ranked second, reflecting strong daily periodicity.

**D. Weather and Ramadan Impact Analysis**

During Ramadan (March 1-30, 2025), median load was approximately 75A versus 160A during non-Ramadan periods - a 53% reduction reflecting changes in daily routines and commercial activity.

**V. DISCUSSION**

The superior XGBoost performance stems from effective nonlinear relationship capture, regularization mechanisms, and handling of mixed feature types. The 1.3% gap between XGBoost and LightGBM suggests both are suitable. The demonstrated MAPE of 3.25% meets operational requirements for load balancing, transformer assessment, and maintenance scheduling.

**VI. CONCLUSION**

This paper presented a comprehensive machine learning approach for short-term load forecasting at the 11kV Khaboorah feeder in Oman. XGBoost achieved the best performance (R2=0.8094, RMSE=6.827A, MAPE=3.25%). Feature engineering with rolling statistics and lag features significantly improves accuracy. Ramadan period analysis revealed 53% load reduction, providing valuable demand-side management insights.

**ACKNOWLEDGMENT**

The author thanks Mazoon Electricity Company (MZEC) for SCADA data access. Support from University of Technology and Applied Sciences (UTAS), Sohar campus, is gratefully acknowledged.

**REFERENCES**

[1] T. Hong and S. Fan, "Probabilistic electric load forecasting: A tutorial review," Int. J. Forecasting, vol. 32, no. 3, pp. 914-938, 2016.

[2] H. K. Alfares and M. Nazeeruddin, "Electric load forecasting: Literature survey," Int. J. Systems Science, vol. 33, no. 1, pp. 23-34, 2002.

[3] Authority for Public Services Regulation, "Oman Electricity and Water Sector Annual Report 2024," Muscat, 2024.

[4] T. Chen and C. Guestrin, "XGBoost: A scalable tree boosting system," Proc. ACM SIGKDD, 2016, pp. 785-794.

[5] J. Zheng et al., "Electric load forecasting using LSTM," Proc. CISS, 2017, pp. 1-6.

[6] G. Ke et al., "LightGBM: Efficient gradient boosting," NIPS, 2017, pp. 3146-3154.

[7] S. Haben et al., "Short term load forecasting at low voltage level," Int. J. Forecasting, vol. 35, pp. 1469-1484, 2019.

[8] Y. Wang et al., "Conditional residual modeling for load forecasting," IEEE Trans. Power Syst., vol. 33, pp. 7327-7330, 2018.

[9] A. Al-Musawi et al., "Short-term load forecasting with weather data," Applied Energy, vol. 259, 2020.

[10] S. Rahman and H. Zareipour, "Data-driven load forecasting," Electric Power Syst. Research, vol. 175, 2019.

[11] R. Sevlian and R. Rajagopal, "Scaling law for load forecasting," Int. J. Electrical Power, vol. 98, pp. 350-361, 2018.

[12] H. Shi et al., "Deep learning for household load forecasting," IEEE Trans. Smart Grid, vol. 9, pp. 5271-5280, 2018.

[13] H. M. Al-Hamadi and S. A. Soliman, "Short-term load forecasting based on Kalman filtering," Electric Power Syst. Research, vol. 68, pp. 47-59, 2004.

[14] B. Liu et al., "Probabilistic load forecasting via quantile regression," IEEE Trans. Smart Grid, vol. 8, pp. 730-737, 2017.

[15] L. Breiman, "Random forests," Machine Learning, vol. 45, pp. 5-32, 2001.
