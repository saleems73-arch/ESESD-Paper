**Short-Term Load Forecasting for 11kV Distribution Feeder Using XGBoost: A Case Study of Khaboorah Substation in Oman**

Abdul Saleem Shaik

*Department of Electrical and Electronics Engineering*

*University of Technology and Applied Sciences (UTAS)*

*Sohar, Oman*

*abdulsaleem.shaik@utas.edu.om*

***Abstract***

*Short-term load forecasting (STLF) is essential for operational planning in electrical distribution systems, including generation scheduling, demand-side management, and grid stability enhancement. This paper presents an optimized XGBoost machine learning model for accurate load prediction at an 11kV distribution feeder in Khaboorah, Oman. A comprehensive dataset spanning March 2025 to August 2025 was collected from SCADA measurements, incorporating hourly load readings from eight feeder lines, temporal features, and weather parameters. The proposed methodology employs sophisticated feature engineering including lag variables (1-168 hours), rolling statistics, and cyclical time encodings. Four machine learning models—XGBoost, LightGBM, Random Forest, and Ridge Regression—were systematically evaluated. Results demonstrate that the tuned XGBoost model achieved superior performance with a Root Mean Square Error (RMSE) of 6.827 A, R² score of 0.8094, and Mean Absolute Percentage Error (MAPE) of 3.25%. Feature importance analysis revealed that rolling mean statistics and 24-hour lag features are the dominant predictors. Additionally, the study identified a significant 53% reduction in load during the Ramadan period, highlighting the importance of cultural factors in regional load forecasting applications.*

***Keywords: ****Short-term load forecasting, XGBoost, machine learning, distribution feeder, feature engineering, Ramadan load patterns*

**I. INTRODUCTION**

*A. Background and Motivation*

The increasing complexity of modern electrical distribution networks demands accurate load forecasting for effective grid management and operational planning. Short-term load forecasting (STLF) plays a crucial role in various aspects of power system operations, including generation scheduling, demand-side management, equipment maintenance planning, and grid stability enhancement [4]. At the distribution level, particularly for 11kV feeders serving residential and commercial loads, accurate predictions are essential for preventing equipment overloading, reducing operational costs, and maintaining voltage stability [2].

The Sultanate of Oman has experienced rapid growth in electricity demand, driven by economic development, population growth, and increasing urbanization. The distribution network faces unique challenges including extreme summer temperatures exceeding 45°C, cultural factors such as Ramadan affecting consumption patterns, and seasonal variations in demand [15]. Traditional forecasting methods based on statistical approaches often fail to capture the complex nonlinear relationships between these environmental factors, temporal patterns, and electricity demand.

The Gulf Cooperation Council (GCC) region, including Oman, faces unique challenges in load forecasting due to extreme summer temperatures reaching 44°C, high air conditioning penetration, and significant load variations during religious observances such as Ramadan [6], [7]. Traditional forecasting methods based on linear regression or simple time series models fail to capture these complex nonlinear relationships between environmental factors, cultural patterns, and electricity demand.

*B. Research Objectives and Contributions*

This study addresses the challenge of accurate short-term load forecasting for an 11kV distribution feeder by leveraging advanced machine learning techniques. The main contributions of this paper are:

1) Development of an optimized XGBoost-based STLF model specifically designed for 11kV distribution feeders in the GCC region.

2) Comprehensive feature engineering incorporating lag variables, rolling statistics, and temporal encodings to capture load patterns at multiple time scales.

3) Systematic comparison of four machine learning algorithms with quantitative performance evaluation.

4) Analysis of Ramadan's impact on load patterns, demonstrating a 53% reduction during the holy month.

**II. LITERATURE REVIEW**

*A. Machine Learning in Load Forecasting*

Machine learning techniques have emerged as powerful tools for electrical load forecasting, demonstrating superior performance compared to traditional statistical methods. Hong and Fan [4] provided a comprehensive review of probabilistic load forecasting methods, establishing benchmarks for forecast evaluation. Recent studies have shown that gradient boosting methods, particularly XGBoost, achieve excellent results in structured data problems including load forecasting [1]. Chen and Guestrin introduced XGBoost as a scalable tree boosting system that has become the algorithm of choice for many machine learning competitions and practical applications [1].

Ke et al. [3] introduced LightGBM, a highly efficient gradient boosting decision tree framework using histogram-based algorithms for faster training while maintaining accuracy. Kong et al. [8] demonstrated the effectiveness of LSTM recurrent neural networks for short-term residential load forecasting, achieving significant improvements over conventional methods. Deep learning approaches have been explored by Shi et al. [12], who proposed a novel pooling deep RNN for household load forecasting.

*B. Regional Load Forecasting in GCC*

Swaroop [6] applied multilayer perceptron networks for load forecasting in Oman's Al Batinah region, demonstrating the need for localized models to address the country's rapid demand growth. Al-Hamadi and Soliman [7] utilized Kalman filtering for short-term forecasting in the GCC, noting the significant impact of temperature and cultural factors on consumption patterns. These studies highlight the importance of incorporating regional characteristics, including extreme weather conditions and religious observances, into forecasting models for improved accuracy in Middle Eastern distribution systems.

*C. Feature Engineering for Load Forecasting*

The importance of feature engineering in load forecasting has been extensively documented. Haben et al. [5] reviewed low voltage load forecasting methods, emphasizing the significance of incorporating calendar variables, weather data, and customer behavior patterns. Studies by Wang et al. [13] demonstrated that lag features and rolling statistics significantly improve forecast accuracy through their bi-directional LSTM method with attention mechanism. Chen et al. [2] utilized SVR models with carefully engineered features for baseline demand calculation in commercial buildings.

**III. METHODOLOGY**

*A. Data Collection and Description*

The dataset comprises 4,193 hourly observations from March 8, 2025 to August 31, 2025, collected via SCADA systems from the Khaboorah 33/11kV distribution substation in the North Batinah governorate of Oman. Eight feeder lines (KHBR01_K_LN01 through KHBR01_K_LN08) were monitored, with phase current measurements recorded in amperes. The substation serves a mixed residential and commercial area with approximately 5,000 connected customers.

*B. Feature Engineering*

A comprehensive feature engineering approach was implemented to capture temporal patterns at multiple scales:

• Temporal Features: Hour of day (0-23), day of week (0-6), day of month, week of year, and month. Sine and cosine encodings were applied to capture cyclical patterns: hour_sin = sin(2π × hour/24), hour_cos = cos(2π × hour/24).

• Lag Features: Historical load values at 1, 2, 3, 6, 12, 24, 48, and 168 hours prior to the forecast time, capturing short-term dynamics and weekly periodicity.

• Rolling Statistics: Moving window mean and standard deviation computed over 6, 12, 24, and 48 hour windows to capture trend and volatility information.

• Weather Features: Temperature (°C) and relative humidity (%) obtained from the nearest meteorological station. A binary indicator identifies the Ramadan period (March 1-30, 2025).

*C. Machine Learning Models*

Four machine learning algorithms were systematically evaluated for this study: XGBoost [1] - an optimized gradient boosting framework with L1 and L2 regularization to prevent overfitting; LightGBM [3] - a histogram-based gradient boosting framework with leaf-wise tree growth strategy; Random Forest - an ensemble of decision trees using bagging and feature randomization; Ridge Regression - a regularized linear model serving as a baseline for comparison.

*D. Model Evaluation Metrics*

Model performance was evaluated using four complementary metrics: Mean Absolute Error (MAE), Root Mean Square Error (RMSE), Coefficient of Determination (R²), and Mean Absolute Percentage Error (MAPE). The dataset was split 80/20 chronologically to preserve temporal ordering and prevent data leakage.

**IV. RESULTS**

*A. Data Exploration*

Initial data exploration revealed important characteristics of the load profile at the Khaboorah substation. Fig. 1 presents a comprehensive analysis of the load data showing the time series pattern, distribution histogram, hourly patterns, and daily patterns by day of week.

Fig. 1. Data exploration analysis showing (a) time series of load readings, (b) distribution histogram, (c) average hourly pattern, and (d) average daily pattern by day of week.

The load distribution exhibits a bimodal pattern with peaks at approximately 70-80A during Ramadan and 160-170A during non-Ramadan periods. Clear diurnal patterns emerge with peak loads occurring at 14:00 hours (2 PM) when air conditioning demand reaches maximum due to afternoon heat. Temperature ranged from 25°C in March to 44°C in August, showing strong correlation with load demand.

*B. Model Performance Comparison*

Table I presents the comprehensive performance comparison of all four models across training and testing datasets. XGBoost achieved the best overall performance with test R² of 0.8094 and RMSE of 6.827A. The model comparison visualization in Fig. 2 provides a graphical representation of these results.

**TABLE I: MODEL PERFORMANCE COMPARISON (MAE AND RMSE IN AMPERES)**

| Model | Train MAE | Test MAE | Train RMSE | Test RMSE |
| --- | --- | --- | --- | --- |
| XGBoost | 1.600 | 4.923 | 2.156 | 6.827 |
| LightGBM | 3.048 | 5.061 | 4.232 | 7.015 |
| Random Forest | 3.910 | 6.149 | 6.087 | 8.565 |
| Ridge Regression | 10.598 | 7.153 | 14.476 | 8.978 |

**TABLE II: R² SCORES AND MAPE COMPARISON**

| Model | Train R² | Test R² | Test MAPE |
| --- | --- | --- | --- |
| XGBoost | 0.9965 | 0.8094 | 3.25% |
| LightGBM | 0.9867 | 0.7988 | 3.35% |
| Random Forest | 0.9725 | 0.7000 | 4.04% |
| Ridge Regression | 0.8443 | 0.6704 | 4.59% |

Fig. 2. Model performance comparison showing (a) R² scores across models, (b) MAE and RMSE comparison, (c) predictions vs actual values, and (d) Test RMSE ranking.

LightGBM demonstrated comparable performance with test R² of 0.7988, falling only 1.3% behind XGBoost. Random Forest showed moderate performance with R² of 0.7000, while Ridge Regression served as a linear baseline with R² of 0.6704. The performance gap between XGBoost and Ridge Regression (20.7% improvement in R²) demonstrates the value of nonlinear models for this application.

*C. XGBoost Model Evaluation*

Fig. 3 presents detailed evaluation of the XGBoost model including actual vs predicted comparisons, scatter plots, residual analysis, and feature importance. The model demonstrates excellent fit with minimal systematic bias, as evidenced by the symmetric residual distribution centered near zero.

Fig. 3. XGBoost model evaluation showing (a) actual vs predicted comparison, (b) scatter plot of predictions, (c) residuals distribution, and (d) top 15 feature importance.

*D. Feature Importance Analysis*

Table III presents the top 10 features ranked by importance from the XGBoost model. Rolling statistics (6-hour rolling mean) emerged as the most influential predictor with 28.4% importance, capturing short-term load trends. The 24-hour lag feature ranked second at 19.8%, reflecting strong daily periodicity in consumption patterns.

**TABLE III: TOP 10 FEATURE IMPORTANCE RANKINGS (XGBOOST)**

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

Fig. 4. Feature importance comparison across XGBoost, Random Forest, and LightGBM models.

Fig. 4 compares feature importance rankings across three tree-based models. While the specific rankings differ slightly, all models consistently identify rolling statistics and lag features as dominant predictors. The 168-hour lag (one week) captures weekly periodicity, while temperature contributes to modeling seasonal and diurnal variations.

*E. Weather and Ramadan Impact Analysis*

Fig. 5 presents comprehensive analysis of weather features and their relationship with load patterns. The analysis reveals strong temperature-load correlation during summer months, with loads increasing approximately 3A per degree Celsius above 35°C. The Ramadan period analysis shows a dramatic 53% reduction in median load (approximately 75A versus 160A during non-Ramadan periods).

Fig. 5. Weather features analysis showing (a) temperature distribution by season, (b) temperature vs humidity correlation, (c) diurnal temperature pattern, and (d) load comparison during Ramadan vs non-Ramadan periods.

**V. DISCUSSION**

*A. Model Performance Analysis*

The superior performance of XGBoost can be attributed to several factors: (1) Effective capture of nonlinear relationships between features and load demand through gradient boosting; (2) Built-in regularization mechanisms (L1 and L2) that prevent overfitting despite high feature dimensionality; (3) Efficient handling of mixed feature types including categorical, continuous, and cyclical variables; (4) Ability to model feature interactions without explicit specification.

*B. Feature Engineering Insights*

The dominance of rolling statistics in feature importance rankings validates the hypothesis that recent load trends provide strong predictive signals for short-term forecasting. The 6-hour rolling mean captures intra-day patterns including morning ramp-up, afternoon peak, and evening decline. The significance of the 24-hour and 168-hour lags confirms the importance of daily and weekly periodicity in residential and commercial consumption patterns.

*C. Regional Considerations*

The 53% load reduction during Ramadan represents a critical finding for regional utilities. This substantial decrease reflects changes in daily routines, commercial operating hours, and residential activities during the holy month. The finding aligns with observations from other GCC countries and underscores the necessity of incorporating cultural calendar features in forecasting models for this region.

*D. Practical Applications*

The developed model can be integrated with existing SCADA infrastructure for real-time load prediction, enabling:

1) Proactive transformer loading management to prevent equipment overloading during peak hours (identified as 2 PM).

2) Optimized maintenance scheduling during predicted low-load periods.

3) Enhanced demand response planning during Ramadan when loads decrease by 53%.

4) Improved voltage regulation through anticipated demand patterns.

**VI. CONCLUSION**

This paper presented a comprehensive machine learning approach for short-term load forecasting at the 11kV Khaboorah distribution feeder in Oman. The study evaluated four machine learning models using real operational data from March to August 2025. XGBoost achieved the best overall performance with an R² score of 0.8094, RMSE of 6.827A, and MAPE of 3.25%, outperforming other models by 1.3-20.7% in R² score.

The feature engineering approach incorporating rolling statistics, lag features, and temporal encodings proved highly effective, with the 6-hour rolling mean and 24-hour lag emerging as dominant predictors. The analysis of Ramadan period impact revealed a significant 53% load reduction, providing valuable insights for demand-side management in regions with similar cultural patterns.

Future work will focus on extending the model to multiple substations, incorporating additional weather parameters, and developing ensemble approaches combining the strengths of XGBoost and deep learning methods for improved accuracy.

**ACKNOWLEDGMENT**

The author thanks Mazoon Electricity Company (MZEC) for providing SCADA data access and technical support. Support from the University of Technology and Applied Sciences (UTAS), Sohar campus, is gratefully acknowledged.

**REFERENCES**

[1] T. Chen and C. Guestrin, "XGBoost: A scalable tree boosting system," in Proc. 22nd ACM SIGKDD Int. Conf. Knowl. Discovery Data Mining, 2016, pp. 785-794.

[2] Y. Chen, P. Xu, Y. Chu, W. Li, Y. Wu, L. Ni, Y. Bao, and K. Wang, "Short-term electrical load forecasting using the Support Vector Regression (SVR) model to calculate the demand response baseline for office buildings," Applied Energy, vol. 195, pp. 659-670, 2017.

[3] G. Ke et al., "LightGBM: A highly efficient gradient boosting decision tree," in Advances in Neural Information Processing Systems, vol. 30, 2017.

[4] T. Hong and S. Fan, "Probabilistic electric load forecasting: A tutorial review," Int. J. Forecasting, vol. 32, no. 3, pp. 914-938, 2016.

[5] R. Haben, C. Singleton, and P. Grindrod, "Review of low voltage load forecasting: Methods, applications, and recommendations," Applied Energy, vol. 304, p. 117798, 2021.

[6] R. Swaroop, "Short-term load forecasting using artificial neural network for Al Batinah region in Oman," J. Eng. Sci. Tech., vol. 7, no. 4, pp. 498-504, 2012.

[7] H. M. Al-Hamadi and S. A. Soliman, "Short-term electric load forecasting based on Kalman filtering algorithm with moving window weather and load model," Electric Power Systems Research, vol. 68, no. 1, pp. 47-59, 2004.

[8] W. Kong, Z. Y. Dong, Y. Jia, D. J. Hill, Y. Xu, and Y. Zhang, "Short-term residential load forecasting based on LSTM recurrent neural network," IEEE Trans. Smart Grid, vol. 10, no. 1, pp. 841-851, 2019.

[9] S. Haben, S. Arber, V. Chiodo, P. Sherlock, and P. Sherlock, "Review and integration of data-driven load forecasting using machine learning and deep learning," Energy Reports, vol. 7, pp. 5234-5252, 2021.

[10] M. Q. Raza and A. Khosravi, "A review on artificial intelligence based load demand forecasting techniques for smart grid and buildings," Renewable and Sustainable Energy Reviews, vol. 50, pp. 1352-1372, 2015.

[11] A. Ahmad, N. Javaid, M. Guizani, N. Alrajeh, and Z. A. Khan, "An accurate and fast converging short-term load forecasting model for industrial applications in a smart grid," IEEE Trans. Ind. Informatics, vol. 13, no. 5, pp. 2587-2596, 2017.

[12] H. Shi, M. Xu, and R. Li, "Deep learning for household load forecasting—A novel pooling deep RNN," IEEE Trans. Smart Grid, vol. 9, no. 5, pp. 5271-5280, 2018.

[13] S. Wang, X. Wang, S. Wang, and D. Wang, "Bi-directional long short-term memory method based on attention mechanism and rolling update for short-term load forecasting," Int. J. Electrical Power Energy Syst., vol. 109, pp. 470-479, 2019.

[14] E. Zivot and J. Wang, "Rolling analysis of time series," in Modeling Financial Time Series with S-PLUS, New York: Springer, 2006, pp. 313-360.

[15] Authority for Electricity Regulation Oman, "Annual Report 2024," Muscat, Oman, 2024.

[16] X. Zhang, J. Wang, and K. Zhang, "Short-term electric load forecasting based on singular spectrum analysis and support vector machine optimized by cuckoo search algorithm," Electric Power Systems Research, vol. 146, pp. 270-285, 2017.

[17] L. Hernandez, C. Baladrón, J. M. Aguiar, B. Carro, A. J. Sánchez-Esguevillas, and J. Lloret, "Artificial neural networks for short-term load forecasting in microgrids environment," Energy, vol. 75, pp. 252-264, 2014.

[18] OSTI, "A cross-dimensional analysis of data-driven short-term load forecasting," U.S. Dept. Energy, Tech. Rep. 2997924, 2025.
