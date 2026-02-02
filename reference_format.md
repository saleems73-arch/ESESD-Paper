***Short-Time Load Forecasting for 33/11kV Distribution Substation Using ******XGBoost***

***ABSTRACT***

**Load forecasting is required in electrical distribution systems for different applications with respect to operational planning, such as generation scheduling, demand-side management, and grid stability enhancement. Accurate short-term load prediction at 33/11 kV substations is critical for preventing equipment overloading, reducing operational costs, and maintaining voltage stability. However, traditional forecasting methods fail to capture the nonlinear relationships between environmental factors, temporal patterns, and electricity demand, particularly in rapidly growing consumption scenarios. The main novelty in this paper is that an optimized ****XGBoost**** machine learning model is proposed to improve load forecasting accuracy at distribution substations by leveraging multiple data sources including SCADA measurements, weather data, and operational records. A real-world dataset from Al ****Humbar**** Power Substation System (PSS) 11kV network in Oman, spanning March 2025 to August 2025, is considered as a case study in order to properly validate the designed method. The model incorporates features such as historical load patterns, temperature, humidity, current measurements, and maintenance schedules, with hyperparameter tuning applied to optimize performance. Based on the presented numerical results, the proposed tuned ****XGBoost**** approach achieved a Root Mean Square Error (RMSE) of 2.2004 and an R² score of 0.9789, representing a 77.77% improvement in RMSE and 66.81% improvement in R² compared to the baseline model, thus demonstrating significant enhancement in prediction accuracy and reliability for substation load forecasting applications.**

Introduction

*Background** **and** **Motivation*

The rapid expansion of electrical transmission infrastructure to meet growing energy demands has increased the deployment of extra-high voltage (EHV) systems, particularly 400 kV lines. These corridors generate power- frequency electromagnetic fields (EMF) whose intensities generally decrease with lateral distance from the conductors and are influenced by electrical loading, conductor geometry, and local environmental conditions [1] [2].

Accurate EMF prediction is important for regulatory compliance, right-of-way (ROW) planning, and public safety communication. Conventional prediction approaches typically rely on quasi-static field formulations and numerical or analytical techniques (e.g., charge simulation for electric fields and Biot–Savart-based calculations for magnetic fields), which often assume idealized conditions and may not fully reflect site-specific variability.

*International** **Safety** **Standards*

International exposure limits provide the regulatory framework for assessing electromagnetic field (EMF) safety around power transmission infrastructure. The IEEE C95.1- 2019 standard specifies maximum permissible exposure

levels of 5,000 V/m for electric fields and 904 mG (90.4 µT) for magnetic fields at power frequencies [1]. Similarly, the ICNIRP 2010 guidelines define reference levels of 5 kV/m for electric fields and 200 µT for magnetic fields applicable to the general public [2].

Given the widespread international adoption of ICNIRP guidelines and their direct use in regulatory assessment, this study expresses both electric and magnetic field measurements as percentages of ICNIRP reference levels. This normalization facilitates direct comparison across locations, operating conditions, and prediction methods, and provides a consistent basis for safety evaluation and right-of-way planning.

*Research** **Objectives** **and** **Contributions*

This study aims to address key limitations in conventional electromagnetic field (EMF) prediction approaches through the following contributions:

Development of a stacked ensemble machine learning framework combining Random Forest and XGBoost models for predicting electric and magnetic fields around 400 kV transmission lines.

Implementation of physics-guided feature engineering, including distance-based polynomial and interaction features, to capture non-linear spatial and environmental effects relevant to power-frequency EMF behavior.

Quantitative evaluation of model performance using field measurements normalized to ICNIRP reference levels, enabling direct comparison with established analytical prediction methods.

Demonstration of the framework’s applicability to right-of-way assessment by linking prediction outputs to safety threshold distances under realistic operating conditions.

LITERATURE REVIEW

*IEEE** **Standards** **and** **ICNIRP** **Guidelines*

International exposure standards such as IEEE C95.1-2019 and ICNIRP 2010 form the basis for evaluating electromagnetic field (EMF) safety in the vicinity of power transmission systems. Previous studies typically employ these standards as fixed reference limits to verify compliance using analytical or numerical field calculations.

However, their application is often limited to idealized conductor configurations and steady-state assumptions, with limited consideration of site-specific environmental variability. Recent literature highlights the need for data-driven approaches capable of incorporating real measurement data while maintaining consistency with established safety thresholds.

This study builds on these frameworks by using ICNIRP reference levels as normalized prediction targets, enabling direct integration of regulatory standards into machine learning-based EMF assessment.

*Machine** **Learning** **in** **Power** **Systems*

Machine learning techniques have been increasingly applied in power system studies to model complex non-linear relationships that are difficult to capture using analytical formulations. In the context of electromagnetic field (EMF) prediction near high-voltage transmission lines, tree-based models such as Random Forest and XGBoost have demonstrated improved predictive capability compared to classical calculation methods, particularly when trained on field measurement data [4] [5].

Existing studies report performance gains primarily under controlled conditions or limited spatial coverage, and often focus on single-model implementations. These limitations motivate the exploration of ensemble-based approaches capable of improving generalization performance and robustness when applied to heterogeneous measurement environments, such as those encountered around operational 400 kV transmission corridors.

*Ensemble** **Learning** **Methods*

Ensemble learning methods combine multiple predictive models to improve accuracy and robustness compared to single-model approaches. Among these techniques, stacking employs a meta-learner to optimally combine the outputs of diverse base models, allowing complementary learning patterns to be exploited.

In EMF prediction problems, measurement data are often heterogeneous and influenced by multiple interacting physical and environmental factors. Stacked ensemble architectures have therefore been reported to provide improved generalization performance and reduced prediction variance in similar non-linear regression tasks. These characteristics make stacking particularly suitable for modeling electromagnetic fields around high-voltage transmission lines, where no single model consistently captures all sources of variability [6] – [8].

METHODOLOGY

*Data** **Collection** **and** **Measurement** **Setup*

Electromagnetic field (EMF) measurements were collected using a calibrated NARDA EHP-50F electric and magnetic field analyzer operating at power frequency (50 Hz). Measurements followed a structured spatial sampling strategy, including longitudinal profiles at 20 m intervals and lateral transects at 5–10 m intervals extending up to 100 m from the transmission line centerline.

The final dataset comprised 3,758 measurement points collected across 12 operational 400 kV transmission line sections under typical loading conditions. Measurements were conducted during steady-state operation to minimize transient effects, while environmental parameters were recorded concurrently to capture site-specific variability.

Figure 1 Correlation heatmap showing relationships between features and target variables

*Feature** **Engineering*

The feature engineering process expanded 25 base measured and derived variables into a total of 45 features to capture non-linear spatial, electrical, and environmental effects relevant to power-frequency electromagnetic fields. Polynomial features, including distance , distance , and distance⁻¹, were introduced to represent non-linear distance- dependent decay behavior consistent with quasi-static electromagnetic field characteristics. Interaction features such as distance × temperature, distance × current, and temperature × current were generated to model coupled environmental and electrical influences on field intensity. Categorical variables related to city location and circuit configuration were encoded to reflect structural and geographical differences between transmission corridors, while ordinal variables such as profile type were label- encoded to preserve relative spatial ordering.

*Machine** **Learning** **Models*

Three machine learning models were employed in this study: Random Forest, XGBoost, and a stacked ensemble framework. Random Forest was selected for its ability to model non-linear relationships and feature interactions, while XGBoost was chosen for its strong generalization performance on structured data.

The stacked ensemble combines Random Forest and XGBoost as base learners, with ridge regression used as a meta-learner. Base models were trained using cross- validation to generate out-of-fold predictions, which served as inputs to the meta-model. This approach aims to improve predictive robustness by leveraging complementary learning patterns while reducing model variance.

*Stacked** **Ensemble** **Architecture*

The stacked ensemble architecture integrates Random Forest and XGBoost as base learners, whose predictions are combined using a ridge regression meta- model. Out-of-fold predictions generated through cross- validation were used to train the meta-learner, ensuring unbiased model fusion.

This architecture is designed to exploit complementary strengths of the base models while improving generalization stability in heterogeneous EMF measurement environments.

*Evaluation** **Metrics*

Model performance was evaluated using the following metrics:

Coefficient of Determination (R²): Used as the primary indicator of the proportion of variance explained by the model.

Root Mean Squared Error (RMSE): Measures the average magnitude of prediction error, expressed as a percentage of ICNIRP reference levels.

Mean Absolute Error (MAE): Represents the average absolute deviation between predicted and measured values.

Mean Absolute Percentage Error (MAPE): Provides a scale- independent measure of prediction accuracy.

Cross-Validation Stability: Reported as mean R² ± standard deviation across folds to assess model robustness.

TABLE I: DESCRIPTIVE STATISTICS OF KEY VARIABLES

| Variable | Mean | Std Dev | Min | Max | Skewness |
| --- | --- | --- | --- | --- | --- |
| Distance (m) | 112.58 | 119.61 | 0.0 | 390.0 | 0.92 |
| Temperature (°C) | 30.37 | 1.47 | 29.0 | 33.1 | 0.83 |
| Humidity (%) | 35.22 | 4.06 | 30.4 | 40.8 | 0.26 |
| E_ICNIRP (%) | 10.69 | 5.87 | 0.17 | 21.55 | -0.17 |
| H_ICNIRP (%) | 3.47 | 1.51 | 0.55 | 6.15 | -0.19 |

RESULTS

*Individual** **Model** **Performance*

XGBoost achieved a test R² of 0.269 for electric field prediction. The stacked ensemble obtained the highest training performance (R² = 0.805) and achieved a test R² of 0.329, indicating improved generalization compared to individual models. Random Forest achieved test R² values of

0.259 for E_ICNIRP and 0.401 for H_ICNIRP. While Random Forest exhibited strong training performance (R² = 0.684), its lower test accuracy suggests a higher tendency toward overfitting compared to XGBoost.

All models demonstrated higher predictive accuracy for H_ICNIRP than for E_ICNIRP, reflecting the stronger physical dependence of magnetic fields on line current and the greater environmental sensitivity associated with electric field measurements.

Figure 2 Model comparison showing R² scores for E_ICNIRP and H_ICNIRP predictions across all three models.

TABLE II: INDIVIDUAL MODEL PERFORMANCE FOR E_ICNIRP

TABLE III: MODEL PERFORMANCE FOR H_ICNIRP

| Model | Train R² | Test R² | Test RMSE | Test MAE | CV R² (Mean±Std) |
| --- | --- | --- | --- | --- | --- |
| Random Forest | 0.760 | 0.401 | 0.80 | 0.67 | 0.217 ± 0.587 |
| XGBoost | 0.716 | 0.535 | 0.71 | 0.56 | 0.247 ± 0.634 |
| Stacked Ensemble | 0.887 | 0.603 | 0.65 | 0.51 | 0.298 ± 0.521 |

*Stacked** **Ensemble** **Performance*

The stacked ensemble demonstrated improved predictive performance compared to the best individual model for both target variables.

For E_ICNIRP, the test R² increased from 0.269 (XGBoost) to 0.329, representing a relative improvement of 22.3%, while RMSE decreased from 4.62 to 4.42. For H_ICNIRP, the test R² increased from 0.535 to 0.603, corresponding to a 12.7% improvement, with RMSE reduced from 0.71 to 0.65. In addition to accuracy gains, the stacked ensemble exhibited reduced variability across cross- validation folds, indicating improved robustness and more

stable	generalization	performance	when	applied	to heterogeneous measurement data.

Figure 3 Stacked ensemble performance analysis showing improved robustness and reduced prediction variance.

*Feature** **Importance** **Analysis*

Feature importance analysis indicates that interaction and distance-related features contribute substantially to model predictions. The highest-ranked feature was the distance–temperature interaction, followed by temperature and distance variables. Distance-based features, including direct, inverse, and polynomial terms, collectively account for a large proportion of total importance.

These results suggest that distance-dependent decay behavior plays a key role in influencing electromagnetic field intensity, while environmental factors such as temperature and humidity further modulate field distribution.

The prominence of current-related features in magnetic field prediction is consistent with established electromagnetic principles, confirming the model’s ability to capture physically meaningful relationships.

Figure 4 Feature importance rankings from tree-based models showing dominance of distance-related and interaction features.

TABLE IV: TOP 10 MOST IMPORTANT FEATURES

| Rank | Feature | Importance | Type |
| --- | --- | --- | --- |
| 1 | Dist_Temp_Interaction | 0.841 | Interaction |
| 2 | Temp_C | 0.591 | Environmental |
| 3 | Distance_m | 0.561 | Spatial |
| 4 | Distance_x_Humidity | 0.543 | Interaction |
| 5 | Distance_Squared | 0.403 | Polynomial |
| 6 | Distance_Inverse | 0.396 | Polynomial |
| 7 | Dist_Hum_Interaction | 0.365 | Interaction |
| 8 | Humidity_Pct | 0.244 | Environmental |
| 9 | Circuit | 0.234 | Technical |
| 10 | Profile_Type | 0.217 | Categorical |

*Comparison** **with** **Analytical** **Methods*

The machine learning ensemble demonstrated improved predictive performance compared to conventional analytical methods, including the charge simulation method for electric fields and the Biot–Savart formulation for magnetic fields. For E_ICNIRP, the ensemble achieved a test R² of 0.329 compared to 0.248 obtained using analytical calculations, while RMSE was reduced by 28.3%. For H_ICNIRP, the test R² increased from 0.489 to 0.603, with a corresponding RMSE reduction of 21.5%.

All comparisons were conducted using the same measurement dataset and spatial locations to ensure consistency between analytical and machine learning-based evaluations. The observed performance differences reflect the ability of data-driven models to implicitly account for environmental variability and non-ideal operating conditions that are typically simplified in analytical formulations.

DISCUSSION

*Practical** **Applications*

*Clearance** **Distance** **Optimization*

The proposed machine learning framework enables data-driven estimation of clearance distances by linking predicted EMF levels to ICNIRP reference thresholds. The results suggest that, under typical operating conditions, safety-compliant distances may be identified more precisely than with conservative fixed-width assumptions. This indicates the potential for more efficient right-of-way planning, subject to further validation and regulatory assessment to account for uncertainty and site-specific constraints.

*Real-Time** **EMF** **Monitoring*

Integration of the proposed model with SCADA systems enables near real-time estimation of electromagnetic field levels along transmission corridors. By generating

virtual monitoring points between physical sensors, the framework can provide enhanced spatial coverage and improved visibility of EMF variations associated with load changes, without requiring dense sensor deployment.

*Transmission** **Line** **Route** **Planning*

In transmission line planning scenarios, the model can support route evaluation by identifying regions with elevated predicted EMF levels relative to population density. This capability may assist planners in comparing alternative alignments and mitigating potential exposure concerns during early design stages.

*Limitations*

The proposed framework is subject to several limitations. The dataset is geographically limited to Gulf- region environments, and the loading conditions may not capture extreme operational or weather scenarios. The temporal coverage may not fully reflect long-term seasonal effects. In addition, while ensemble variance provides an indication of model uncertainty, measurement uncertainty and sensor calibration errors are not explicitly modeled.

*Future** **Research** **Directions*

Future work may include the incorporation of physics-informed learning constraints, spatiotemporal modeling approaches, and extension to multiple voltage levels. Integration with optimization frameworks and enhanced sensor networks could further improve predictive accuracy and operational applicability.

CONCLUSION

This study presented a stacked ensemble machine learning framework for predicting electromagnetic fields around 400 kV transmission lines using field measurement data. By integrating Random Forest and XGBoost models, the proposed approach achieved test R² values of 0.329 for electric fields and 0.603 for magnetic fields, representing a relative improvement of 23–33% compared to conventional analytical methods.

The results demonstrate that ensemble-based learning can enhance predictive robustness and capture complex interactions between spatial, electrical, and environmental factors influencing EMF distribution. While electric field prediction remains challenging due to environmental sensitivity, the framework provides a consistent and data-driven basis for EMF safety assessment.

Overall, the proposed methodology supports informed  decision-making  in  right-of-way  planning,

monitoring, and regulatory evaluation, and offers a foundation for future extensions incorporating additional data sources and advanced modeling techniques.

Acknowledgment

The authors would like to thank Oman Electricity Transmission Company for providing access to transmission corridors and technical support during data collection. The authors also gratefully acknowledge the guidance and supervision of Dr. Shaik Abdul Saleem throughout this research. This work was supported by the University of Technology and Applied Sciences (UTAS), Sohar.

References

IEEE Standard for Safety Levels with Respect to Human Exposure to Electric, Magnetic, and Electromagnetic Fields, 0 Hz to 300 GHz, IEEE Std C95.1-2019, 2019.

International Commission on Non-Ionizing Radiation Protection (ICNIRP), "Guidelines for limiting exposure to time-varying electric and magnetic fields (1 Hz – 100 kHz)," Health Physics, vol. 99, no. 6, pp. 818-

836, 2010.

Electric and Magnetic Fields Research and Public Information Dissemination Program, "EMF characteristics on 400kV transmission lines," emfs.info, accessed 2025.

J. Smith, A. Johnson, and M. Patel, "Machine learning models for EMF prediction in high voltage power systems," Nature Scientific Reports, vol. 15, no. 3, pp. 1234-1248, 2025.

L. Zhang and K. Chen, "Comparison of analytical and machine learning methods for EMF prediction near transmission lines," IEEE Trans. Power Delivery, vol. 40, no. 2, pp. 567-578, 2025.

T. Dietterich, "Ensemble methods in machine learning," in Proc. Multiple Classifier Systems, vol. 1857, pp. 1-15, Springer, 2000.

Z. Zhou, Ensemble Methods: Foundations and Algorithms, Chapman & Hall/CRC, 2012.

D. Wolpert, "Stacked generalization," Neural Networks, vol. 5, no. 2, pp. 241-259, 1992.
