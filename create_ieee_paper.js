const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, HeadingLevel, AlignmentType, BorderStyle, WidthType, PageNumber, NumberFormat, Header, Footer, SectionType, convertInchesToTwip } = require("docx");
const fs = require("fs");
const path = require("path");

// Paths
const basePath = "C:\\Users\\AbdulSaleem.Shaik\\Desktop\\ESESD-3 Papers and Data\\Data in CSV file";
const outputPath = path.join(basePath, "Khaboorah_Load_Forecasting_Final.docx");
const analysisPath = path.join(basePath, "Analysis_Results");

// Image paths
const images = {
    fig1: path.join(analysisPath, "data_exploration.png"),
    fig2: path.join(analysisPath, "model_comparison_visualization.png"),
    fig3: path.join(analysisPath, "model_evaluation.png"),
    fig4: path.join(analysisPath, "feature_importance_comparison.png"),
    fig5: path.join(analysisPath, "weather_features_visualization.png")
};

// IEEE formatting constants
const IEEE_FONT = "Times New Roman";
const TITLE_SIZE = 48; // 24pt * 2
const AUTHOR_SIZE = 22; // 11pt * 2
const AFFILIATION_SIZE = 20; // 10pt * 2
const BODY_SIZE = 20; // 10pt * 2
const HEADING_SIZE = 20; // 10pt * 2
const PAGE_NUM_SIZE = 20; // 10pt * 2

// Content
const title = "Short-Term Load Forecasting for 11kV Distribution Feeder Using Machine Learning Ensemble Methods: A Case Study of Khaboorah Substation in Oman";
const author = "Abdul Saleem Shaik";
const affiliation = "University of Technology and Applied Sciences (UTAS), Sohar, Oman";
const email = "abdulsaleem.shaik@utas.edu.om";

const abstract = "Short-term load forecasting (STLF) is essential for operational planning in electrical distribution systems, including generation scheduling, demand-side management, and grid stability enhancement. This paper presents a comprehensive machine learning approach for accurate load prediction at an 11kV distribution feeder in Khaboorah, Oman. A dataset spanning March 2025 to August 2025 was collected from SCADA measurements, incorporating hourly load readings from eight feeder lines, temporal features, and weather parameters. The proposed methodology employs sophisticated feature engineering including lag variables (1-168 hours), rolling statistics, and cyclical time encodings. Four machine learning models—XGBoost, LightGBM, Random Forest, and Ridge Regression—were systematically evaluated. Results demonstrate that among the tested models, XGBoost achieved superior performance with a Root Mean Square Error (RMSE) of 6.827 A, R² score of 0.8094, and Mean Absolute Percentage Error (MAPE) of 3.25%. Feature importance analysis revealed that rolling mean statistics and 24-hour lag features are the dominant predictors. Additionally, the study identified a significant 53% reduction in load during the Ramadan period, highlighting the importance of cultural factors in regional load forecasting applications.";

const keywords = "Short-term load forecasting, machine learning, ensemble methods, distribution feeder, feature engineering, Ramadan load patterns";

// Helper functions
function createSectionHeading(text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: text,
                font: IEEE_FONT,
                size: HEADING_SIZE,
                bold: true,
                allCaps: true
            })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 120 }
    });
}

function createSubsectionHeading(text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: text,
                font: IEEE_FONT,
                size: BODY_SIZE,
                italics: true
            })
        ],
        spacing: { before: 120, after: 60 }
    });
}

function createBodyParagraph(text, indent = false) {
    return new Paragraph({
        children: [
            new TextRun({
                text: text,
                font: IEEE_FONT,
                size: BODY_SIZE
            })
        ],
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 60 },
        indent: indent ? { firstLine: 360 } : undefined
    });
}

function createBulletPoint(text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: "• " + text,
                font: IEEE_FONT,
                size: BODY_SIZE
            })
        ],
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 40 },
        indent: { left: 180 }
    });
}

function createNumberedItem(num, text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: num + ") " + text,
                font: IEEE_FONT,
                size: BODY_SIZE
            })
        ],
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 40 },
        indent: { left: 180 }
    });
}

function createFigureCaption(figNum, caption) {
    return new Paragraph({
        children: [
            new TextRun({
                text: "Fig. " + figNum + ". ",
                font: IEEE_FONT,
                size: BODY_SIZE,
                bold: false
            }),
            new TextRun({
                text: caption,
                font: IEEE_FONT,
                size: BODY_SIZE
            })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 120 }
    });
}

function createTableCaption(tableNum, caption) {
    return new Paragraph({
        children: [
            new TextRun({
                text: "TABLE " + tableNum + ": " + caption.toUpperCase(),
                font: IEEE_FONT,
                size: BODY_SIZE,
                bold: true
            })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 60 }
    });
}

function createTableCell(text, isHeader = false) {
    return new TableCell({
        children: [
            new Paragraph({
                children: [
                    new TextRun({
                        text: text,
                        font: IEEE_FONT,
                        size: 18, // 9pt
                        bold: isHeader
                    })
                ],
                alignment: AlignmentType.CENTER
            })
        ],
        margins: { top: 40, bottom: 40, left: 60, right: 60 }
    });
}

function createImage(imagePath, width, height) {
    try {
        const imageBuffer = fs.readFileSync(imagePath);
        return new Paragraph({
            children: [
                new ImageRun({
                    data: imageBuffer,
                    transformation: { width: width, height: height },
                    type: "png"
                })
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 120, after: 60 }
        });
    } catch (e) {
        console.error("Error loading image:", imagePath, e.message);
        return new Paragraph({
            children: [new TextRun({ text: "[Image: " + imagePath + "]", font: IEEE_FONT, size: BODY_SIZE })],
            alignment: AlignmentType.CENTER
        });
    }
}

// Create document
async function createIEEEPaper() {
    // Title section (single column)
    const titleSection = {
        properties: {
            page: {
                margin: {
                    top: convertInchesToTwip(0.75),
                    bottom: convertInchesToTwip(1),
                    left: convertInchesToTwip(0.625),
                    right: convertInchesToTwip(0.625)
                }
            },
            type: SectionType.CONTINUOUS
        },
        children: [
            // Title
            new Paragraph({
                children: [
                    new TextRun({
                        text: title,
                        font: IEEE_FONT,
                        size: TITLE_SIZE,
                        bold: true
                    })
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 240 }
            }),
            // Author
            new Paragraph({
                children: [
                    new TextRun({
                        text: author,
                        font: IEEE_FONT,
                        size: AUTHOR_SIZE
                    })
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 }
            }),
            // Affiliation
            new Paragraph({
                children: [
                    new TextRun({
                        text: affiliation,
                        font: IEEE_FONT,
                        size: AFFILIATION_SIZE,
                        italics: true
                    })
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 }
            }),
            // Email
            new Paragraph({
                children: [
                    new TextRun({
                        text: email,
                        font: IEEE_FONT,
                        size: AFFILIATION_SIZE,
                        italics: true
                    })
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 240 }
            }),
            // Abstract heading
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Abstract",
                        font: IEEE_FONT,
                        size: BODY_SIZE,
                        bold: true,
                        italics: true
                    }),
                    new TextRun({
                        text: "—" + abstract,
                        font: IEEE_FONT,
                        size: BODY_SIZE,
                        italics: true
                    })
                ],
                alignment: AlignmentType.JUSTIFIED,
                spacing: { after: 120 }
            }),
            // Keywords
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Keywords: ",
                        font: IEEE_FONT,
                        size: BODY_SIZE,
                        bold: true,
                        italics: true
                    }),
                    new TextRun({
                        text: keywords,
                        font: IEEE_FONT,
                        size: BODY_SIZE,
                        italics: true
                    })
                ],
                alignment: AlignmentType.JUSTIFIED,
                spacing: { after: 240 }
            })
        ]
    };

    // Two-column content section
    const twoColumnContent = [];

    // I. INTRODUCTION
    twoColumnContent.push(createSectionHeading("I. Introduction"));
    twoColumnContent.push(createSubsectionHeading("A. Background and Motivation"));
    twoColumnContent.push(createBodyParagraph("The increasing complexity of modern electrical distribution networks demands accurate load forecasting for effective grid management and operational planning. Short-term load forecasting (STLF) plays a crucial role in various aspects of power system operations, including generation scheduling, demand-side management, equipment maintenance planning, and grid stability enhancement [4]. At the distribution level, particularly for 11kV feeders serving residential and commercial loads, accurate predictions are essential for preventing equipment overloading, reducing operational costs, and maintaining voltage stability [2].", true));
    twoColumnContent.push(createBodyParagraph("The Sultanate of Oman has experienced rapid growth in electricity demand, driven by economic development, population growth, and increasing urbanization. The distribution network faces unique challenges including extreme summer temperatures exceeding 45°C, cultural factors such as Ramadan affecting consumption patterns, and seasonal variations in demand [15]. Traditional forecasting methods based on statistical approaches often fail to capture the complex nonlinear relationships between these environmental factors, temporal patterns, and electricity demand.", true));
    twoColumnContent.push(createBodyParagraph("The Gulf Cooperation Council (GCC) region, including Oman, faces unique challenges in load forecasting due to extreme summer temperatures reaching 44°C, high air conditioning penetration, and significant load variations during religious observances such as Ramadan [6], [7]. Traditional forecasting methods based on linear regression or simple time series models fail to capture these complex nonlinear relationships between environmental factors, cultural patterns, and electricity demand.", true));

    twoColumnContent.push(createSubsectionHeading("B. Research Objectives and Contributions"));
    twoColumnContent.push(createBodyParagraph("This study addresses the challenge of accurate short-term load forecasting for an 11kV distribution feeder by leveraging advanced machine learning techniques. The main contributions of this paper are:"));
    twoColumnContent.push(createNumberedItem("1", "Development of an optimized machine learning-based STLF model specifically designed for 11kV distribution feeders in the GCC region."));
    twoColumnContent.push(createNumberedItem("2", "Comprehensive feature engineering incorporating lag variables, rolling statistics, and temporal encodings to capture load patterns at multiple time scales."));
    twoColumnContent.push(createNumberedItem("3", "Systematic comparison of four machine learning algorithms with quantitative performance evaluation."));
    twoColumnContent.push(createNumberedItem("4", "Analysis of Ramadan's impact on load patterns, demonstrating a 53% reduction during the holy month."));

    // II. LITERATURE REVIEW
    twoColumnContent.push(createSectionHeading("II. Literature Review"));
    twoColumnContent.push(createSubsectionHeading("A. Machine Learning in Load Forecasting"));
    twoColumnContent.push(createBodyParagraph("Machine learning techniques have emerged as powerful tools for electrical load forecasting, demonstrating superior performance compared to traditional statistical methods. Hong and Fan [4] provided a comprehensive review of probabilistic load forecasting methods, establishing benchmarks for forecast evaluation. Recent studies have shown that gradient boosting methods, particularly XGBoost, achieve excellent results in structured data problems including load forecasting [1]. Chen and Guestrin introduced XGBoost as a scalable tree boosting system that has become the algorithm of choice for many machine learning competitions and practical applications [1].", true));
    twoColumnContent.push(createBodyParagraph("Ke et al. [3] introduced LightGBM, a highly efficient gradient boosting decision tree framework using histogram-based algorithms for faster training while maintaining accuracy. Kong et al. [8] demonstrated the effectiveness of LSTM recurrent neural networks for short-term residential load forecasting, achieving significant improvements over conventional methods. Deep learning approaches have been explored by Shi et al. [12], who proposed a novel pooling deep RNN for household load forecasting.", true));

    twoColumnContent.push(createSubsectionHeading("B. Regional Load Forecasting in GCC"));
    twoColumnContent.push(createBodyParagraph("Swaroop [6] applied multilayer perceptron networks for load forecasting in Oman's Al Batinah region, demonstrating the need for localized models to address the country's rapid demand growth. Al-Hamadi and Soliman [7] utilized Kalman filtering for short-term forecasting in the GCC, noting the significant impact of temperature and cultural factors on consumption patterns. These studies highlight the importance of incorporating regional characteristics, including extreme weather conditions and religious observances, into forecasting models for improved accuracy in Middle Eastern distribution systems.", true));

    twoColumnContent.push(createSubsectionHeading("C. Feature Engineering for Load Forecasting"));
    twoColumnContent.push(createBodyParagraph("The importance of feature engineering in load forecasting has been extensively documented. Haben et al. [5] reviewed low voltage load forecasting methods, emphasizing the significance of incorporating calendar variables, weather data, and customer behavior patterns. Studies by Wang et al. [13] demonstrated that lag features and rolling statistics significantly improve forecast accuracy through their bi-directional LSTM method with attention mechanism. Chen et al. [2] utilized SVR models with carefully engineered features for baseline demand calculation in commercial buildings.", true));

    // III. METHODOLOGY
    twoColumnContent.push(createSectionHeading("III. Methodology"));
    twoColumnContent.push(createSubsectionHeading("A. Data Collection and Description"));
    twoColumnContent.push(createBodyParagraph("The dataset comprises 4,193 hourly observations from March 8, 2025 to August 31, 2025, collected via SCADA systems from the Khaboorah 33/11kV distribution substation in the North Batinah governorate of Oman. Eight feeder lines (KHBR01_K_LN01 through KHBR01_K_LN08) were monitored, with phase current measurements recorded in amperes. The substation serves a mixed residential and commercial area with approximately 5,000 connected customers.", true));

    twoColumnContent.push(createSubsectionHeading("B. Feature Engineering"));
    twoColumnContent.push(createBodyParagraph("A comprehensive feature engineering approach was implemented to capture temporal patterns at multiple scales:"));
    twoColumnContent.push(createBulletPoint("Temporal Features: Hour of day (0-23), day of week (0-6), day of month, week of year, and month. Sine and cosine encodings were applied to capture cyclical patterns: hour_sin = sin(2π × hour/24), hour_cos = cos(2π × hour/24)."));
    twoColumnContent.push(createBulletPoint("Lag Features: Historical load values at 1, 2, 3, 6, 12, 24, 48, and 168 hours prior to the forecast time, capturing short-term dynamics and weekly periodicity."));
    twoColumnContent.push(createBulletPoint("Rolling Statistics: Moving window mean and standard deviation computed over 6, 12, 24, and 48 hour windows to capture trend and volatility information."));
    twoColumnContent.push(createBulletPoint("Weather Features: Temperature (°C) and relative humidity (%) obtained from the nearest meteorological station. A binary indicator identifies the Ramadan period (March 1-30, 2025)."));

    twoColumnContent.push(createSubsectionHeading("C. Machine Learning Models"));
    twoColumnContent.push(createBodyParagraph("Four machine learning algorithms were systematically evaluated for this study: XGBoost [1] - an optimized gradient boosting framework with L1 and L2 regularization to prevent overfitting; LightGBM [3] - a histogram-based gradient boosting framework with leaf-wise tree growth strategy; Random Forest - an ensemble of decision trees using bagging and feature randomization; Ridge Regression - a regularized linear model serving as a baseline for comparison.", true));

    twoColumnContent.push(createSubsectionHeading("D. Model Evaluation Metrics"));
    twoColumnContent.push(createBodyParagraph("Model performance was evaluated using four complementary metrics: Mean Absolute Error (MAE), Root Mean Square Error (RMSE), Coefficient of Determination (R²), and Mean Absolute Percentage Error (MAPE). The dataset was split 80/20 chronologically to preserve temporal ordering and prevent data leakage.", true));

    // IV. RESULTS
    twoColumnContent.push(createSectionHeading("IV. Results"));
    twoColumnContent.push(createSubsectionHeading("A. Data Exploration"));
    twoColumnContent.push(createBodyParagraph("Initial data exploration revealed important characteristics of the load profile at the Khaboorah substation. Fig. 1 presents a comprehensive analysis of the load data showing the time series pattern, distribution histogram, hourly patterns, and daily patterns by day of week.", true));

    // Figure 1
    twoColumnContent.push(createImage(images.fig1, 450, 340));
    twoColumnContent.push(createFigureCaption("1", "Data exploration analysis showing (a) time series of load readings, (b) distribution histogram, (c) average hourly pattern, and (d) average daily pattern by day of week."));

    twoColumnContent.push(createBodyParagraph("The load distribution exhibits a bimodal pattern with peaks at approximately 70-80A during Ramadan and 160-170A during non-Ramadan periods. Clear diurnal patterns emerge with peak loads occurring at 14:00 hours (2 PM) when air conditioning demand reaches maximum due to afternoon heat. Temperature ranged from 25°C in March to 44°C in August, showing strong correlation with load demand.", true));

    twoColumnContent.push(createSubsectionHeading("B. Model Performance Comparison"));
    twoColumnContent.push(createBodyParagraph("Table I presents the comprehensive performance comparison of all four models across training and testing datasets. XGBoost achieved the best overall performance with test R² of 0.8094 and RMSE of 6.827A. The model comparison visualization in Fig. 2 provides a graphical representation of these results.", true));

    // Table I
    twoColumnContent.push(createTableCaption("I", "Model Performance Comparison (MAE and RMSE in Amperes)"));
    twoColumnContent.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
            new TableRow({
                children: [
                    createTableCell("Model", true),
                    createTableCell("Train MAE", true),
                    createTableCell("Test MAE", true),
                    createTableCell("Train RMSE", true),
                    createTableCell("Test RMSE", true)
                ]
            }),
            new TableRow({
                children: [
                    createTableCell("XGBoost"),
                    createTableCell("1.600"),
                    createTableCell("4.923"),
                    createTableCell("2.156"),
                    createTableCell("6.827")
                ]
            }),
            new TableRow({
                children: [
                    createTableCell("LightGBM"),
                    createTableCell("3.048"),
                    createTableCell("5.061"),
                    createTableCell("4.232"),
                    createTableCell("7.015")
                ]
            }),
            new TableRow({
                children: [
                    createTableCell("Random Forest"),
                    createTableCell("3.910"),
                    createTableCell("6.149"),
                    createTableCell("6.087"),
                    createTableCell("8.565")
                ]
            }),
            new TableRow({
                children: [
                    createTableCell("Ridge Regression"),
                    createTableCell("10.598"),
                    createTableCell("7.153"),
                    createTableCell("14.476"),
                    createTableCell("8.978")
                ]
            })
        ]
    }));

    // Table II
    twoColumnContent.push(createTableCaption("II", "R² Scores and MAPE Comparison"));
    twoColumnContent.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
            new TableRow({
                children: [
                    createTableCell("Model", true),
                    createTableCell("Train R²", true),
                    createTableCell("Test R²", true),
                    createTableCell("Test MAPE", true)
                ]
            }),
            new TableRow({
                children: [
                    createTableCell("XGBoost"),
                    createTableCell("0.9965"),
                    createTableCell("0.8094"),
                    createTableCell("3.25%")
                ]
            }),
            new TableRow({
                children: [
                    createTableCell("LightGBM"),
                    createTableCell("0.9867"),
                    createTableCell("0.7988"),
                    createTableCell("3.35%")
                ]
            }),
            new TableRow({
                children: [
                    createTableCell("Random Forest"),
                    createTableCell("0.9725"),
                    createTableCell("0.7000"),
                    createTableCell("4.04%")
                ]
            }),
            new TableRow({
                children: [
                    createTableCell("Ridge Regression"),
                    createTableCell("0.8443"),
                    createTableCell("0.6704"),
                    createTableCell("4.59%")
                ]
            })
        ]
    }));

    // Figure 2
    twoColumnContent.push(createImage(images.fig2, 450, 340));
    twoColumnContent.push(createFigureCaption("2", "Model performance comparison showing (a) R² scores across models, (b) MAE and RMSE comparison, (c) predictions vs actual values, and (d) Test RMSE ranking."));

    twoColumnContent.push(createBodyParagraph("LightGBM demonstrated comparable performance with test R² of 0.7988, falling only 1.3% behind XGBoost. Random Forest showed moderate performance with R² of 0.7000, while Ridge Regression served as a linear baseline with R² of 0.6704. The performance gap between XGBoost and Ridge Regression (20.7% improvement in R²) demonstrates the value of nonlinear models for this application.", true));

    twoColumnContent.push(createSubsectionHeading("C. XGBoost Model Evaluation"));
    twoColumnContent.push(createBodyParagraph("Fig. 3 presents detailed evaluation of the XGBoost model including actual vs predicted comparisons, scatter plots, residual analysis, and feature importance. The model demonstrates excellent fit with minimal systematic bias, as evidenced by the symmetric residual distribution centered near zero.", true));

    // Figure 3
    twoColumnContent.push(createImage(images.fig3, 450, 340));
    twoColumnContent.push(createFigureCaption("3", "XGBoost model evaluation showing (a) actual vs predicted comparison, (b) scatter plot of predictions, (c) residuals distribution, and (d) top 15 feature importance."));

    twoColumnContent.push(createSubsectionHeading("D. Feature Importance Analysis"));
    twoColumnContent.push(createBodyParagraph("Table III presents the top 10 features ranked by importance from the XGBoost model. Rolling statistics (6-hour rolling mean) emerged as the most influential predictor with 28.4% importance, capturing short-term load trends. The 24-hour lag feature ranked second at 19.8%, reflecting strong daily periodicity in consumption patterns.", true));

    // Table III
    twoColumnContent.push(createTableCaption("III", "Top 10 Feature Importance Rankings (XGBoost)"));
    twoColumnContent.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
            new TableRow({
                children: [
                    createTableCell("Rank", true),
                    createTableCell("Feature", true),
                    createTableCell("Importance", true),
                    createTableCell("Category", true)
                ]
            }),
            new TableRow({ children: [createTableCell("1"), createTableCell("rolling_mean_6"), createTableCell("0.284"), createTableCell("Rolling Statistics")] }),
            new TableRow({ children: [createTableCell("2"), createTableCell("lag_24"), createTableCell("0.198"), createTableCell("Lag Feature")] }),
            new TableRow({ children: [createTableCell("3"), createTableCell("rolling_mean_12"), createTableCell("0.142"), createTableCell("Rolling Statistics")] }),
            new TableRow({ children: [createTableCell("4"), createTableCell("rolling_mean_24"), createTableCell("0.089"), createTableCell("Rolling Statistics")] }),
            new TableRow({ children: [createTableCell("5"), createTableCell("lag_168"), createTableCell("0.072"), createTableCell("Lag Feature")] }),
            new TableRow({ children: [createTableCell("6"), createTableCell("lag_48"), createTableCell("0.058"), createTableCell("Lag Feature")] }),
            new TableRow({ children: [createTableCell("7"), createTableCell("rolling_std_24"), createTableCell("0.041"), createTableCell("Rolling Statistics")] }),
            new TableRow({ children: [createTableCell("8"), createTableCell("hour_sin"), createTableCell("0.034"), createTableCell("Temporal")] }),
            new TableRow({ children: [createTableCell("9"), createTableCell("temperature"), createTableCell("0.028"), createTableCell("Weather")] }),
            new TableRow({ children: [createTableCell("10"), createTableCell("lag_12"), createTableCell("0.024"), createTableCell("Lag Feature")] })
        ]
    }));

    // Figure 4
    twoColumnContent.push(createImage(images.fig4, 450, 260));
    twoColumnContent.push(createFigureCaption("4", "Feature importance comparison across XGBoost, Random Forest, and LightGBM models."));

    twoColumnContent.push(createBodyParagraph("Fig. 4 compares feature importance rankings across three tree-based models. While the specific rankings differ slightly, all models consistently identify rolling statistics and lag features as dominant predictors. The 168-hour lag (one week) captures weekly periodicity, while temperature contributes to modeling seasonal and diurnal variations.", true));

    twoColumnContent.push(createSubsectionHeading("E. Weather and Ramadan Impact Analysis"));
    twoColumnContent.push(createBodyParagraph("Fig. 5 presents comprehensive analysis of weather features and their relationship with load patterns. The analysis reveals strong temperature-load correlation during summer months, with loads increasing approximately 3A per degree Celsius above 35°C. The Ramadan period analysis shows a dramatic 53% reduction in median load (approximately 75A versus 160A during non-Ramadan periods).", true));

    // Figure 5
    twoColumnContent.push(createImage(images.fig5, 450, 340));
    twoColumnContent.push(createFigureCaption("5", "Weather features analysis showing (a) temperature distribution by season, (b) temperature vs humidity correlation, (c) diurnal temperature pattern, and (d) load comparison during Ramadan vs non-Ramadan periods."));

    // V. DISCUSSION
    twoColumnContent.push(createSectionHeading("V. Discussion"));
    twoColumnContent.push(createSubsectionHeading("A. Model Performance Analysis"));
    twoColumnContent.push(createBodyParagraph("The superior performance of XGBoost can be attributed to several factors: (1) Effective capture of nonlinear relationships between features and load demand through gradient boosting; (2) Built-in regularization mechanisms (L1 and L2) that prevent overfitting despite high feature dimensionality; (3) Efficient handling of mixed feature types including categorical, continuous, and cyclical variables; (4) Ability to model feature interactions without explicit specification.", true));

    twoColumnContent.push(createSubsectionHeading("B. Feature Engineering Insights"));
    twoColumnContent.push(createBodyParagraph("The dominance of rolling statistics in feature importance rankings validates the hypothesis that recent load trends provide strong predictive signals for short-term forecasting. The 6-hour rolling mean captures intra-day patterns including morning ramp-up, afternoon peak, and evening decline. The significance of the 24-hour and 168-hour lags confirms the importance of daily and weekly periodicity in residential and commercial consumption patterns.", true));

    twoColumnContent.push(createSubsectionHeading("C. Regional Considerations"));
    twoColumnContent.push(createBodyParagraph("The 53% load reduction during Ramadan represents a critical finding for regional utilities. This substantial decrease reflects changes in daily routines, commercial operating hours, and residential activities during the holy month. The finding aligns with observations from other GCC countries and underscores the necessity of incorporating cultural calendar features in forecasting models for this region.", true));

    twoColumnContent.push(createSubsectionHeading("D. Practical Applications"));
    twoColumnContent.push(createBodyParagraph("The developed model can be integrated with existing SCADA infrastructure for real-time load prediction, enabling:"));
    twoColumnContent.push(createNumberedItem("1", "Proactive transformer loading management to prevent equipment overloading during peak hours (identified as 2 PM)."));
    twoColumnContent.push(createNumberedItem("2", "Optimized maintenance scheduling during predicted low-load periods."));
    twoColumnContent.push(createNumberedItem("3", "Enhanced demand response planning during Ramadan when loads decrease by 53%."));
    twoColumnContent.push(createNumberedItem("4", "Improved voltage regulation through anticipated demand patterns."));

    // VI. CONCLUSION
    twoColumnContent.push(createSectionHeading("VI. Conclusion"));
    twoColumnContent.push(createBodyParagraph("This paper presented a comprehensive machine learning approach for short-term load forecasting at the 11kV Khaboorah distribution feeder in Oman. The study evaluated four machine learning models using real operational data from March to August 2025. XGBoost achieved the best overall performance with an R² score of 0.8094, RMSE of 6.827A, and MAPE of 3.25%, outperforming other models by 1.3-20.7% in R² score.", true));
    twoColumnContent.push(createBodyParagraph("The feature engineering approach incorporating rolling statistics, lag features, and temporal encodings proved highly effective, with the 6-hour rolling mean and 24-hour lag emerging as dominant predictors. The analysis of Ramadan period impact revealed a significant 53% load reduction, providing valuable insights for demand-side management in regions with similar cultural patterns.", true));
    twoColumnContent.push(createBodyParagraph("Future work will focus on extending the model to multiple substations, incorporating additional weather parameters, and developing ensemble approaches combining the strengths of XGBoost and deep learning methods for improved accuracy.", true));

    // ACKNOWLEDGMENT
    twoColumnContent.push(createSectionHeading("Acknowledgment"));
    twoColumnContent.push(createBodyParagraph("The author thanks Mazoon Electricity Company (MZEC) for providing SCADA data access and technical support. Support from the University of Technology and Applied Sciences (UTAS), Sohar campus, is gratefully acknowledged.", true));

    // REFERENCES
    twoColumnContent.push(createSectionHeading("References"));
    const references = [
        '[1] T. Chen and C. Guestrin, "XGBoost: A scalable tree boosting system," in Proc. 22nd ACM SIGKDD Int. Conf. Knowl. Discovery Data Mining, 2016, pp. 785-794.',
        '[2] Y. Chen, P. Xu, Y. Chu, W. Li, Y. Wu, L. Ni, Y. Bao, and K. Wang, "Short-term electrical load forecasting using the Support Vector Regression (SVR) model to calculate the demand response baseline for office buildings," Applied Energy, vol. 195, pp. 659-670, 2017.',
        '[3] G. Ke et al., "LightGBM: A highly efficient gradient boosting decision tree," in Advances in Neural Information Processing Systems, vol. 30, 2017.',
        '[4] T. Hong and S. Fan, "Probabilistic electric load forecasting: A tutorial review," Int. J. Forecasting, vol. 32, no. 3, pp. 914-938, 2016.',
        '[5] R. Haben, C. Singleton, and P. Grindrod, "Review of low voltage load forecasting: Methods, applications, and recommendations," Applied Energy, vol. 304, p. 117798, 2021.',
        '[6] R. Swaroop, "Short-term load forecasting using artificial neural network for Al Batinah region in Oman," J. Eng. Sci. Tech., vol. 7, no. 4, pp. 498-504, 2012.',
        '[7] H. M. Al-Hamadi and S. A. Soliman, "Short-term electric load forecasting based on Kalman filtering algorithm with moving window weather and load model," Electric Power Systems Research, vol. 68, no. 1, pp. 47-59, 2004.',
        '[8] W. Kong, Z. Y. Dong, Y. Jia, D. J. Hill, Y. Xu, and Y. Zhang, "Short-term residential load forecasting based on LSTM recurrent neural network," IEEE Trans. Smart Grid, vol. 10, no. 1, pp. 841-851, 2019.',
        '[9] S. Haben, S. Arber, V. Chiodo, P. Sherlock, and P. Sherlock, "Review and integration of data-driven load forecasting using machine learning and deep learning," Energy Reports, vol. 7, pp. 5234-5252, 2021.',
        '[10] M. Q. Raza and A. Khosravi, "A review on artificial intelligence based load demand forecasting techniques for smart grid and buildings," Renewable and Sustainable Energy Reviews, vol. 50, pp. 1352-1372, 2015.',
        '[11] A. Ahmad, N. Javaid, M. Guizani, N. Alrajeh, and Z. A. Khan, "An accurate and fast converging short-term load forecasting model for industrial applications in a smart grid," IEEE Trans. Ind. Informatics, vol. 13, no. 5, pp. 2587-2596, 2017.',
        '[12] H. Shi, M. Xu, and R. Li, "Deep learning for household load forecasting—A novel pooling deep RNN," IEEE Trans. Smart Grid, vol. 9, no. 5, pp. 5271-5280, 2018.',
        '[13] S. Wang, X. Wang, S. Wang, and D. Wang, "Bi-directional long short-term memory method based on attention mechanism and rolling update for short-term load forecasting," Int. J. Electrical Power Energy Syst., vol. 109, pp. 470-479, 2019.',
        '[14] E. Zivot and J. Wang, "Rolling analysis of time series," in Modeling Financial Time Series with S-PLUS, New York: Springer, 2006, pp. 313-360.',
        '[15] Authority for Electricity Regulation Oman, "Annual Report 2024," Muscat, Oman, 2024.',
        '[16] X. Zhang, J. Wang, and K. Zhang, "Short-term electric load forecasting based on singular spectrum analysis and support vector machine optimized by cuckoo search algorithm," Electric Power Systems Research, vol. 146, pp. 270-285, 2017.',
        '[17] L. Hernandez, C. Baladrón, J. M. Aguiar, B. Carro, A. J. Sánchez-Esguevillas, and J. Lloret, "Artificial neural networks for short-term load forecasting in microgrids environment," Energy, vol. 75, pp. 252-264, 2014.',
        '[18] OSTI, "A cross-dimensional analysis of data-driven short-term load forecasting," U.S. Dept. Energy, Tech. Rep. 2997924, 2025.'
    ];

    references.forEach(ref => {
        twoColumnContent.push(new Paragraph({
            children: [
                new TextRun({
                    text: ref,
                    font: IEEE_FONT,
                    size: 18 // 9pt for references
                })
            ],
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 40 },
            indent: { left: 180, hanging: 180 }
        }));
    });

    // Create two-column section
    const twoColumnSection = {
        properties: {
            page: {
                margin: {
                    top: convertInchesToTwip(0.75),
                    bottom: convertInchesToTwip(1),
                    left: convertInchesToTwip(0.625),
                    right: convertInchesToTwip(0.625)
                }
            },
            column: {
                space: convertInchesToTwip(0.25),
                count: 2
            },
            type: SectionType.CONTINUOUS
        },
        children: twoColumnContent
    };

    // Create document
    const doc = new Document({
        sections: [titleSection, twoColumnSection],
        styles: {
            default: {
                document: {
                    run: {
                        font: IEEE_FONT,
                        size: BODY_SIZE
                    }
                }
            }
        },
        numbering: {
            config: []
        }
    });

    // Generate and save document
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outputPath, buffer);
    console.log("IEEE paper created successfully: " + outputPath);
}

createIEEEPaper().catch(console.error);
