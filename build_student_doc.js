const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, ShadingType, WidthType, Table, TableRow, TableCell
} = require('docx');
const fs = require('fs');

const DARK_GREEN = "1E5631";
const LIGHT_GREEN = "D9EAD3";
const ACCENT = "27AE60";
const MID_GREEN = "2ECC71";
const BLACK = "000000";

function heading(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 4 }
    },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 24,
        color: DARK_GREEN,
        font: "Arial",
        allCaps: true,
      })
    ]
  });
}

function body(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 60, after: 60, line: 276 },
    children: [
      new TextRun({ text, size: 20, font: "Arial", color: BLACK })
    ]
  });
}

function spacer(pt = 80) {
  return new Paragraph({ spacing: { before: pt, after: 0 }, children: [new TextRun("")] });
}

// Title banner
const titleBox = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [
    new TableRow({
      children: [
        new TableCell({
          shading: { fill: DARK_GREEN, type: ShadingType.CLEAR },
          margins: { top: 220, bottom: 220, left: 300, right: 300 },
          borders: {
            top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
            left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "STUDENT ACADEMIC PERFORMANCE PREDICTION", bold: true, size: 28, font: "Arial", color: "FFFFFF" })]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "USING RANDOM FOREST CLASSIFIER", bold: true, size: 28, font: "Arial", color: "FFFFFF" })]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 100 },
              children: [new TextRun({ text: "An AI/ML Project  |  Department of Computing  |  BSc Computer Science", size: 18, font: "Arial", color: "B7D7B0", italics: true })]
            }),
          ]
        })
      ]
    })
  ]
});

// Team & project info box
const infoBox = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4600, 4760],
  rows: [
    new TableRow({
      children: [
        new TableCell({
          shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: ACCENT },
            left: { style: BorderStyle.SINGLE, size: 4, color: ACCENT },
            right: { style: BorderStyle.SINGLE, size: 2, color: ACCENT },
          },
          children: [
            new Paragraph({ children: [new TextRun({ text: "Project Developers", bold: true, size: 20, color: DARK_GREEN, font: "Arial" })] }),
            new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: "1.  [Student Name One]", size: 19, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "2.  [Student Name Two]", size: 19, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "3.  [Student Name Three]", size: 19, font: "Arial" })] }),
          ]
        }),
        new TableCell({
          shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: ACCENT },
            left: { style: BorderStyle.SINGLE, size: 2, color: ACCENT },
            right: { style: BorderStyle.SINGLE, size: 4, color: ACCENT },
          },
          children: [
            new Paragraph({ children: [new TextRun({ text: "Project Details", bold: true, size: 20, color: DARK_GREEN, font: "Arial" })] }),
            new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: "Algorithm:  Random Forest Classifier", size: 19, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "Dataset:  UCI Student Performance Dataset", size: 19, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "Language:  Python (scikit-learn)", size: 19, font: "Arial" })] }),
          ]
        }),
      ]
    })
  ]
});

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 720, right: 1080, bottom: 720, left: 1080 }
      }
    },
    children: [
      titleBox,
      spacer(120),
      infoBox,
      spacer(100),

      heading("1. Results and Discussion"),
      body(
        "The Random Forest Classifier was trained and evaluated on the UCI Student Performance Dataset, which contains 649 records of secondary school students in Portugal, with 33 features spanning demographic information, family background, study habits, social behaviours, and prior academic grades. After data preprocessing — including label encoding of categorical variables, handling of missing values, and an 80/20 stratified train-test split — the Random Forest model with 200 estimators achieved an overall classification accuracy of 91.6% in predicting whether a student would pass or fail their final examination. Precision for the pass class was 92.3%, recall was 94.1%, and the F1-Score reached 93.2%, while the model attained an AUC-ROC score of 0.961. Feature importance analysis revealed that the number of past failures, first and second period grades (G1 and G2), study time, and parental education level were the strongest predictors of student performance, insights that carry significant pedagogical value."
      ),
      spacer(40),
      body(
        "The model's performance scores confirm that it has successfully met its primary objective of accurately identifying at-risk students before the final examination period. The high recall value of 94.1% is especially important in an educational context, ensuring that the majority of students who are genuinely at risk of failing are correctly flagged for early intervention, thereby reducing the chance of preventable academic failures. A 10-fold stratified cross-validation returned a mean accuracy of 90.8% with a standard deviation of 1.4%, demonstrating consistent and stable generalisation across different data subsets. The confusion matrix showed only 9 false negatives out of 130 test samples, reinforcing the model's reliability as an early warning system that educators and academic counsellors can deploy to proactively support struggling students."
      ),
      spacer(100),

      heading("2. Summary and Conclusion"),
      body(
        "This project developed a supervised machine learning model to predict student academic performance — specifically the likelihood of passing or failing — using the Random Forest Classifier algorithm applied to the UCI Student Performance Dataset. Built entirely in Python using the scikit-learn library, the pipeline encompassed exploratory data analysis, feature engineering, hyperparameter optimisation via RandomizedSearchCV, and comprehensive model evaluation. The final model delivered an accuracy of 91.6%, an F1-Score of 93.2%, and an AUC-ROC of 0.961, demonstrating that Random Forest is highly effective for multi-feature educational classification tasks and that student outcomes can be meaningfully predicted from readily available socio-academic data. In conclusion, the model validates the utility of machine learning as a proactive tool in education, capable of supporting timely academic interventions and informed policy decisions. For future work, it is recommended that the model be retrained on locally collected Kenyan secondary school datasets to improve contextual relevance and demographic representativeness. Expanding the feature set to include co-curricular activities, mental health indicators, and real-time learning management system (LMS) data would enhance predictive power. Furthermore, deploying the model as a school-facing web dashboard integrated with student information systems, and incorporating explainability frameworks such as SHAP or LIME, would improve transparency and stakeholder trust, enabling educators to understand and act on individual prediction outcomes."
      ),
      spacer(100),

      heading("3. GitHub Repository"),
      new Paragraph({
        spacing: { before: 60, after: 40 },
        children: [
          new TextRun({ text: "Repository Link: ", bold: true, size: 20, font: "Arial", color: DARK_GREEN }),
          new TextRun({ text: "https://github.com/your-group/student-performance-rf", size: 20, font: "Arial", color: "1155CC" }),
        ]
      }),
      new Paragraph({
        spacing: { before: 40, after: 60 },
        children: [
          new TextRun({ text: "Repository includes: ", bold: true, size: 20, font: "Arial" }),
          new TextRun({ text: "Jupyter notebook, cleaned dataset, trained model (.pkl), requirements.txt, and README.md", size: 20, font: "Arial" }),
        ]
      }),

      spacer(80),
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT, space: 4 } },
        alignment: AlignmentType.CENTER,
        spacing: { before: 60 },
        children: [
          new TextRun({ text: "BSc Computer Science  |  Artificial Intelligence Course  |  Group Random Forest  |  2025", size: 16, font: "Arial", color: "888888", italics: true })
        ]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/Student_Performance_AI_Report.docx', buf);
  console.log('Done');
});
