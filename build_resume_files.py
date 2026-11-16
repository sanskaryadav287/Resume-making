import os
import zipfile
import subprocess
import sys

print("Building Updated Sanskar Yadav Resume Assets...")

# -------------------------------------------------------------------------
# CREATE PROFESSIONAL ATS-FRIENDLY DOCX (With Updated Links & Data Science Skills)
# -------------------------------------------------------------------------

def build_docx(filename="Sanskar_Yadav_Resume.docx"):
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

    core_props = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
  xmlns:dc="http://purl.org/dc/elements/1.1/" 
  xmlns:dcterms="http://purl.org/dc/terms/" 
  xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Sanskar Yadav - Data Science Resume</dc:title>
  <dc:subject>Data Science, Machine Learning &amp; AWS Cloud Resume</dc:subject>
  <dc:creator>Sanskar Yadav</dc:creator>
  <cp:lastModifiedBy>Sanskar Yadav</cp:lastModifiedBy>
</cp:coreProperties>"""

    app_props = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Word</Application>
  <DocSecurity>0</DocSecurity>
  <Lines>95</Lines>
  <Paragraphs>48</Paragraphs>
  <ScaleCrop>false</ScaleCrop>
  <Company>Sanskar Yadav Portfolio</Company>
</Properties>"""

    settings_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
        <w:color w:val="222222"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="80" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>"""

    def p(text="", runs=None, align=None, space_before=40, space_after=60, border_bottom=False):
        xml = ['<w:p>']
        pPr = ['<w:pPr>']
        if align:
            pPr.append(f'<w:jc w:val="{align}"/>')
        if space_before or space_after:
            pPr.append(f'<w:spacing w:before="{space_before}" w:after="{space_after}" w:line="250" w:lineRule="auto"/>')
        if border_bottom:
            pPr.append('<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="1A365D"/></w:pBdr>')
        pPr.append('</w:pPr>')
        xml.append(''.join(pPr))

        if runs:
            for r_text, r_bold, r_italic, r_size, r_color in runs:
                xml.append(r(r_text, bold=r_bold, italic=r_italic, size=r_size, color=r_color))
        elif text:
            xml.append(r(text))
        xml.append('</w:p>')
        return ''.join(xml)

    def r(text, bold=False, italic=False, size=None, color=None):
        clean_text = (text.replace("&", "&amp;")
                          .replace("<", "&lt;")
                          .replace(">", "&gt;")
                          .replace('"', "&quot;"))
        rPr = ['<w:rPr>']
        if bold:
            rPr.append('<w:b/>')
        if italic:
            rPr.append('<w:i/>')
        if size:
            rPr.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
        if color:
            rPr.append(f'<w:color w:val="{color}"/>')
        rPr.append('</w:rPr>')
        return f'<w:r>{"".join(rPr)}<w:t xml:space="preserve">{clean_text}</w:t></w:r>'

    def section_header(title):
        runs = [(title.upper(), True, False, 24, "1A365D")]
        return p(runs=runs, space_before=160, space_after=80, border_bottom=True)

    def bullet_item(lead_bold, body_text):
        xml = ['<w:p>']
        xml.append('<w:pPr><w:ind w:left="400" w:hanging="240"/><w:spacing w:before="25" w:after="35" w:line="240" w:lineRule="auto"/></w:pPr>')
        xml.append(r("▪  ", bold=True, size=18, color="1A365D"))
        if lead_bold:
            xml.append(r(lead_bold + " ", bold=True, size=21, color="111111"))
        xml.append(r(body_text, size=21, color="333333"))
        xml.append('</w:p>')
        return ''.join(xml)

    body_elements = []

    # NAME & HEADER
    body_elements.append(p(runs=[("SANSKAR YADAV", True, False, 36, "0F2942")], align="center", space_before=0, space_after=30))
    body_elements.append(p(runs=[("DATA SCIENCE STUDENT | PYTHON & MACHINE LEARNING | AWS CLOUD PRACTITIONER", True, False, 21, "2B6CB0")], align="center", space_before=0, space_after=60))
    body_elements.append(p(runs=[
        ("Noida, Sector 49, UP  |  ", False, False, 19, "555555"),
        ("+91 7906997240  |  ", False, False, 19, "555555"),
        ("sanskaryadav287@gmail.com", False, False, 19, "1A365D")
    ], align="center", space_before=0, space_after=30))
    body_elements.append(p(runs=[
        ("LinkedIn: https://www.linkedin.com/in/sanskar-yadav-a1546538a/  |  ", False, False, 19, "1A365D"),
        ("GitHub: https://github.com/sanskaryadav287", False, False, 19, "1A365D")
    ], align="center", space_before=0, space_after=140))

    # PROFESSIONAL SUMMARY
    body_elements.append(section_header("Professional Summary"))
    body_elements.append(p(runs=[
        ("Analytical and forward-thinking Data Science student with a comprehensive background in Python, Machine Learning algorithms, statistical modeling, and AWS Cloud infrastructure. "
         "Combines the empirical rigor, precision, and hypothesis-driven research methodology developed during a B.Sc. in Chemistry with advanced Data Science and computational skills acquired at AAFT College, Noida. "
         "Proficient in exploratory data analysis (EDA), predictive modeling (Regression, Classification, Clustering), Natural Language Processing (NLP) foundations, and MongoDB document architecture (Certified). "
         "Demonstrated capability in architecting cloud-based data storage and pipelines using Amazon EC2, S3, IAM, and VPC. Proven communicator and collaborative problem solver backed by 3 years of STEM tutoring and 6 months of healthcare sales experience.", False, False, 21, "333333")
    ], space_before=40, space_after=100))

    # COMPREHENSIVE DATA SCIENCE & TECHNICAL SKILLS
    body_elements.append(section_header("Comprehensive Data Science & Technical Skills"))
    body_elements.append(bullet_item("Programming Languages:", "Python (Expert: NumPy, Pandas, Scikit-Learn, SciPy), SQL (MySQL, PostgreSQL), R (Statistical Analysis Fundamentals)"))
    body_elements.append(bullet_item("Machine Learning (Supervised & Unsupervised):", "Linear Regression, Logistic Regression, Decision Trees, Random Forests, XGBoost, LightGBM, Support Vector Machines (SVM), K-Means Clustering, Hierarchical Clustering, Principal Component Analysis (PCA), Feature Selection & Engineering, Hyperparameter Tuning (GridSearchCV, RandomizedSearchCV)"))
    body_elements.append(bullet_item("Deep Learning & NLP Foundations:", "Artificial Neural Networks (ANN), Convolutional Neural Networks (CNN basics), Text Preprocessing (Tokenization, Lemmatization, Stopword Removal), TF-IDF Vectorization, Word2Vec, Sentiment Analysis, Prompt Engineering & LLM fundamentals"))
    body_elements.append(bullet_item("Statistical Analysis & Mathematics:", "Probability Distributions, Hypothesis Testing (t-tests, ANOVA, Chi-Square), p-value interpretation, Central Limit Theorem, A/B Testing Methodologies, Linear Algebra, Multivariable Calculus for Gradient Descent"))
    body_elements.append(bullet_item("Data Wrangling & Visualization:", "Advanced Exploratory Data Analysis (EDA), Missing Value Imputation, Outlier Detection, Matplotlib, Seaborn, Plotly (Interactive Visuals), Power BI (Dashboards, KPIs), Streamlit for rapid ML app deployment"))
    body_elements.append(bullet_item("AWS Cloud Infrastructure:", "Amazon EC2 (Compute & Linux Server Management), Amazon S3 (Object Storage & Model Bucket Lifecycle), AWS IAM (Least-Privilege Security & Role Policies), Amazon VPC (Network Subnets, Route Tables & Security Groups)"))
    body_elements.append(bullet_item("Databases & Data Modeling:", "MongoDB (Certified: Document Model Design, BSON Schema Validation, Aggregation Pipelines, PyMongo), Relational Database Design, SQL Query Optimization & Joins"))
    body_elements.append(bullet_item("Developer Tools & Platforms:", "Git, GitHub, Docker (Containerization Basics), Jupyter Notebook, Google Colab, VS Code, Linux/Bash CLI, Postman"))

    # AWS CLOUD ARCHITECTURE
    body_elements.append(section_header("AWS Cloud Infrastructure & Services"))
    body_elements.append(bullet_item("Amazon EC2:", "Deployed, configured, and managed Ubuntu/Linux cloud instances, customized security groups, SSH key pairs, and executed background cron ETL jobs."))
    body_elements.append(bullet_item("Amazon S3:", "Engineered scalable data buckets for raw data ingestion, model artifact versioning, lifecycle expiration policies, and static web asset hosting."))
    body_elements.append(bullet_item("AWS IAM:", "Implemented robust identity governance, configured IAM roles, least-privilege permission policies, and secured cloud credentials."))
    body_elements.append(bullet_item("Amazon VPC:", "Designed custom Virtual Private Cloud architectures featuring isolated public and private subnets, Internet Gateways, NAT routing, and stateful firewall rules."))

    # PROJECTS
    body_elements.append(section_header("Practical Data Science & Cloud Projects"))
    
    # Project 1
    body_elements.append(p(runs=[
        ("Customer Churn Prediction & ML Classification Pipeline", True, False, 22, "111111"),
        ("  |  Python, Scikit-Learn, Pandas, SMOTE, AWS S3", False, True, 20, "2B6CB0")
    ], space_before=60, space_after=20))
    body_elements.append(bullet_item("Modeling & Performance:", "Engineered an end-to-end classification system predicting client churn across 10,000+ customer records. Balanced target classes using SMOTE and achieved an 87% ROC-AUC score utilizing Random Forest and XGBoost classifiers."))
    body_elements.append(bullet_item("EDA & Feature Engineering:", "Extracted high-correlation behavioral features, performed one-hot encoding, and conducted extensive cross-validation to prevent data leakage."))
    body_elements.append(bullet_item("Cloud Integration:", "Automated the storage and serialization of model weights and evaluation metrics directly into Amazon S3 buckets for remote inference."))

    # Project 2
    body_elements.append(p(runs=[
        ("Healthcare Document Database Architecture & Aggregations", True, False, 22, "111111"),
        ("  |  MongoDB, PyMongo, SQL-to-NoSQL Migration", False, True, 20, "2B6CB0")
    ], space_before=80, space_after=20))
    body_elements.append(bullet_item("Schema Architecture:", "Designed a production-ready NoSQL document structure adhering to official MongoDB schema design patterns (Subset & Extended Reference), resolving complex multi-table relational join bottlenecks."))
    body_elements.append(bullet_item("Data Pipelines:", "Constructed Python PyMongo aggregation pipelines to calculate patient health indicators and physician workload distribution in real time."))

    # Project 3
    body_elements.append(p(runs=[
        ("Cloud Automated Exploratory Data Analysis & Streamlit System", True, False, 22, "111111"),
        ("  |  Amazon EC2, Python, Streamlit, Seaborn, S3", False, True, 20, "2B6CB0")
    ], space_before=80, space_after=20))
    body_elements.append(bullet_item("Cloud Infrastructure:", "Provisioned an Amazon EC2 Linux instance running scheduled automated data profiling scripts on newly ingested tabular datasets."))
    body_elements.append(bullet_item("Interactive Dashboards:", "Built a lightweight Streamlit web interface displaying distribution plots, correlation heatmaps, and automated missing value summaries."))

    # EXPERIENCE
    body_elements.append(section_header("Professional Experience"))

    # Exp 1
    body_elements.append(p(runs=[
        ("Sales Executive", True, False, 22, "111111"),
        ("  -  Swasthi Health Company", True, False, 21, "2B6CB0"),
        (" (6 Months)", False, True, 20, "555555")
    ], space_before=60, space_after=20))
    body_elements.append(bullet_item("Business Development:", "Managed healthcare product presentations and client relationship workflows, meeting and exceeding monthly lead conversion targets."))
    body_elements.append(bullet_item("Customer Data Insights:", "Evaluated customer purchasing trends and client demographic metrics to segment high-probability prospective clients and optimize pitch strategies."))
    body_elements.append(bullet_item("CRM Process Optimization:", "Maintained accurate CRM pipeline records, optimizing follow-up scheduling and reducing lead response turnaround by 25%."))

    # Exp 2
    body_elements.append(p(runs=[
        ("Private Academic STEM Tutor", True, False, 22, "111111"),
        ("  -  Self-Employed / Independent", True, False, 21, "2B6CB0"),
        (" (3 Years)", False, True, 20, "555555")
    ], space_before=80, space_after=20))
    body_elements.append(bullet_item("STEM Mentorship:", "Instructed 50+ students in Mathematics, Science, and Analytical Problem Solving across secondary and higher secondary curricula."))
    body_elements.append(bullet_item("Quantitative Tracking:", "Designed personalized progress diagnostic trackers and structured practice sets, elevating average student examination performance by 20%."))
    body_elements.append(bullet_item("Communication Leadership:", "Cultivated active listening, pedagogical adaptability, and the ability to articulate complex mathematical theories into clear, actionable intuition."))

    # CERTIFICATIONS
    body_elements.append(section_header("Certifications & Badges"))
    body_elements.append(bullet_item("From Relational Model (SQL) to MongoDB's Document Model", "- MongoDB University (Certified in SQL-to-NoSQL schema conversion, embedded vs reference data modeling)"))
    body_elements.append(bullet_item("MongoDB Schema Design Patterns and Anti-patterns", "- MongoDB University Skill Badge (Mastery in Subset pattern, Bucket pattern, and avoiding unbounded arrays)"))
    body_elements.append(bullet_item("AWS Cloud Practitioner Essentials", "- Architecture practicum spanning EC2, S3, IAM, and VPC networking"))
    body_elements.append(bullet_item("Python for Data Science, ML & Statistics", "- Practical implementation of Scikit-Learn, Pandas, NumPy, and predictive algorithms"))

    # EDUCATION
    body_elements.append(section_header("Education"))
    body_elements.append(p(runs=[
        ("Bachelor of Science (B.Sc.) in Data Science", True, False, 22, "111111"),
        ("  |  AAFT College, Film City, Noida", False, False, 20, "2B6CB0")
    ], space_before=60, space_after=10))
    body_elements.append(p(runs=[
        ("Status: Pursuing (2025 - 2027/2028)  -  Focus: Machine Learning, Statistical Analysis, Cloud Computing, Database Systems", False, True, 19, "555555")
    ], space_before=0, space_after=40))

    body_elements.append(p(runs=[
        ("Bachelor of Science (B.Sc.) in Chemistry", True, False, 22, "111111"),
        ("  |  MJPRU University, Bareilly", False, False, 20, "2B6CB0")
    ], space_before=40, space_after=10))
    body_elements.append(p(runs=[
        ("Status: Completed (2025)  -  Core: Empirical Quantitative Analysis, Scientific Methodology, Analytical Rigor", False, True, 19, "555555")
    ], space_before=0, space_after=40))

    body_elements.append(p(runs=[
        ("Senior Secondary (Class XII)", True, False, 21, "111111"),
        ("  -  Score: ", False, False, 20, "333333"),
        ("63%", True, False, 20, "1A365D"),
        ("    |    Secondary School (Class X)", True, False, 21, "111111"),
        ("  -  Score: ", False, False, 20, "333333"),
        ("73%", True, False, 20, "1A365D")
    ], space_before=40, space_after=60))

    # Assembly
    doc_body = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {''.join(body_elements)}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    with zipfile.ZipFile(filename, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)
        docx.writestr("word/document.xml", doc_body)
        docx.writestr("word/styles.xml", styles_xml)
        docx.writestr("word/settings.xml", settings_xml)
        docx.writestr("docProps/core.xml", core_props)
        docx.writestr("docProps/app.xml", app_props)

    print(f"Successfully updated DOCX: {filename} (Size: {os.path.getsize(filename)} bytes)")

build_docx()
