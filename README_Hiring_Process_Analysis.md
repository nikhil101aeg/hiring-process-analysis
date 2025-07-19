# 📊 Hiring Process Analytics

## 📘 Project Description
This project analyzes a dataset of job applicants and hiring outcomes using **Microsoft Excel**. It involves cleaning, calculating metrics, identifying trends, and visualizing key aspects of the hiring process — such as gender distribution, salary analysis, and applicant distribution across roles and departments.

---

## 🧱 Dataset Overview
The dataset includes the following fields:
- `Application ID`: Unique identifier for each applicant
- `Interview Taken On`: Date of interview
- `Status`: Outcome (Hired or Rejected)
- `Event Name`: Gender of the applicant
- `Department`: Department applied to
- `Post Name`: Role applied for
- `Offered Salary`: Final salary offer

---

## ⚙️ Tech Stack
- **Microsoft Excel for Microsoft 365** (Version 2505 Build 16.0.18827.20102 - 32-bit)
- Excel Functions: `COUNTIFS`, `SUM`, `COUNT`, `AVERAGE`, `Pivot Table`, `Group Selection`

---

## 🧠 Project Approach
1. **Data Cleaning**: Identified and removed null values. Outliers (salary > $100,000) were filtered out to ensure accuracy in calculations.
2. **Data Transformation**: Converted offered salary to currency format.
3. **Problem Solving**: Tackled multiple business questions using Excel functions and pivot tables.

---

## ✅ Problem Statements & Solutions

### 👥 1. Gender Distribution of Hires
**Goal**: Determine how many male and female applicants were hired.
**Solution**: Used `COUNTIFS` to apply multiple conditions.
```excel
=COUNTIFS(C:C, "Hired", D:D, "Male")
=COUNTIFS(C:C, "Hired", D:D, "Female")
```

### 💵 2. Average Salary Offered
**Goal**: Calculate the average salary offered to applicants.
**Solution**: Used built-in Excel functions.
```excel
=SUM(G:G)/COUNT(G:G)
=AVERAGE(G:G)
```

### 📈 3. Salary Distribution (Class Intervals)
**Goal**: Create salary range buckets and count applicants in each.
**Solution**: Used a **pivot table**, then grouped data using “Group Selection”.
- Start: Lowest salary in dataset
- End: Highest salary
- Interval: $10,000

---

## 📊 Key Insights
- Gender hiring trends are visible via simple filters.
- Average salary was efficiently calculated using Excel built-in formulas.
- Pivot tables provided an easy way to visualize salary distributions.

---

## 📂 Project Files
- `Project-4_Hiring_Process.xlsx` – Dataset with applied Excel queries and visualizations
- `README.md` – Project summary and insights

---

## 👨‍💻 Author
**[Your Name]**  
Aspiring Data Analyst | Excel | Data Cleaning | Visualization

---

🔗 [View on GitHub](https://github.com/your-username/hiring-process-analysis)