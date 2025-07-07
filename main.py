import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Absolute paths to data files
students_path = r"C:\Users\Vishal M\Desktop\PSGCT 2023-2027\Projects\Student Information Report\students.csv"
grades_path = r"C:\Users\Vishal M\Desktop\PSGCT 2023-2027\Projects\Student Information Report\grades.xlsx"
db_path = r"C:\Users\Vishal M\Desktop\PSGCT 2023-2027\Projects\Student Information Report\student_info_system.db"

# 1. Load data
students_df = pd.read_csv(students_path)
grades_df = pd.read_excel(grades_path, engine="openpyxl")

# 2. Connect to SQLite
conn = sqlite3.connect(db_path)

# 3. Store data in SQLite
students_df.to_sql("students", conn, if_exists="replace", index=False)
grades_df.to_sql("grades", conn, if_exists="replace", index=False)

# 4. Data Cleaning
students_df = students_df.fillna({"age": students_df["age"].mean(), "email": "noemail@domain.com"})
grades_df.dropna(subset=["grade"], inplace=True)

# 5. Summary Statistics
print("\n--- Students Summary ---")
print(students_df.describe())
print("\n--- Grades Summary ---")
print(grades_df.describe())

# 6. Visualizations
plt.figure(figsize=(6, 4))
sns.histplot(students_df["age"], bins=10, kde=True)
plt.title("Age Distribution")
plt.xlabel("Age")
plt.ylabel("Frequency")
plt.tight_layout()
plt.show()

plt.figure(figsize=(6, 4))
sns.histplot(grades_df["grade"], bins=10, kde=True)
plt.title("Grade Distribution")
plt.xlabel("Grade")
plt.ylabel("Frequency")
plt.tight_layout()
plt.show()

# 7. Aggregation
avg_grade_by_course = grades_df.groupby("course_id").agg({"grade": "mean"})
print("\n--- Average Grade by Course ---")
print(avg_grade_by_course)

# 8. Advanced Analysis
grade_matrix = grades_df.pivot_table(index="student_id", columns="course_id", values="grade")
cov_matrix = grade_matrix.cov()
print("\n--- Covariance Matrix ---")
print(cov_matrix)

# 9. Visualization
plt.figure(figsize=(8, 5))
avg_grade_by_course.plot(kind="bar", legend=False)
plt.title("Average Grade by Course")
plt.xlabel("Course ID")
plt.ylabel("Average Grade")
plt.tight_layout()
plt.show()

merged_df = pd.merge(students_df, grades_df, on="student_id")
print("\n--- Merged Student and Grades ---")
print(merged_df.head())

# 10. SQL Reporting
query = """
SELECT students.*, AVG(grades.grade) AS avg_grade
FROM students
LEFT JOIN grades ON students.student_id = grades.student_id
GROUP BY students.student_id
ORDER BY students.student_id
"""
report_df = pd.read_sql_query(query, conn)
print("\n--- SQL Report ---")
print(report_df)
# Save Report to XLSX
# Export selected outputs to an Excel workbook
excel_path = r"C:\Users\Vishal M\Desktop\PSGCT 2023-2027\Projects\Student Information Report\student_report.xlsx"

with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
    # Summary statistics for students and grades
    students_df.describe().to_excel(writer, sheet_name="Students_Summary")
    grades_df.describe().to_excel(writer, sheet_name="Grades_Summary")
    
    # Final SQL-based student performance report
    report_df.to_excel(writer, sheet_name="SQL_Report", index=False)

print(f"\n Excel report saved to:\n{excel_path}")
#Close the connection
conn.close()
