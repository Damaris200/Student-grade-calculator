# 🎓 Student Grade Calculator

> A Kotlin console application that reads student scores from an Excel file, calculates grades automatically, and writes the results back into the same file.


## 📖 About the Project

This is a **pair programming** assignment that produces a **Student Grade Calculator** in Kotlin.

- Takes an Excel file (`.xlsx`) as input containing student ID, Name, CA score and Exam score
- Calculates each student's total and assigns a grade automatically
- Writes the grade, grade points and remarks **back into the same Excel file**
- Prints a full summary in the console including class average, pass and fail count


**Step by step flow:**

1. Program asks for the path to your Excel file
2. `ExcelParser` reads each student row — ID, Name, CA, Exam
3. `Student` model calculates the total (`CA + Exam`) and determines the grade
4. Results are printed to the console in a formatted table
5. `ExcelWriter` opens the same file and fills in Total, Grade, Grade Points and Remarks columns
6. A summary section is added at the bottom of the sheet

---

## 📊 Grading Scale

| Grade | Total Score | Grade Points | Remarks       |
|-------|------------|--------------|---------------|
| A     | 80 – 100   | 4.0          | Excellent     |
| B+    | 70 – 79    | 3.5          | Very Good     |
| B     | 60 – 69    | 3.0          | Good          |
| C+    | 55 – 59    | 2.5          | Above Average |
| C     | 50 – 54    | 2.0          | Average       |
| D+    | 45 – 49    | 1.5          | Below Average |
| D     | 40 – 44    | 1.0          | Pass          |
| F     | 0  – 39    | 0.0          | Fail          |



### 1. Clone the repository

```bash
git clone https://github.com/your-username/Student-grade.git
cd Student-grade
```

### 2. Open in IntelliJ IDEA

- Open IntelliJ IDEA
- Click **File → Open** and select the `Student-grade` folder
- Wait for Gradle to sync automatically

### 3. Set the JDK

- Go to **File → Project Structure**
- Under **Project SDK** select **temurin-17** or any JDK 17
- Click **Apply → OK**

### 4. Set Gradle JVM

- Go to **File → Settings → Build, Execution, Deployment → Build Tools → Gradle**
- Set **Gradle JVM** to **temurin-17**
- Click **Apply → OK**

---

## ▶️ Usage

### 1. Prepare your Excel file

Create an Excel file (`.xlsx`) in Microsoft Excel with this format:

| A  | B            | C  | D    |
|----|-------------|-----|------|
| ID | Student Name | CA | Exam |
| 1  | John Doe    | 35  | 52   |
| 2  | Jane Smith  | 38  | 57   |

> ⚠️ Make sure to **close the Excel file** before running the program.

### 2. Run the program

- Open `Main.kt` in IntelliJ
- Click the green ▶️ **Run** button next to `fun main`
- The console will prompt:

```
Enter full path to your Excel file (.xlsx):
```

### 3. Enter the file path

Type the full path to your Excel file — **no quotes**:

```
C:\Users\yourname\Desktop\students.xlsx
```

### 4. Check the results

- The console prints a full grade table
- Open your Excel file — it now has **Total, Grade, Grade Points and Remarks** filled in automatically

---

## 📄 Excel File Format

### Input (what you provide)

| ID | Student Name  | CA | Exam |
|----|---------------|----|------|
| 1  | John Doe      | 35 | 52   |
| 2  | Jane Smith    | 38 | 57   |

- **Column A** = Student ID
- **Column B** = Student Name  
- **Column C** = CA score (max 40)
- **Column D** = Exam score (max 60)
- **Row 1** = Header (skipped by the program)

### Output (what the program fills in)

| ID | Student Name | Total| Grade | Grade Points | Remarks   |
|----|--------------|------|-------|--------------|-----------|
| 1  | John Doe      87.00 | A     | 4.0          | Excellent |
| 2  | Jane Smith    95.00 | A     | 4.0          | Excellent |


