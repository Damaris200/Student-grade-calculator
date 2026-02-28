package com.gradecalc.model

data class Student(
    val id: Int,
    val name: String,
    val ca: Double,       // Continuous Assessment — max 40
    val exam: Double,     // Exam score — max 60
    val rowIndex: Int     // Row number in Excel (for writing back)
) {
    val total: Double = ca + exam          // Total out of 100
    val grade: String = calculateGrade(total)
    val remarks: String = gradeRemarks(grade)
    val gradePoints: Double = gradeToPoints(grade)
}

fun calculateGrade(total: Double): String = when {
    total >= 80 -> "A"
    total >= 70 -> "B+"
    total >= 60 -> "B"
    total >= 55 -> "C+"
    total >= 50 -> "C"
    total >= 45 -> "D+"
    total >= 40 -> "D"
    else        -> "F"
}

fun gradeRemarks(grade: String): String = when (grade) {
    "A"  -> "Excellent"
    "B+" -> "Very Good"
    "B"  -> "Good"
    "C+" -> "Above Average"
    "C"  -> "Average"
    "D+" -> "Below Average"
    "D"  -> "Pass"
    else -> "Fail"
}

fun gradeToPoints(grade: String): Double = when (grade) {
    "A"  -> 4.0
    "B+" -> 3.5
    "B"  -> 3.0
    "C+" -> 2.5
    "C"  -> 2.0
    "D+" -> 1.5
    "D"  -> 1.0
    else -> 0.0
}