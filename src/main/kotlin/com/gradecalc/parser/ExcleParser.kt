package com.gradecalc.parser

import com.gradecalc.model.Student
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File

/**
 * Reads the Excel file.
 *
 * Expected columns:
 * A = ID | B = Student Name | C = CA | D = Exam | E = Grade (empty, will be filled)
 * Row 1 = Header (skipped)
 * Row 2 onward = Student data
 */
object ExcelParser {

    data class ParseResult(
        val students: List<Student>,
        val errors: List<String>
    )

    fun parse(file: File): ParseResult {
        require(file.exists()) { "File not found: ${file.absolutePath}" }
        require(file.extension.lowercase() in listOf("xlsx", "xls")) {
            "Only .xlsx and .xls files are supported."
        }

        val workbook = WorkbookFactory.create(file)
        val sheet = workbook.getSheetAt(0)
        val students = mutableListOf<Student>()
        val errors = mutableListOf<String>()

        // Start from row 1 — skip header row 0
        for (rowIdx in 1..sheet.lastRowNum) {
            val row = sheet.getRow(rowIdx) ?: continue

            // Column A (index 0) = ID
            val idCell = row.getCell(0)
            val id: Int = when {
                idCell == null -> rowIdx
                idCell.cellType == CellType.NUMERIC -> idCell.numericCellValue.toInt()
                idCell.cellType == CellType.STRING ->
                    idCell.stringCellValue.trim().toIntOrNull() ?: rowIdx
                else -> rowIdx
            }

            // Column B (index 1) = Student Name
            val nameCell = row.getCell(1)
            val name: String? = when {
                nameCell == null -> null
                nameCell.cellType == CellType.STRING -> nameCell.stringCellValue.trim()
                nameCell.cellType == CellType.NUMERIC -> nameCell.numericCellValue.toInt().toString()
                else -> null
            }

            if (name.isNullOrBlank()) {
                errors.add("Row ${rowIdx + 1}: skipped — no student name in column B.")
                continue
            }

            // Column C (index 2) = CA score
            val caCell = row.getCell(2)
            val ca: Double? = when {
                caCell == null -> null
                caCell.cellType == CellType.NUMERIC -> caCell.numericCellValue
                caCell.cellType == CellType.STRING ->
                    caCell.stringCellValue.trim().toDoubleOrNull()
                else -> null
            }

            if (ca == null) {
                errors.add("Row ${rowIdx + 1} ($name): skipped — no CA score in column C.")
                continue
            }

            // Column D (index 3) = Exam score
            val examCell = row.getCell(3)
            val exam: Double? = when {
                examCell == null -> null
                examCell.cellType == CellType.NUMERIC -> examCell.numericCellValue
                examCell.cellType == CellType.STRING ->
                    examCell.stringCellValue.trim().toDoubleOrNull()
                else -> null
            }

            if (exam == null) {
                errors.add("Row ${rowIdx + 1} ($name): skipped — no Exam score in column D.")
                continue
            }

            students.add(
                Student(
                    id = id,
                    name = name,
                    ca = ca.coerceIn(0.0, 40.0),
                    exam = exam.coerceIn(0.0, 60.0),
                    rowIndex = rowIdx
                )
            )
        }

        workbook.close()
        return ParseResult(students, errors)
    }
}