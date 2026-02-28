package com.gradecalc.writer

import com.gradecalc.model.Student
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File

object ExcelWriter {

    // ── Shared helper to build styles ────────────────────────────────────────

    private fun buildStyles(workbook: Workbook): Map<String, CellStyle> {

        val headerFont: Font = workbook.createFont()
        headerFont.bold = true
        headerFont.fontHeightInPoints = 11
        headerFont.color = IndexedColors.WHITE.index

        val headerStyle: CellStyle = workbook.createCellStyle()
        headerStyle.setFont(headerFont)
        headerStyle.fillForegroundColor = IndexedColors.DARK_BLUE.index
        headerStyle.fillPattern = FillPatternType.SOLID_FOREGROUND
        headerStyle.alignment = HorizontalAlignment.CENTER
        headerStyle.borderBottom = BorderStyle.THIN

        val centerStyle: CellStyle = workbook.createCellStyle()
        centerStyle.alignment = HorizontalAlignment.CENTER

        val numberStyle: CellStyle = workbook.createCellStyle()
        numberStyle.dataFormat = workbook.createDataFormat().getFormat("0.00")
        numberStyle.alignment = HorizontalAlignment.CENTER

        val boldStyle: CellStyle = workbook.createCellStyle()
        val boldFont: Font = workbook.createFont()
        boldFont.bold = true
        boldStyle.setFont(boldFont)

        return mapOf(
            "header" to headerStyle,
            "center" to centerStyle,
            "number" to numberStyle,
            "bold"   to boldStyle
        )
    }

    private fun gradeStyle(workbook: Workbook, grade: String): CellStyle {
        val bgColor: Short = when (grade) {
            "A"  -> IndexedColors.LIGHT_GREEN.index
            "B+" -> IndexedColors.LIGHT_BLUE.index
            "B"  -> IndexedColors.LIGHT_CORNFLOWER_BLUE.index
            "C+" -> IndexedColors.YELLOW.index
            "C"  -> IndexedColors.LIGHT_YELLOW.index
            "D+" -> IndexedColors.TAN.index
            "D"  -> IndexedColors.ORANGE.index
            else -> IndexedColors.ROSE.index
        }
        val style: CellStyle = workbook.createCellStyle()
        style.fillForegroundColor = bgColor
        style.fillPattern = FillPatternType.SOLID_FOREGROUND
        style.alignment = HorizontalAlignment.CENTER
        val font: Font = workbook.createFont()
        font.bold = true
        style.setFont(font)
        return style
    }

    private fun addSummaryRows(
        workbook: Workbook,
        sheet: org.apache.poi.ss.usermodel.Sheet,
        students: List<Student>,
        startRow: Int,
        boldStyle: CellStyle
    ) {
        val classAvg = students.map { it.total }.average()
        val passing  = students.count { it.grade != "F" }
        val failing  = students.count { it.grade == "F" }
        val highest  = students.maxOf { it.total }
        val lowest   = students.minOf { it.total }

        fun addRow(offset: Int, label: String, value: String) {
            val r = sheet.createRow(startRow + offset)
            val labelCell = r.createCell(0)
            labelCell.setCellValue(label)
            labelCell.cellStyle = boldStyle
            r.createCell(1).setCellValue(value)
        }

        addRow(0, "── SUMMARY ──",    "")
        addRow(1, "Total Students",   "${students.size}")
        addRow(2, "Class Average",    "%.2f".format(classAvg))
        addRow(3, "Highest Total",    "%.2f".format(highest))
        addRow(4, "Lowest Total",     "%.2f".format(lowest))
        addRow(5, "Passing Students", "$passing")
        addRow(6, "Failing Students", "$failing")
    }

    // ── Function 1: Write grades back to the SAME file ───────────────────────

    fun writeBack(students: List<Student>, inputFile: File) {

        val workbook: Workbook = inputFile.inputStream().use { XSSFWorkbook(it) }
        val sheet = workbook.getSheetAt(0)
        val styles = buildStyles(workbook)

        val headerStyle = styles["header"]!!
        val centerStyle = styles["center"]!!
        val numberStyle = styles["number"]!!
        val boldStyle   = styles["bold"]!!

        // Add new column headers to existing header row
        val headerRow = sheet.getRow(0) ?: sheet.createRow(0)

        val totalHeader = headerRow.createCell(4)
        totalHeader.setCellValue("Total")
        totalHeader.cellStyle = headerStyle

        val gradeHeader = headerRow.createCell(5)
        gradeHeader.setCellValue("Grade")
        gradeHeader.cellStyle = headerStyle

        val pointsHeader = headerRow.createCell(6)
        pointsHeader.setCellValue("Grade Points")
        pointsHeader.cellStyle = headerStyle

        val remarksHeader = headerRow.createCell(7)
        remarksHeader.setCellValue("Remarks")
        remarksHeader.cellStyle = headerStyle

        // Fill in grade data for each student row
        for (student in students) {
            val row = sheet.getRow(student.rowIndex) ?: sheet.createRow(student.rowIndex)

            val totalCell = row.createCell(4)
            totalCell.setCellValue(student.total)
            totalCell.cellStyle = numberStyle

            val gradeCell = row.createCell(5)
            gradeCell.setCellValue(student.grade)
            gradeCell.cellStyle = gradeStyle(workbook, student.grade)

            val pointsCell = row.createCell(6)
            pointsCell.setCellValue(student.gradePoints)
            pointsCell.cellStyle = numberStyle

            val remarksCell = row.createCell(7)
            remarksCell.setCellValue(student.remarks)
            remarksCell.cellStyle = centerStyle
        }

        addSummaryRows(workbook, sheet, students, students.size + 3, boldStyle)

        for (col in 0..7) sheet.autoSizeColumn(col)

        inputFile.outputStream().use { workbook.write(it) }
        workbook.close()
    }

    // ── Function 2: Create a NEW file with grades only (no CA or Exam) ───────

    fun writeToNewFile(students: List<Student>, outputFile: File) {

        val workbook: Workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Grade Results")
        val styles = buildStyles(workbook)

        val headerStyle = styles["header"]!!
        val centerStyle = styles["center"]!!
        val numberStyle = styles["number"]!!
        val boldStyle   = styles["bold"]!!

        // ── Header Row ────────────────────────────────────────────────────
        // NO CA and NO Exam columns — only grade results
        val headers = listOf(
            "No.",
            "Student Name",
            "Total",
            "Grade",
            "Grade Points",
            "Remarks"
        )

        val headerRow = sheet.createRow(0)
        for ((col, title) in headers.withIndex()) {
            val cell = headerRow.createCell(col)
            cell.setCellValue(title)
            cell.cellStyle = headerStyle
        }

        // ── Data Rows ─────────────────────────────────────────────────────

        for ((index, student) in students.withIndex()) {
            val row = sheet.createRow(index + 1)

            // Col 0 — No.
            val noCell = row.createCell(0)
            noCell.setCellValue((index + 1).toDouble())
            noCell.cellStyle = centerStyle

            // Col 1 — Student Name
            row.createCell(1).setCellValue(student.name)

            // Col 2 — Total (CA + Exam combined)
            val totalCell = row.createCell(2)
            totalCell.setCellValue(student.total)
            totalCell.cellStyle = numberStyle

            // Col 3 — Grade (color coded)
            val gradeCell = row.createCell(3)
            gradeCell.setCellValue(student.grade)
            gradeCell.cellStyle = gradeStyle(workbook, student.grade)

            // Col 4 — Grade Points
            val pointsCell = row.createCell(4)
            pointsCell.setCellValue(student.gradePoints)
            pointsCell.cellStyle = numberStyle

            // Col 5 — Remarks
            val remarksCell = row.createCell(5)
            remarksCell.setCellValue(student.remarks)
            remarksCell.cellStyle = centerStyle
        }

        // ── Summary Section ───────────────────────────────────────────────

        addSummaryRows(workbook, sheet, students, students.size + 3, boldStyle)

        // ── Auto-size all 6 columns & Save ───────────────────────────────

        for (col in 0..5) sheet.autoSizeColumn(col)

        outputFile.parentFile?.mkdirs()
        outputFile.outputStream().use { workbook.write(it) }
        workbook.close()
    }
}