package com.gradecalc

import com.gradecalc.parser.ExcelParser
import com.gradecalc.writer.ExcelWriter
import java.io.File

fun main(args: Array<String>) {
    println("╔════════════════════════════════════════════════════╗")
    println("║         Student Grade Calculator                   ║")
    println("╚════════════════════════════════════════════════════╝")
    println()
    println("Step 1 → Reads your Excel file (ID, Name, CA, Exam)")
    println("Step 2 → Calculates grades and shows results in console")
    println("Step 3 → Creates a NEW Excel file with Grade Results only")
    println()

    // ── Get input file path ───────────────────────────────────────────────
    val inputPath: String = if (args.isNotEmpty()) {
        args[0]
    } else {
        print("Enter full path to your Excel file (.xlsx): ")
        readLine()?.trim() ?: ""
    }

    if (inputPath.isBlank()) {
        println("❌ No file path entered. Exiting.")
        return
    }

    val inputFile = File(inputPath)
    if (!inputFile.exists()) {
        println("❌ File not found: ${inputFile.absolutePath}")
        return
    }

    // ── Parse ─────────────────────────────────────────────────────────────
    println("\n📂 Reading: ${inputFile.name} ...")

    val result = try {
        ExcelParser.parse(inputFile)
    } catch (e: Exception) {
        println("❌ Error reading file: ${e.message}")
        return
    }

    // Print warnings
    if (result.errors.isNotEmpty()) {
        println("\n⚠  Warnings:")
        result.errors.forEach { println("   - $it") }
    }

    if (result.students.isEmpty()) {
        println("❌ No valid student data found. Make sure columns are: ID | Name | CA | Exam")
        return
    }

    println("✅ Found ${result.students.size} students.\n")

    // ── Print results to console ──────────────────────────────────────────
    val line = "─".repeat(72)
    println(line)
    println(
        "%-4s  %-22s  %5s  %5s  %6s  %5s  %6s  %-14s".format(
            "No.", "Name", "CA", "Exam", "Total", "Grade", "Points", "Remarks"
        )
    )
    println(line)

    result.students.forEachIndexed { i, s ->
        println(
            "%-4d  %-22s  %5.1f  %5.1f  %6.1f  %5s  %6.1f  %-14s".format(
                i + 1, s.name.take(22), s.ca, s.exam,
                s.total, s.grade, s.gradePoints, s.remarks
            )
        )
    }

    println(line)
    val avg     = result.students.map { it.total }.average()
    val passing = result.students.count { it.grade != "F" }
    val failing = result.students.count { it.grade == "F" }
    println("Class Average : ${"%.2f".format(avg)}")
    println("Passing       : $passing   |   Failing: $failing")
    println(line)

    // ── Generate NEW output file ──────────────────────────────────────────
    // Default output file sits next to the input file e.g. students_grades.xlsx
    val defaultOutput = inputFile
        .resolveSibling("${inputFile.nameWithoutExtension}_grades.xlsx")
        .absolutePath

    print("\nEnter path for the new grades file (press Enter for default [$defaultOutput]): ")
    val userOutput = readLine()?.trim()

// Auto-fix: if user typed a folder path, append the filename automatically
    val resolvedOutput = when {
        userOutput.isNullOrBlank() -> defaultOutput
        File(userOutput).isDirectory -> "$userOutput\\${inputFile.nameWithoutExtension}_grades.xlsx"
        !userOutput.endsWith(".xlsx", ignoreCase = true) -> "$userOutput.xlsx"
        else -> userOutput
    }

    val outputFile = File(resolvedOutput)
    println("📁 Output file: ${outputFile.absolutePath}")

    println("\n💾 Creating new grades file: ${outputFile.name} ...")

    try {
        ExcelWriter.writeToNewFile(result.students, outputFile)
        println("✅ Done! New grades file created at:")
        println("   ${outputFile.absolutePath}")
        println()
        println("📌 Note: Your original file '${inputFile.name}' was NOT modified.")
    } catch (e: Exception) {
        println("❌ Failed to create output file: ${e.message}")
    }
}