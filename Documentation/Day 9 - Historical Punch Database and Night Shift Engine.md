# Day 9 - Historical Attendance Database & Night Shift Processing

## Objective

Continue building the attendance processing engine by introducing a permanent historical punch database and implementing the first version of the Night Shift processing engine.

The objective was to eliminate duplicate records, preserve historical punch data, and automatically rebuild attendance databases from accumulated attendance history.

---

## Features Implemented

### 1. Permanent Raw Punch Database

Implemented a historical attendance database that stores every unique ACS punch.

Instead of processing only the latest ACS report, the system now maintains a cumulative punch history.

Benefits:

- Historical attendance is preserved
- Previous day punch-outs can be detected automatically
- Duplicate imports no longer create duplicate records

---

### 2. Duplicate Punch Prevention

Implemented duplicate validation before inserting records into the Raw Punch Database.

Each punch is validated using:

- Employee ID
- Punch Date
- Punch Time

If the same punch already exists, it is skipped.

---

### 3. Attendance Database Rebuild Architecture

Changed the processing workflow from append mode to rebuild mode.

Every time Generate Attendance runs:

1. Import latest ACS report
2. Update Raw Punch Database
3. Clear Morning Shift Database
4. Clear Night Shift Database
5. Rebuild attendance databases from historical punches

This guarantees data consistency and automatically updates incomplete attendance records when new punches become available.

---

### 4. Morning Shift Engine Updated

The Morning Shift engine now processes attendance using the historical Raw Punch Database instead of the temporary Raw Import sheet.

Benefits:

- Previous day punch-outs are detected automatically
- Attendance remains accurate across multiple ACS uploads
- Attendance records are rebuilt from complete punch history

---

### 5. Night Shift Processing Engine

Implemented the initial Night Shift processing engine.

Features:

- Night Shift dictionary creation
- Night IN detection
- Night OUT detection
- Attendance date calculation
- Total hour calculation
- Present / UNK status generation

---

### 6. Time Value Handling Fix

Discovered that Excel stores time-only values as numeric Double values instead of Date values.

Updated the engine to correctly process both Date and Numeric time formats, ensuring reliable time detection during Night Shift processing.

---

## Important Discovery

During testing, a business rule conflict was identified.

The following time period overlaps:

Morning Shift Punch Out:
00:00 AM – 03:00 AM

Night Shift Punch In:
00:00 AM – 03:00 AM

A midnight punch cannot be classified correctly using time alone.

Example:

12:30 AM may represent:

- Morning Shift Punch Out
- Night Shift Punch In

This revealed a limitation in maintaining separate Morning and Night processing engines.

---

## Architectural Decision

After analysing the overlap issue, the processing strategy will be redesigned.

Instead of classifying individual punches, the next implementation will classify complete attendance sessions.

The future engine will:

- Read all punches for an employee
- Build one attendance session
- Determine the correct shift
- Write the record to the appropriate attendance database

This session-based architecture eliminates midnight ambiguity and provides a more reliable attendance engine.

---

## Current Workflow

ACS Report

↓

Raw Import

↓

Raw Punch Database (Permanent History)

↓

Attendance Processing Engine

↓

Morning Shift Database

↓

Night Shift Database

---

## Result

The attendance engine now supports:

- Historical punch storage
- Duplicate prevention
- Automatic attendance rebuilding
- Morning Shift processing
- Initial Night Shift processing

The project is now ready for the session-based attendance engine refactoring planned for Day 10.