# Day 4 - Header Mapping Engine

## Objective:

The ACS Report may change column headers in the future versions.
Hardcoded column references would cause the system to fail.

The Header Mapping Engine was implemented to dynamically identify column locations based on header names.

## Implementation

A Dictionary object (gHeaderMap) is created.

The system scans Row 1 of the Raw Import sheet.

Each header name is stored together with its column number.

Example:

``` Text
Employee ID -> 1
First Name -> 2
Date -> 4
Time -> 5
```

## Validation

The system verifies that all mandatory headers exist:

- Employee ID
- First Name
- Department
- Date
- Time
- Punch State

If any required header is missing, processing stops and an error message is displayed.

## Result

The Attendance Automation System is now independent of ACS column order changes.
Future modules can reference columns dynamically using gHeaderMap.

