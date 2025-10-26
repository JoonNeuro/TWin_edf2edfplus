# TWIN EDF to EDF+ Converter

A comprehensive Python toolkit for converting EDF files to EDF+ format with Excel event integration, designed specifically for EEG data processing workflows.

## Overview

This project provides automated conversion of EDF (European Data Format) files to EDF+ format by integrating event data from Excel files. It includes preprocessing tools for relative time calculation, duration correction, and comprehensive logging for debugging and analysis.

## Features

- **EDF to EDF+ Conversion**: Convert EDF files to EDF+ format using MNE Python
- **Excel Event Integration**: Automatically match and integrate events from Excel files
- **Duration Correction**: Fix duration mismatches between EDF header and actual data
- **Zero Padding Detection**: Exclude events in zero-padded areas
- **Event Status Tracking**: Track which events are included/excluded and why
- **Batch Processing**: Process multiple EDF files recursively
- **Backup Management**: Automatic backup of original files
- **Comprehensive Logging**: Detailed logging with timestamps for debugging
- **Rollback Support**: Easy restoration of original files

## Project Structure

```
twin_edf2edfplus/
├── edf2edfplus.py          # Main conversion script
├── relative.py              # Relative time calculation for Excel files
├── rollback.py             # Rollback utility
├── edf2set.m               # MATLAB conversion script
├── README.md               # This file
├── .gitignore              # Git ignore rules
└── patient_directories/     # Patient data directories (ignored by git)
    ├── PATIENT_A/          # Patient A data
    └── PATIENT_B/          # Patient B data
```

**Note**: Patient data directories and files are ignored by git for privacy and security reasons. Only the core Python scripts, MATLAB files, and documentation are tracked.

## Requirements

### Python Dependencies

- `mne` - EEG/MEG data processing
- `pandas` - Excel file processing
- `numpy` - Numerical operations
- `openpyxl` - Excel file reading

### Installation

```bash
pip install mne pandas numpy openpyxl
```

## Usage

### 1. Main Conversion

Run the EDF to EDF+ conversion:

```bash
python edf2edfplus.py
```

This will:

- Automatically run `relative.py` to calculate relative times
- Process all EDF files in the directory and subdirectories
- Convert EDF files to EDF+ format with integrated events
- Fix duration mismatches and apply zero padding when needed
- Track event inclusion/exclusion status
- Create backup files automatically
- Generate detailed logs for monitoring

### 2. Manual Preprocessing (Optional)

You can also run `relative.py` separately to preprocess Excel files:

```bash
python relative.py
```

This step:

- Processes all Excel files in the directory and subdirectories
- Calculates relative times based on EDF metadata start time
- Adds relative times to Column E and status to Column F of each Excel file
- Can be run independently for testing or debugging

### 3. Rollback (if needed)

To restore original files and remove converted files:

```bash
python rollback.py
```

## File Naming Convention

The system expects files to follow this naming pattern:

- **EDF files**: `{patient_number}_{YYYYMMDD}_{HHMM}.edf`
- **Excel files**: `{patient_number}_{YYYYMMDD}_{HHMM}.xlsx`

### Example:

```
PATIENT_A_20190717_1000.edf
PATIENT_A_20190717_1000.xlsx
PATIENT_B_20210218_0931.edf
PATIENT_B_20210218_0931.xlsx
```

## Excel File Format

### Event File Structure

- **Column C**: Time stamps (HH:MM:SS.ss format)
- **Column D**: Event descriptions (e.g., "Asleep", "Eyes Closed", "Move", "Photic On/Off")
- **Column E**: Relative times (auto-calculated from EDF start time)
- **Column F**: Event status (auto-updated during conversion)
  - `✅ INCLUDED`: Event successfully added to EDF+
  - `❌ ZERO_PADDING`: Event in zero-padded area (excluded)
  - `❌ BEFORE_START`: Event before EDF start time (excluded)
  - `❌ AFTER_END`: Event after EDF end time (excluded)
  - `⏳ PENDING`: Event pending processing
  - `❌ NO_TIME`: Invalid time format

## Key Features Explained

### Duration Correction

The system automatically detects and fixes duration mismatches:

- **Header vs Data**: Compares EDF header duration with actual data duration
- **Zero Padding**: Adds zero padding when data is shorter than header duration
- **Threshold**: Uses 0.01-second threshold to detect meaningful differences
- **Event Exclusion**: Excludes events in zero-padded areas (no real signal)

### Event Status Tracking

Each event is tracked with detailed status information:

- **Inclusion Tracking**: Shows which events were successfully added to EDF+
- **Exclusion Reasons**: Provides specific reasons for excluded events
- **Excel Updates**: Automatically updates Excel files with status information
- **Debugging Support**: Helps identify timing and data issues

### Relative Time Calculation

The `relative.py` script processes all Excel files to calculate relative event times:

- **Reference Point**: Uses EDF metadata start time as the reference (0 seconds)
- **Auto-calculation**: Calculates relative times for all events
- **Excel Integration**: Adds calculated relative times to Column E and status to Column F
- **EDF-based**: Uses actual EDF start time rather than first event time

## Logging

All operations are logged to timestamped files:

- Format: `log_edf2edfplus_YYYYMMDDHHMMSS.log`
- Includes timestamps, processing steps, and debugging information
- Both console and file output for real-time monitoring

## MATLAB Integration

The project includes `edf2set.m` for MATLAB-based EDF to EEGLAB SET conversion, providing an alternative workflow for EEG analysis.

## Error Handling

The system includes robust error handling for:

- Missing Excel files
- Invalid time formats
- Duration mismatches
- Zero padding detection
- Event status tracking
- File I/O errors

## Troubleshooting

### Common Issues

1. **Missing Excel files**: Ensure Excel files follow the correct naming convention
2. **Time format errors**: Verify time stamps are in HH:MM:SS.ss format
3. **Duration mismatches**: The system automatically handles MNE duration reading issues
4. **Zero padding events**: Events in zero-padded areas are automatically excluded
5. **Event status**: Check Column F in Excel files for detailed inclusion/exclusion reasons

### Debug Information

Check the log files for detailed processing information:

- Event loading and validation
- Duration calculations and corrections
- Zero padding detection
- Event timing adjustments
- File conversion status

## Data Privacy & Security

- **Patient Data**: All patient data directories and files are excluded from git tracking
- **Privacy Protection**: Only core scripts and documentation are version controlled
- **Local Processing**: All data processing happens locally on your machine
- **No Data Transmission**: No patient data is transmitted or stored externally

## Contributing

This project is designed for specific EEG data processing workflows. For modifications or improvements, please ensure compatibility with the existing file structure and naming conventions.

## License

This project is intended for research and educational purposes. Please ensure compliance with data privacy regulations when processing patient data.

## Support

For issues or questions regarding this toolkit, please refer to the log files for detailed error information and processing status.
