# PowerPoint Shape Capture Tool

A Python GUI application that captures selected shapes and table data from PowerPoint presentations.

## Features

- Capture selected shapes from PowerPoint slides
- Extract shape properties (ID, bounding box, color, text)
- Handle table data (specific cells, rows, or columns)
- Real-time validation of table inputs
- Export data as JSON with slide screenshots

## Requirements

- Windows OS
- Python 3.7+
- Microsoft PowerPoint

## Installation

1. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Capturing Data

1. Open PowerPoint and select the shapes you want to capture
2. Run the capture application:

```bash
python powerpoint_shape_capture.py
```

3. Fill in the form:
   - **Name**: Enter a name for this capture
   - **Table Cells Section**: (Optional) Add table data
     - Enter the Table ID (shape ID of the table)
     - Select cells, rows, or cols
     - Enter the indices (comma-separated)
     - Add or remove sections as needed

4. Press **OK** (or Enter key) to capture and save the data
5. Press **Exit** (or Esc key) to close the application

### Viewing Test Data

1. Run the test viewer application:

```bash
python test_viewer.py
```

2. Enter the test ID (sequential number from the saved files, e.g., `1`, `2`, `3`)
3. Press **OK** (or Enter key) to load and display the data
4. The viewer will show:
   - Prompt information with all captured data details
   - Slide image with red bounding boxes drawn around captured shapes/tables
5. Adjust the **Box Thickness** slider and click **Apply** to change the outline thickness
6. Press **Exit** (or Esc key) to close the viewer

## Output

The application creates files in the `resources` directory:

- `resources/img/{id}.png` - Screenshot of the current slide (sequential numbering: 1, 2, 3, ...)
- `resources/json/{id}.json` - JSON data with shape and table information (sequential numbering: 1, 2, 3, ...)
- `resources/cloud_presentations/{timestamp}_{filename}.pptx` - Downloaded copy of cloud-based presentations (OneDrive, SharePoint, Teams)

## JSON Format

```json
[
  {
    "name": "Example",
    "path": "C:\\path\\to\\presentation.pptx",
    "slide_number": 1,
    "selection_type": "shape",
    "shape_ids": [1, 2, 3],
    "table_rows": "",
    "table_cols": "",
    "table_cells": "",
    "bbox": [[x, y, width, height], ...],
    "color_rgb": [[r, g, b], ...],
    "text": ["Text 1", "Text 2", ...]
  }
]
```

## Key Features

### Cloud Presentation Support

The application automatically detects if your PowerPoint presentation is stored in the cloud (OneDrive, SharePoint, or Microsoft Teams). When a cloud-based presentation is detected:

- The application automatically downloads a local copy to `resources/cloud_presentations/`
- The local path is used in the JSON output for easier access
- The downloaded file is named with a timestamp to avoid conflicts: `{timestamp}_{original_filename}.pptx`

This ensures you have a local backup and consistent file paths in your JSON data.

### Table Data Validation

- **Cells**: Format must be `row.col` (e.g., `1.1, 2.3`)
- **Rows**: Comma-separated integers (e.g., `1, 2, 3`)
- **Cols**: Comma-separated integers (e.g., `1, 2, 3`)

Invalid formats or non-existent indices will show error messages in red.

### Dynamic Text Preview

Labels next to table input fields show a preview of the cell content:
- Text is truncated to 20 characters with an ellipsis (…)
- Click on the label to expand/collapse the full text

### Key Bindings

- **Enter**: Execute OK button
- **Esc**: Execute Exit button

