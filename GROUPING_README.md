# PowerPoint Shape Grouping Module

This module allows you to group selected red boxes (created by the shape analysis tool) and save the grouping information to a text file.

## Files

- `group_shapes.py` - Comprehensive grouping implementation with detailed logging
- `simple_group_shapes.py` - Simplified version for easy integration
- `workflow_shape_analysis.py` - Complete workflow combining red box creation and grouping
- `test_grouping.py` - Test script to verify grouping functionality
- `server.py` - Updated with grouping API endpoint

## Features

### Shape Grouping
- Reads shape information from red boxes created by `create_shape_boxes.py`
- Extracts original shape IDs from the text inside red boxes
- Calculates group bounds (min left, min top, max right, max bottom)
- Interactive grouping with user input
- Saves groups to timestamped text files

### Workflow
1. **Create Red Boxes**: Use `create_shape_boxes.py` to create red boxes representing shapes
2. **Select Red Boxes**: In PowerPoint, select the red boxes you want to group
3. **Group Shapes**: Run the grouping tool to create groups
4. **Save Groups**: Groups are automatically saved to a text file

## Usage

### Interactive Grouping

```python
import simple_group_shapes

# Start interactive grouping session
groups = simple_group_shapes.group_shapes_interactive()
```

### API Usage

Start the server:
```bash
python server.py
```

Make a POST request to group shapes:
```bash
curl -X POST "http://127.0.0.1:8000/group_shapes" \
     -H "Content-Type: application/json" \
     -d '{"group_name": "My Group"}'
```

### Complete Workflow

```python
import workflow_shape_analysis

# Run complete workflow (red boxes + grouping)
workflow_shape_analysis.run_complete_workflow()
```

## Group Data Structure

Each group contains:
```python
{
    'group_name': 'Group Name',
    'shape_ids': ['123', '456', '789'],  # Original shape IDs
    'shape_count': 3,
    'bounds': {
        'min_left': 100.0,
        'min_top': 50.0,
        'max_right': 300.0,
        'max_bottom': 200.0,
        'width': 200.0,
        'height': 150.0
    }
}
```

## Text File Output

Groups are saved to files named `shape_groups_YYYYMMDD_HHMMSS.txt` with format:

```
Shape Groups - 2025-01-15 14:30:25
==================================================

Group 1: Header Shapes
  Shape IDs: ['123', '456']
  Count: 2
  Bounds: (100.0, 50.0) to (300.0, 150.0)
  Size: 200.0 x 100.0

Group 2: Content Boxes
  Shape IDs: ['789', '101', '102']
  Count: 3
  Bounds: (50.0, 200.0) to (400.0, 350.0)
  Size: 350.0 x 150.0
```

## Interactive Commands

- **Group Name**: Enter a name to create a group with selected shapes
- **'show'**: Display current groups
- **'exit'**: Finish grouping and save to file

## Requirements

- PowerPoint must be open with a presentation
- Red boxes must be created first using `create_shape_boxes.py`
- Red boxes must contain shape information in the format:
  ```
  text: "original text"
  shape_id: "123"
  ```

## Error Handling

- Handles cases where no shapes are selected
- Validates that selected shapes contain valid shape information
- Graceful error messages for common issues
- File save error handling

## Coordinate System

- Uses PowerPoint's coordinate system (points)
- Origin (0,0) at top-left corner
- X increases right, Y increases down
- Bounds calculated from individual shape positions

## Example Workflow

1. **Create red boxes**:
   ```python
   import create_shape_boxes
   result = create_shape_boxes.create_shape_boxes()
   ```

2. **Group shapes**:
   ```python
   import simple_group_shapes
   groups = simple_group_shapes.group_shapes_interactive()
   ```

3. **View results**:
   - Check the generated text file
   - Use the returned group data structure

## Testing

Run the test script to verify functionality:
```bash
python test_grouping.py
```

Make sure to:
1. Have PowerPoint open
2. Select some red boxes
3. Run the test
