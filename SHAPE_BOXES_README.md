# PowerPoint Shape Boxes Creator

This module provides functionality to create red boxes representing shapes on PowerPoint slides.

## Files

- `create_shape_boxes.py` - Comprehensive implementation with detailed logging
- `simple_shape_boxes.py` - Simplified version for easy integration
- `test_shape_boxes.py` - Test script to verify functionality
- `server.py` - Updated with new API endpoint

## Features

- Gets the active PowerPoint slide
- Extracts all shapes with their properties:
  - Shape ID
  - Text content
  - Position (top, bottom, left, right)
  - Size (width, height)
- Creates a new slide after the current slide
- Generates red boxes with black borders for each original shape
- Each box contains text showing the original text and shape ID

## Usage

### Direct Python Usage

```python
import simple_shape_boxes

# Create red boxes for shapes on the active slide
result = simple_shape_boxes.create_red_boxes_for_shapes()
print(result)
```

### API Usage

Start the server:
```bash
python server.py
```

Make a POST request to:
```
http://127.0.0.1:8000/create_shape_boxes
```

### Test Script

Run the test script:
```bash
python test_shape_boxes.py
```

## Requirements

- PowerPoint must be open with a presentation
- A slide must be active (selected)
- The slide should contain shapes to process

## Return Values

The function returns a dictionary with:
- `success`: Boolean indicating if the operation succeeded
- `message`: Description of the result
- `original_slide`: Index of the original slide
- `new_slide`: Index of the newly created slide
- `shapes_processed`: Number of shapes that were processed

## Error Handling

The function includes try-catch blocks to handle common errors:
- PowerPoint not running
- No active presentation
- No active slide
- Shape access errors

## Shape Properties Extracted

For each shape, the following properties are captured:
- **ID**: Unique PowerPoint shape identifier
- **Text**: Text content (if any)
- **Position**: Top, bottom, left, right coordinates
- **Size**: Width and height dimensions

## Box Styling

The created red boxes have:
- Red fill color (RGB: 255, 0, 0)
- Black border (RGB: 0, 0, 0) with 2-point thickness
- White text (RGB: 255, 255, 255)
- Centered text alignment
- 10-point bold font
