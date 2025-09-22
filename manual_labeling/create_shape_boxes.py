import win32com.client
import resources

def create_shape_boxes(ppt_app):
    """
    Gets the active PowerPoint slide, extracts all shape information,
    and creates a new slide with red boxes representing each original shape.
    """
    
    # Get PowerPoint application
    
    try:
        # Get the active presentation
        presentation = ppt_app.ActivePresentation
        
        # Get the active slide
        active_slide = ppt_app.ActiveWindow.View.Slide
        
        # Get slide dimensions
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight
        
        print(f"Processing slide {active_slide.SlideIndex} of {presentation.Slides.Count}")
        print(f"Slide dimensions: {slide_width} x {slide_height} points")
        
        # Get all shapes from the active slide
        shapes = active_slide.Shapes
        shape_data = []
        
        print(f"Found {shapes.Count} shapes on the slide")
        
        # Extract information from each shape
        for i in range(1, shapes.Count + 1):
            shape = shapes.Item(i)
            
            # Get shape properties
            shape_id = str(shape.Id)
            top = shape.Top
            left = shape.Left
            width = shape.Width
            height = shape.Height
            bottom = top + height
            right = left + width
            
            # Get text content if available
            text_content = ""
            if shape.HasTextFrame:
                if shape.TextFrame.HasText:
                    text_content = shape.TextFrame.TextRange.Text.strip()
            
            shape_info = {
                'shape_id': shape_id,
                'text': text_content,
                'top': top,
                'bottom': bottom,
                'left': left,
                'right': right,
                'width': width,
                'height': height
            }
            
            shape_data.append(shape_info)
            print(f"Shape {i}: ID={shape_id}, Text='{text_content[:50]}...', Position=({left}, {top}), Size=({width}, {height})")
        
        # Create a new slide after the current slide
        current_slide_index = active_slide.SlideIndex
        new_slide = presentation.Slides.AddSlide(current_slide_index + 1, active_slide.CustomLayout)
        
        print(f"Created new slide at index {current_slide_index + 1}")
        
        # Create red boxes for each original shape
        for shape_info in shape_data:
            # Create a rectangle shape
            new_shape = new_slide.Shapes.AddShape(
                Type=1,  # msoShapeRectangle
                Left=shape_info['left'],
                Top=shape_info['top'],
                Width=shape_info['width'],
                Height=shape_info['height']
            )
            
            # Set the shape properties
            # Fill with red color
            new_shape.Fill.ForeColor.RGB = resources.rgb_to_int((255, 0, 0))  # Red
            
            # Set black border
            new_shape.Line.ForeColor.RGB = resources.rgb_to_int((0, 0, 0))  # Black
            new_shape.Line.Weight = 2  # 2 points thick
            
            # Add text to the shape
            if new_shape.HasTextFrame:
                text_frame = new_shape.TextFrame
                text_frame.TextRange.Text = shape_info['shape_id']
                #text_frame.TextRange.Text = f"text: \"{shape_info['text']}\"\nshape_id: \"{shape_info['shape_id']}\""
                
                # Format the text
                text_range = text_frame.TextRange
                text_range.Font.Size = 10
                text_range.Font.Bold = True
                text_range.Font.Color.RGB = resources.rgb_to_int((255, 255, 255))  # White text
                
                
        print(f"Created {len(shape_data)} red boxes on the new slide")
        print("Task completed successfully!")
        
        return {
            'status': 'success',
            'original_slide_index': current_slide_index,
            'new_slide_index': current_slide_index + 1,
            'shapes_processed': len(shape_data),
            'slide_dimensions': {
                'width': slide_width,
                'height': slide_height
            },
            'shape_data': shape_data
        }
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return {
            'status': 'error',
            'message': str(e)
        }
    
    finally:
        # Don't close PowerPoint - let user keep working
        pass

