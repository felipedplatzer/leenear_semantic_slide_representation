

def _contains(parent, child):
    """Check if parent section contains child section"""
    return (parent['start_char'] <= child['start_char'] and 
            parent['end_char'] >= child['end_char'] and
            not (parent['start_char'] == child['start_char'] and parent['end_char'] == child['end_char']))

def _get_next_root_index(tree):
    """Get next root level index (01, 02, 03, ...)"""
    root_sections = [s for s in tree if '.' not in s.get('index', '')]
    return f"{len(root_sections) + 1:02d}"

def _get_next_child_index(tree, parent_index):
    """Get next child index under parent (01.01, 01.02, ...)"""
    # Find all children of this parent
    children = [s for s in tree if s.get('index', '').startswith(parent_index + '.')]
    
    # Get the next child number
    if children:
        # Extract the last part of the index and increment
        last_child = children[-1]['index']
        last_number = int(last_child.split('.')[-1])
        next_number = last_number + 1
    else:
        next_number = 1
    
    return f"{parent_index}.{next_number:02d}"

def structure_text_sections(text_sections):
    """
    Arrange text sections in a tree structure based on their position.
    
    Rules:
    1. Text sections that fully encapsulate others should be in the level above
    2. Text sections that are encapsulated by the same parent should be ordered by position
    3. Each text section gets an 'index' attribute with tree position
    
    Args:
        text_sections: List of dicts with 'start_char' and 'end_char' keys
        
    Returns:
        List of dicts with same attributes plus 'index' attribute
    """
    if not text_sections:
        return []
    
    # Create a copy to avoid modifying original
    sections = [section.copy() for section in text_sections]
    
    # Sort by start position, then by end position (descending for proper nesting)
    sections.sort(key=lambda x: (x['start_char'], -x['end_char']))
    
    # Build tree structure
    tree = []
    stack = []  # Stack to track parent sections
    
    for section in sections:
        # Remove sections from stack that don't contain this section
        while stack and not _contains(stack[-1], section):
            stack.pop()
        
        # Add to tree with appropriate index
        if stack:
            # This section is nested under the top of stack
            parent_index = stack[-1]['index']
            section['index'] = _get_next_child_index(tree, parent_index)
        else:
            # This is a root level section
            section['index'] = _get_next_root_index(tree)
        
        # Add to tree
        tree.append(section)
        
        # Push to stack if this section could contain others
        stack.append(section)
    
    return tree
