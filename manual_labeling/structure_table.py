

def _contains(parent, child):
    """
    Check if parent section contains child section based on cells.
    A parent contains a child if the child's cells are a subset of the parent's cells.
    """
    parent_cells = set(parent['cells'])
    child_cells = set(child['cells'])
    
    # Parent must contain all child cells plus at least one more
    return child_cells.issubset(parent_cells) and len(child_cells) < len(parent_cells)

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

def structure_table_sections(table_sections):
    """
    Arrange table sections in a tree structure based on their cells.
    
    Rules:
    1. Table sections that fully encapsulate others (contain all their cells plus more) should be in the level above
    2. Table sections that are encapsulated by the same parent should be ordered by position
    3. Each table section gets an 'index' attribute with tree position
    
    Args:
        table_sections: List of dicts with 'cells' key (array of cell coordinates like "0.1", "1.2")
        
    Returns:
        List of dicts with same attributes plus 'index' attribute
    """
    if not table_sections:
        return []
    
    # Create a copy to avoid modifying original
    sections = [section.copy() for section in table_sections]
    
    # Sort by number of cells (descending) to process larger sections first
    # This ensures parent sections come before their children
    sections.sort(key=lambda x: -len(x['cells']))
    
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

