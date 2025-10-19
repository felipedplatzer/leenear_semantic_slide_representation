

def _contains(parent, child):
    """
    Check if parent section contains child section based on cells.
    A parent contains a child if the child's cells are a subset of the parent's cells.
    """
    parent_cells = set(parent['cells'])
    child_cells = set(child['cells'])
    
    # Parent must contain all child cells plus at least one more
    is_subset = child_cells.issubset(parent_cells)
    is_smaller = len(child_cells) < len(parent_cells)
    result = is_subset and is_smaller
    
    # Debug output for Group 1 and Big 4 columns
    if parent.get('label') == 'Group 1' and 'Big 4' in child.get('label', ''):
        print(f"  DEBUG _contains: Checking if Group 1 contains {child.get('label')}")
        print(f"    Parent cells (first 10): {sorted(list(parent_cells))[:10]}")
        print(f"    Child cells (first 10): {sorted(list(child_cells))[:10]}")
        print(f"    is_subset={is_subset}, is_smaller={is_smaller}, result={result}")
    
    # Debug output for successful containment
    if result:
        print(f"  DEBUG _contains: '{parent.get('label', '?')}' ({len(parent['cells'])} cells) CONTAINS '{child.get('label', '?')}' ({len(child['cells'])} cells)")
    
    return result

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
    
    # Debug: Print first few cells of sections including some columns
    print("\nDEBUG: Cell contents for sections:")
    col_sections = [s for s in sections if s.get('section_type') in ['col', 'group_of_cols']]
    row_sections = [s for s in sections if s.get('section_type') in ['row', 'group_of_rows']]
    
    if col_sections:
        print(f"  Group/Col example: '{col_sections[0].get('label', '?')}' ({col_sections[0].get('section_type', '?')}): {col_sections[0]['cells'][:5]}...")
        if len(col_sections) > 1:
            print(f"  Col example: '{col_sections[1].get('label', '?')}' ({col_sections[1].get('section_type', '?')}): {col_sections[1]['cells'][:5]}...")
    if row_sections:
        print(f"  Row example: '{row_sections[0].get('label', '?')}' ({row_sections[0].get('section_type', '?')}): {row_sections[0]['cells'][:5]}...")
    
    # Build tree structure
    tree = []
    
    for section in sections:
        # Find the best parent for this section
        # Check all previously processed sections (not just a stack)
        # Find the smallest section that contains this one
        best_parent = None
        best_parent_size = float('inf')
        
        for potential_parent in tree:
            if _contains(potential_parent, section):
                parent_size = len(potential_parent['cells'])
                if parent_size < best_parent_size:
                    best_parent = potential_parent
                    best_parent_size = parent_size
        
        # Add to tree with appropriate index
        if best_parent:
            # This section is nested under the best parent
            parent_index = best_parent['index']
            section['index'] = _get_next_child_index(tree, parent_index)
            print(f"DEBUG structure_table: Section '{section.get('label', '?')}' ({section.get('section_type', '?')}, {len(section['cells'])} cells) nested under parent '{best_parent.get('label', '?')}' index {parent_index}")
        else:
            # This is a root level section
            section['index'] = _get_next_root_index(tree)
            print(f"DEBUG structure_table: Section '{section.get('label', '?')}' ({section.get('section_type', '?')}, {len(section['cells'])} cells) at root level, index {section['index']}")
        
        # Add to tree
        tree.append(section)
    
    return tree

