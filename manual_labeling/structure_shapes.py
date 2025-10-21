import re

def get_parent(d, i, dict_list):
    d_id_list = d['shape_id']
    # correct for single item lists
    if type(d_id_list) != list:
        d_id_list = [d_id_list]
    
    # Check if this is a table section (row, col, or group)
    section_type = d.get('section_type', '')
    is_table_section = section_type in ['row', 'col', 'group_of_rows', 'group_of_cols']
    
    #remove self and items that are not lists
    ref_list = [x for j, x in enumerate(dict_list) if j != i] 
    ref_list = [x for x in ref_list if type(x['shape_id']) == list]
    if len(ref_list) == 0:
        return None
    
    # Special handling for table sections
    if is_table_section:
        # Get the table ID from shape_name field
        table_id = d.get('shape_name')
        if table_id is None:
            # Fallback to standard logic if we can't find table ID
            return get_parent_standard(d, d_id_list, ref_list)
        
        # Find potential parents
        potential_parents = []
        for r in ref_list:
            r_id_list = r['shape_id']
            r_section_type = r.get('section_type', '')
            
            # Check if r is the parent table
            # The table should have shape_id = [table_id] (single element matching table_id)
            if len(r_id_list) == 1 and r_id_list[0] == table_id:
                # This is the table itself
                potential_parents.append(r)
            # Check if r is another table section from the same table that contains this one (superset)
            elif r_section_type in ['row', 'col', 'group_of_rows', 'group_of_cols']:
                # Must be from the same table
                if r.get('shape_name') == table_id:
                    # Use superset/subset logic for groups vs. rows/cols
                    if all([x in r_id_list for x in d_id_list]):
                        potential_parents.append(r)
        
        # Get smallest parent
        if len(potential_parents) == 0:
            return None
        else:
            smallest_parent = min(potential_parents, key=lambda x: len(x['shape_id']))
            return smallest_parent['i']
    else:
        # Standard logic for non-table sections
        return get_parent_standard(d, d_id_list, ref_list)

def get_parent_standard(d, d_id_list, ref_list):
    """Standard parent-finding logic using superset/subset of shape_ids."""
    potential_parents = []
    for r in ref_list:
        r_id_list = r['shape_id']
        if all([x in r_id_list for x in d_id_list]):
            potential_parents.append(r)
    # get smallest parent
    if len(potential_parents) == 0:
        return None
    else:
        smallest_parent = min(potential_parents, key=lambda x: len(x['shape_id']))
        return smallest_parent['i']


def get_parents(dict_list):
    for i,d in enumerate(dict_list):
        d['parent_i'] = get_parent(d, i, dict_list)
    return dict_list

def generate_tree_indexes(dict_list):
    """Generate hierarchical indexes for tree structure based on parent_i relationships."""
    # Sort siblings by top (ascending) then left (ascending)
    def sort_key(item):
        return (item.get('top', 0), item.get('left', 0))
    
    # Build parent-children mapping
    children_map = {}
    for item in dict_list:
        parent_i = item.get('parent_i')
        if parent_i not in children_map:
            children_map[parent_i] = []
        children_map[parent_i].append(item)
    
    # Sort children for each parent
    for parent_i in children_map:
        children_map[parent_i].sort(key=sort_key)
    
    # Assign indexes recursively starting from root (parent_i = None)
    def assign_indexes(parent_i, prefix=""):
        if parent_i not in children_map:
            return
        
        children = children_map[parent_i]
        for i, child in enumerate(children, 1):
            child['index'] = f"{prefix}{i}" if prefix else f"{i}"
            assign_indexes(child['i'], f"{child['index']}.")
    
    # Start with root elements (parent_i = None)
    assign_indexes(None)
    return dict_list

def duplicate_overlaid_shapes_for_all_parents(dict_list):
    """
    Find individual shapes that appear in multiple table sections (overlaid shapes)
    and duplicate them so they appear as children under ALL containing rows and columns.
    
    An overlaid shape is identified by:
    - section_type == 'individual_shape'
    - shape_id (single ID) appears in the shape_id lists of multiple table sections
    
    Note: Overlaid shapes are nested under individual rows and cols only, 
    not under groups of rows or groups of cols.
    """
    print("\n=== DEBUG: Starting duplicate_overlaid_shapes_for_all_parents ===")
    
    # Identify individual shapes with single shape_id
    # These could be shapes with section_type='individual_shape' or shapes without 
    # a table-related section_type that have a single shape_id
    individual_shapes = []
    for i, item in enumerate(dict_list):
        section_type = item.get('section_type', '')
        # Skip table sections themselves
        if section_type in ['row', 'col', 'group_of_rows', 'group_of_cols']:
            continue
        
        shape_ids = item.get('shape_id', [])
        if isinstance(shape_ids, list) and len(shape_ids) == 1:
            parent_i = item.get('parent_i', 'None')
            individual_shapes.append((i, item, shape_ids[0]))
            print(f"DEBUG: Found potential overlaid shape {shape_ids[0]} with section_type='{section_type}', parent_i={parent_i}")
    
    # For each individual shape, find ALL table sections that contain it
    new_sections = []
    shapes_to_remove = set()  # Track which original individual shapes to remove
    
    for orig_index, shape_item, shape_id in individual_shapes:
        # Find all table sections containing this shape_id
        # Only look at individual rows and cols, NOT groups
        containing_sections = []
        for j, section in enumerate(dict_list):
            section_type = section.get('section_type', '')
            if section_type in ['row', 'col']:  # Exclude 'group_of_rows' and 'group_of_cols'
                section_shape_ids = section.get('shape_id', [])
                if isinstance(section_shape_ids, list) and shape_id in section_shape_ids:
                    containing_sections.append((j, section))
                    print(f"  DEBUG: Shape {shape_id} found in {section_type} '{section.get('label', 'unlabeled')}'")
        
        print(f"DEBUG: Shape {shape_id} found in {len(containing_sections)} total rows/cols")
        
        # If this shape appears in multiple table sections (rows/cols), duplicate it
        if len(containing_sections) > 1:
            section_labels = [f"{s[1].get('section_type')} '{s[1].get('label', 'unlabeled')}'" for s in containing_sections]
            print(f"DEBUG: Overlaid shape {shape_id} appears in {len(containing_sections)} rows/cols - duplicating for: {', '.join(section_labels)}")
            shapes_to_remove.add(orig_index)
            
            # Create a copy for each containing section
            # Store the original indices - we'll map them after removal
            for section_index, section in containing_sections:
                shape_copy = shape_item.copy()
                # Store the original parent index - will be remapped later
                shape_copy['_temp_parent_i'] = section_index
                new_sections.append(shape_copy)
    
    # Remove original individual shapes that were duplicated
    if shapes_to_remove:
        print(f"DEBUG: Removing {len(shapes_to_remove)} original overlaid shapes and adding {len(new_sections)} duplicates")
        
        # Build mapping of old indices to new indices
        old_to_new = {}
        new_index = 0
        for i in range(len(dict_list)):
            if i not in shapes_to_remove:
                old_to_new[i] = new_index
                new_index += 1
        
        # Filter out the removed shapes
        filtered_list = [item for i, item in enumerate(dict_list) if i not in shapes_to_remove]
        
        # Update parent_i references in all items
        for item in filtered_list:
            if item.get('parent_i') is not None:
                old_parent = item['parent_i']
                if old_parent in old_to_new:
                    item['parent_i'] = old_to_new[old_parent]
                else:
                    # Parent was removed (shouldn't happen, but handle it)
                    print(f"WARNING: Parent {old_parent} was removed, setting parent to None")
                    item['parent_i'] = None
        
        # Update parent_i references in new_sections using the mapping
        for item in new_sections:
            old_parent = item.pop('_temp_parent_i')
            if old_parent in old_to_new:
                item['parent_i'] = old_to_new[old_parent]
            else:
                # Parent was removed (shouldn't happen, but handle it)
                print(f"WARNING: Parent {old_parent} was removed, setting parent to None")
                item['parent_i'] = None
        
        # Re-assign 'i' values
        for i, item in enumerate(filtered_list):
            item['i'] = i
        
        # Add new sections with their 'i' values
        start_i = len(filtered_list)
        for j, item in enumerate(new_sections):
            item['i'] = start_i + j
        
        dict_list = filtered_list + new_sections
        print(f"DEBUG: After duplication, dict_list has {len(dict_list)} items")
        
        # Show what was created
        for item in new_sections:
            shape_id = item.get('shape_id', [None])[0]
            parent_i = item.get('parent_i')
            parent_label = 'None'
            if parent_i is not None and parent_i < len(filtered_list):
                parent = filtered_list[parent_i]
                parent_label = f"{parent.get('section_type', '?')} '{parent.get('label', 'unlabeled')}'"
            print(f"  Created duplicate of {shape_id} with parent {parent_i} ({parent_label})")
    else:
        print("DEBUG: No overlaid shapes found to duplicate")
    
    print("=== DEBUG: Finished duplicate_overlaid_shapes_for_all_parents ===\n")
    return dict_list

def generate_structure_main(dict_list):
    for i, d in enumerate(dict_list):
        d['i'] = i
    dict_list = get_parents(dict_list)
    
    # NEW: Duplicate overlaid shapes so they appear under all containing sections
    dict_list = duplicate_overlaid_shapes_for_all_parents(dict_list)
    
    dict_list = generate_tree_indexes(dict_list)
    for d in dict_list:
        # Only format index if it exists
        if 'index' in d:
            d['index'] = re.sub(r'\d+', lambda m: m.group().zfill(2), d['index'])
    return dict_list

def trim_tree(dict_list):
    """
    Remove elements that appear both at root level and nested within other elements.
    After removal, regenerate indices to reflect the new structure.
    """
    # Find all shape_ids that are nested (have a parent)
    nested_shape_ids = set()
    for item in dict_list:
        if item.get('parent_i') is not None:
            # This item is nested, collect its shape_ids
            shape_ids = item.get('shape_id', [])
            if isinstance(shape_ids, list):
                nested_shape_ids.update(shape_ids)
            else:
                nested_shape_ids.add(shape_ids)
    
    # Remove root-level items whose shape_ids are all nested elsewhere
    filtered_list = []
    for item in dict_list:
        # Keep items that have a parent (they're nested)
        if item.get('parent_i') is not None:
            filtered_list.append(item)
        else:
            # For root-level items, only keep if their shape_ids are NOT all nested
            item_shape_ids = item.get('shape_id', [])
            if not isinstance(item_shape_ids, list):
                item_shape_ids = [item_shape_ids]
            
            # Check if ANY of the shape_ids are not nested
            has_non_nested = any(sid not in nested_shape_ids for sid in item_shape_ids)
            if has_non_nested:
                filtered_list.append(item)
    
    # Regenerate i, parent_i, and indexes
    # Create mapping from old i to new i
    old_to_new_i = {}
    for new_i, d in enumerate(filtered_list):
        old_to_new_i[d['i']] = new_i
        d['i'] = new_i
    
    # Update parent_i references using the mapping
    for d in filtered_list:
        if d.get('parent_i') is not None:
            d['parent_i'] = old_to_new_i.get(d['parent_i'], d['parent_i'])
    
    # Regenerate tree indexes
    filtered_list = generate_tree_indexes(filtered_list)
    
    # Format indexes with zero-padding
    for d in filtered_list:
        if 'index' in d:
            d['index'] = re.sub(r'\d+', lambda m: m.group().zfill(2), d['index'])
    
    return filtered_list