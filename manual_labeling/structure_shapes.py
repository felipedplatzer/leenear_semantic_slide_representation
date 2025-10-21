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

def generate_structure_main(dict_list):
    for i, d in enumerate(dict_list):
        d['i'] = i
    dict_list = get_parents(dict_list)
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