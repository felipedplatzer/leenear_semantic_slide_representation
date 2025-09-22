import re

def get_parent(d, i, dict_list):
    d_id_list = d['shape_id']
    # correct for single item lists
    if type(d_id_list) != list:
        d_id_list = [d_id_list]
    #remove self and items that are not lists
    ref_list = [x for j, x in enumerate(dict_list) if j != i] 
    ref_list = [x for x in ref_list if type(x['shape_id']) == list]
    if len(ref_list) == 0:
        return None
    # get potential parents
    else:
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
            smallest_parent_i_temp = smallest_parent['i']
            return smallest_parent_i_temp


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
        d['index'] = re.sub(r'\d+', lambda m: m.group().zfill(2), d['index'])
    return dict_list