import os
import json

# Get the directory where this script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Paths
training_dir = os.path.join(script_dir, 'training_data')

# One image or separate images - toggle this to True or False
ONE_IMAGE_BOOL = False


def get_json_dl(json_path):
    with open(json_path, 'r') as f:
        return json.load(f)

def get_tree_from_index(flat_list):
    """
    Convert a flat list of objects with hierarchical indices into a tree structure.
    
    Input: [{'index': '1', 'label': 'Introduction'}, {'index': '1.1', 'label': 'Overview'}, ...]
    Output: [{'label': 'Introduction', 'sections': [{'label': 'Overview', 'sections': []}]}]
    """
    # Sort by index to ensure proper order
    sorted_items = sorted(flat_list, key=lambda x: [int(part) for part in str(x['index']).split('.')])
    
    # Create a dictionary to store nodes by their full path
    nodes = {}
    root_nodes = []
    
    for item in sorted_items:
        index_parts = [int(part) for part in str(item['index']).split('.')]
        current_path = []
        
        # Build the tree structure
        for i, part in enumerate(index_parts):
            current_path.append(str(part))
            path_key = '.'.join(current_path)
            
            # Create node if it doesn't exist
            if path_key not in nodes:
                node = {
                    'label': item['label'] if i == len(index_parts) - 1 else f"Section {path_key}",
                    'sections': []
                }
                nodes[path_key] = node
                
                # Add to parent or root
                if i == 0:
                    root_nodes.append(node)
                else:
                    parent_path = '.'.join(current_path[:-1])
                    if parent_path in nodes:
                        nodes[parent_path]['sections'].append(node)
    
    return root_nodes

def write_tree(tree, base_name):
    with open(os.path.join(training_dir, f'{base_name}_tree.json'), 'w') as f:
        json.dump(tree, f, indent=2)

def get_trees_from_index(training_dir):
    for filename in os.listdir(training_dir):
        if filename.endswith('.json') and 'tree' not in filename:
            base_name = filename[:-5]  # Remove '.json'
            json_path = os.path.join(training_dir, f'{base_name}.json')
            json_dl = get_json_dl(json_path)
            tree = get_tree_from_index(json_dl)
            write_tree(tree, base_name)

def add_indices_from_ids(dl): # dict with label, shape_id, and maybe other attributes
    """
    Arrange a flat list of dicts into a flat list ordered by containment hierarchy.
    
    Args:
        dl: List of dicts with 'label', 'shape_id' (list), and other attributes
        
    Returns:
        Flat list of sections with 'label' and 'index' attributes, ordered by hierarchy
    """
    if not dl:
        return []
    
    # Convert shape_id to sets for easier comparison
    for item in dl:
        if isinstance(item.get('shape_id'), list):
            item['shape_id_set'] = set(item['shape_id'])
        else:
            item['shape_id_set'] = set([item.get('shape_id', '')])
    
    # Build containment hierarchy
    def assign_indices(nodes, parent_index="", level=0):
        result = []
        # Sort nodes by size (larger sets first) to ensure proper hierarchy
        sorted_nodes = sorted(nodes, key=lambda x: len(x['shape_id_set']), reverse=True)
        
        for i, node in enumerate(sorted_nodes):
            # Create current index
            current_index = f"{parent_index}.{i+1:02d}" if parent_index else f"{i+1:02d}"
            
            # Find children (nodes that are contained in this node)
            children = []
            remaining_nodes = []
            
            for other_node in nodes:
                if (other_node != node and 
                    node['shape_id_set'].issuperset(other_node['shape_id_set']) and 
                    node['shape_id_set'] != other_node['shape_id_set']):
                    children.append(other_node)
                elif other_node != node:
                    remaining_nodes.append(other_node)
            
            # Create section with index
            section = {
                'label': node['label'],
                'index': current_index
            }
            
            # Copy other attributes from original node
            for key, value in node.items():
                if key not in ['label', 'shape_id_set']:
                    section[key] = value
            
            result.append(section)
            
            # Recursively process children
            if children:
                result.extend(assign_indices(children, current_index, level + 1))
        
        return result
    
    # Start with all nodes and build the hierarchy
    return assign_indices(dl)

if __name__ == "__main__":
    get_trees_from_index(training_dir)