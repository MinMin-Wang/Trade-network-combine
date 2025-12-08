import pandas as pd
import numpy as np

def read_data_from_excel(file_path, sheet_name):
    """
    Read data from specified Excel file and worksheet
    
    Args:
        file_path (str): Path to Excel file
        sheet_name (str): Worksheet name
    
    Returns:
        pd.DataFrame: DataFrame containing node relationship data
    """
    return pd.read_excel(file_path, sheet_name=sheet_name)

def generate_adjacency_matrix(data):
    """
    Generate adjacency matrix from DataFrame containing node relationships (with weights)
    
    Args:
        data (pd.DataFrame): DataFrame with 'Source', 'Target', and 'Weight' columns
    
    Returns:
        tuple: (adjacency_matrix, node_names) where:
            - adjacency_matrix: numpy array representing the adjacency matrix
            - node_names: list of node names
    """
    # Extract unique nodes from Source and Target columns
    all_nodes = np.unique(np.concatenate([data['Source'].values, data['Target'].values]))
    node_to_index = {node: index for index, node in enumerate(all_nodes)}
    num_nodes = len(all_nodes)
    
    # Initialize adjacency matrix with zeros
    adjacency_matrix = np.zeros((num_nodes, num_nodes))
    
    # Fill the adjacency matrix with weights
    for _, row in data.iterrows():
        source_index = node_to_index[row['Source']]
        target_index = node_to_index[row['Target']]
        weight = row['Weight']
        adjacency_matrix[source_index][target_index] = weight
        # For undirected graphs, uncomment the line below:
        # adjacency_matrix[target_index][source_index] = weight
    
    return adjacency_matrix, all_nodes

def matrix_to_dataframe(adjacency_matrix, node_names=None):
    """
    Convert adjacency matrix to DataFrame with optional node names as row/column indices
    
    Args:
        adjacency_matrix (np.array): Adjacency matrix as numpy array
        node_names (list): List of node names (optional)
    
    Returns:
        pd.DataFrame: DataFrame representation of the adjacency matrix
    """
    if node_names is not None:
        df = pd.DataFrame(adjacency_matrix, columns=node_names, index=node_names)
    else:
        df = pd.DataFrame(adjacency_matrix)
    return df

def write_to_excel(df, file_path):
    """
    Write DataFrame data to Excel file
    
    Args:
        df (pd.DataFrame): DataFrame containing adjacency matrix data
        file_path (str): Path to save Excel file
    """
    df.to_excel(file_path, index=True, header=True)

if __name__ == "__main__":
    # Configuration
    file_path = "2010colzamat.xlsx"  # Replace with actual Excel file path
    sheet_name = "Sheet1"
    output_file = "matrix_2010_colza.xlsx"  # Output file path
    
    # Read and process data
    data = read_data_from_excel(file_path, sheet_name)
    
    # Ensure node names are strings
    data['Source'] = data['Source'].astype(str)
    data['Target'] = data['Target'].astype(str)
    
    # Generate adjacency matrix
    adjacency_matrix, node_names = generate_adjacency_matrix(data)
    
    # Display information
    print("Node names:", node_names)
    print("Adjacency matrix shape:", adjacency_matrix.shape)
    
    # Convert to DataFrame and save
    df = matrix_to_dataframe(adjacency_matrix, node_names)
    write_to_excel(df, output_file)
    print(f"Adjacency matrix successfully written to {output_file}")
