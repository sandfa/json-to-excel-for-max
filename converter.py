import os
import json
import pandas as pd
import glob

def convert_json_to_excel(input_dir, output_file):
    """
    Reads all JSON files from input_dir and converts them to a single Excel file
    matching the structure of the existing output.xlsx (if available) or a target schema.
    """
    all_data = []
    
    # Find all JSON files in the directory
    json_files = glob.glob(os.path.join(input_dir, "*.json"))
    
    if not json_files:
        print("No JSON files found in the directory.")
        return

    # Try to read the target columns from output.xlsx if it exists
    target_columns = []
    template_path = "output.xlsx"
    if os.path.exists(template_path):
        try:
            template_df = pd.read_excel(template_path)
            target_columns = list(template_df.columns)
            print(f"Loaded {len(target_columns)} columns from {template_path}")
        except Exception as e:
            print(f"Could not read template {template_path}: {e}")
    
    for file_path in json_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                row_data = {}
                
                # 1. Map Basisinformationen
                if "Basisinformationen" in data:
                    basis = data["Basisinformationen"]
                    for key, value in basis.items():
                        # Normalize keys to match Excel (e.g. "&" -> "und")
                        normalized_key = key.replace("&", "und")
                        # Also check for specific mappings if needed
                        if key == "Kurzbeschreibung":
                            normalized_key = "Kurzbeschreibung des Projekts"
                        
                        row_data[normalized_key] = value
                        # Also try exact match if normalization didn't hit
                        row_data[key] = value

                # 2. Map Aufgabenplan
                if "Aufgabenplan" in data:
                    tasks = data["Aufgabenplan"]
                    for i, task in enumerate(tasks):
                        # Main Task ID: 10, 20, 30...
                        task_id_prefix = (i + 1) * 10
                        
                        # Map Main Task Title
                        # Column: A10 Titel
                        row_data[f"A{task_id_prefix} Titel"] = task.get("Aufgabe", "")
                        
                        # Map Subtasks
                        subtasks = task.get("Unteraufgaben", [])
                        for j, subtask in enumerate(subtasks):
                            # Subtask ID: 11, 12... 21, 22...
                            subtask_id = task_id_prefix + j + 1
                            
                            row_data[f"A{subtask_id} Titel"] = subtask.get("Titel", "")
                            row_data[f"A{subtask_id} Linked"] = subtask.get("Verknüpfung", "")
                            row_data[f"A{subtask_id} Dauer"] = subtask.get("Dauer_Tage", "")
                            row_data[f"A{subtask_id} Details"] = subtask.get("Details", "")

                all_data.append(row_data)
                
        except Exception as e:
            print(f"Error reading {file_path}: {e}")

    if all_data:
        df = pd.DataFrame(all_data)
        
        # If we have target columns, ensure we match them
        if target_columns:
            # Reindex to match target columns, adding missing ones as empty
            df = df.reindex(columns=target_columns)
        
        df.to_excel(output_file, index=False)
        print(f"Successfully created {output_file}")
    else:
        print("No data to write.")
