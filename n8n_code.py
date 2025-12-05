# In n8n, input data is accessed via _input.all()
# Each item has a .json property containing the data

for item in _input.all():
    try:
        # Get the input JSON data
        data = item.json
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

        # Replace the item's JSON with the flattened data
        item.json = row_data
        
    except Exception as e:
        # Add error info to the item if something fails
        item.json['error'] = str(e)

# Return the modified items
return _input.all()
