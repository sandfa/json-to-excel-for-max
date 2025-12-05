from converter import convert_json_to_excel
import os

if __name__ == "__main__":
    input_dir = "data"
    output_file = "output/output.xlsx"
    
    # Ensure directories exist
    os.makedirs(input_dir, exist_ok=True)
    if os.path.dirname(output_file):
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    print(f"Starting conversion from '{input_dir}' to '{output_file}'...")
    convert_json_to_excel(input_dir, output_file)
    print("Conversion finished.")
