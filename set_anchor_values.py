def set_anchor_values(file_path, replacements, output_path=None):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()

        for old, new in replacements.items():
            content = content.replace(old, new)

        target_path = output_path if output_path else file_path
        with open(target_path, 'w', encoding='utf-8') as file:
            file.write(content)

        print(f"File updated successfully: {target_path}")
    except Exception as e:
        print(f"Error: {e}")

# Example usage
replacements = {
    "!currentstack!": "currentstack_new",
}

set_anchor_values("Configurator\sample.txt", replacements, "updated_file.txt")