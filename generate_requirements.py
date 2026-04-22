# generate_clean_requirements.py

import os
import ast
import pkg_resources
from pathlib import Path

project_path = Path(r"/")

# Step 1: Collect all top-level imports from Python files
imports = set()

for root, _, files in os.walk(project_path):
    for file in files:
        if file.endswith(".py"):
            with open(os.path.join(root, file), encoding="utf-8") as f:
                try:
                    node = ast.parse(f.read(), filename=file)
                    for n in ast.walk(node):
                        if isinstance(n, ast.Import):
                            for alias in n.names:
                                imports.add(alias.name.split('.')[0])
                        elif isinstance(n, ast.ImportFrom):
                            if n.module:
                                imports.add(n.module.split('.')[0])
                except Exception:
                    pass  # skip files that can't be parsed

# Step 2: Map imports to installed packages
installed = {pkg.key: pkg.version for pkg in pkg_resources.working_set}
requirements = []

for imp in sorted(imports):
    for pkg_key in installed:
        if imp.lower() == pkg_key.replace("-", "_"):
            requirements.append(f"{pkg_key}=={installed[pkg_key]}")
            break

# Step 3: Write to and print requirements.txt
output_file = project_path / "requirements.txt"
with open(output_file, "w", encoding="utf-8") as f:
    for req in sorted(set(requirements)):
        f.write(req + "\n")

print("\n===== requirements.txt CONTENT =====\n")
print(output_file.read_text(encoding="utf-8"))
