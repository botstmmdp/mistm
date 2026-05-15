import os
import glob

html_files = glob.glob("*.html")
target_str = '{ t: "Desempeño Eficaz"'
new_str = '{ t: "Violencia de Género", d: "Encuentros Virtuales - Secretaría de Género", f: "capacitacion.html", tags: "violencia genero mujer diversidad taller proteccion ayuda encuentros virtuales", ico: "fa-hand-holding-heart" }'

updated_count = 0

for file_path in html_files:
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        with open(file_path, 'r', encoding='latin-1') as f:
            content = f.read()
            
    if target_str in content and new_str not in content:
        lines = content.split('\n')
        new_lines = []
        inserted = False
        for line in lines:
            new_lines.append(line)
            if target_str in line and not inserted:
                indent = line[:line.find('{')]
                new_lines.append(indent + new_str + ',')
                inserted = True
                
        if inserted:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(new_lines))
            print(f"Updated {file_path}")
            updated_count += 1

print(f"Total updated: {updated_count}")
