from pathlib import Path
import shutil, zipfile, os
base = Path('output/soft_slides_app')
zip_path = Path('output/soft_slides_app_build.zip')
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
    for p in base.rglob('*'):
        if p.is_file():
            z.write(p, arcname=str(p.relative_to(base.parent)))
print(zip_path)