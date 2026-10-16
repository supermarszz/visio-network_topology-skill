"""Build an installable ZIP containing only this skill, not test/user artifacts."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED

root=Path(__file__).resolve().parents[1]
skill=root/'skills'/'visio-network-topology'
version=(root/'VERSION').read_text(encoding='utf-8').strip()
output=root/'dist'/f'visio-network-topology-{version}.zip'
output.parent.mkdir(exist_ok=True)
files=sorted(p for p in skill.rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.suffix not in ('.pyc','.pyo'))
with ZipFile(output,'w',ZIP_DEFLATED) as archive:
    for path in files:
        archive.write(path,Path(skill.name)/path.relative_to(skill))
print(output)
