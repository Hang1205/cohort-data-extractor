"""Build complete source/desktop ZIPs from a controlled publication inventory."""
from pathlib import Path
import hashlib
import json
import shutil
import sys
import zipfile

ROOT=Path(__file__).resolve().parent
NAME='Cohort_Data_Extractor'
VERSION='1.1.0'
SOURCE_NAMES=['extract_core.py','cohort_extract_gui.py','extract_cli.py','desktop_app.py','desktop_runtime.py','test_extractor.py',
              'build_windows.py','pack_release.py','Find_Python312.cmd','Start_Windows.cmd','Start_Mac.command','LICENSE','README.md',
              'DESKTOP_QUICKSTART.md','CHANGELOG.md','CITATION.cff','THIRD_PARTY_NOTICES.md','BUILD_VERIFICATION.md',
              'requirements.txt','requirements_build_windows.lock.txt','file_options.example.json','.gitignore','vendor.zip','examples.zip']

def source_files():
    files=[ROOT/name for name in SOURCE_NAMES]
    for folder in ('vendor','examples','.github'):
        if (ROOT/folder).exists():files.extend(p for p in (ROOT/folder).rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.suffix!='.pyc')
    assert all(p.is_file() for p in files),'Required source file missing'
    return sorted(files)

def digest(path):
    with path.open('rb') as handle:return hashlib.file_digest(handle,'sha256').hexdigest()

def package(destination=None):
    destination=Path(destination or ROOT/'release_output').resolve();destination.mkdir(parents=True,exist_ok=True)
    src=source_files()
    source_zip=destination/f'{NAME}_Source_v{VERSION}.zip'
    with zipfile.ZipFile(source_zip,'w',zipfile.ZIP_DEFLATED) as z:
        for path in src:z.write(path,NAME+'/'+path.relative_to(ROOT).as_posix())
    desktop=ROOT/'build_output'/NAME
    assert (desktop/(NAME+'.exe')).is_file(),'Build the Windows EXE before packaging'
    for path in src:
        target=desktop/'source'/path.relative_to(ROOT);target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(path,target)
    shutil.copytree(ROOT/'examples',desktop/'examples',dirs_exist_ok=True)
    for name in ('LICENSE','DESKTOP_QUICKSTART.md','BUILD_VERIFICATION.md','THIRD_PARTY_NOTICES.md'):shutil.copyfile(ROOT/name,desktop/name)
    (desktop/'START_HERE.txt').write_text('Extract ALL files. Open Cohort_Data_Extractor.exe. Keep _internal beside the EXE. Complete source is in source/. See DESKTOP_QUICKSTART.md.\n',encoding='utf-8')
    (desktop/'Start_Desktop.cmd').write_text('@echo off\ncd /d "%~dp0"\nif not exist "_internal\\python312.dll" (\n echo Extract the entire ZIP and keep _internal beside the EXE.\n pause\n exit /b 1\n)\nstart "" "%~dp0Cohort_Data_Extractor.exe"\n',encoding='ascii')
    paths=sorted(p for p in desktop.rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.suffix!='.pyc' and p.name not in ('MANIFEST.json','Desktop_Self_Test.json') and 'Diagnostics' not in p.parts)
    (desktop/'MANIFEST.json').write_text(json.dumps({'application':NAME,'version':VERSION,'files':{p.relative_to(desktop).as_posix():digest(p) for p in paths}},indent=2),encoding='utf-8')
    desktop_zip=destination/f'{NAME}_Windows_Desktop_v{VERSION}.zip'
    with zipfile.ZipFile(desktop_zip,'w',zipfile.ZIP_DEFLATED) as z:
        for path in paths+[desktop/'MANIFEST.json']:z.write(path,NAME+'/'+path.relative_to(desktop).as_posix())
    for path in (source_zip,desktop_zip):
        with zipfile.ZipFile(path) as z:assert z.testzip() is None
    sums=destination/f'{NAME}_SHA256SUMS.txt';sums.write_text('\n'.join(digest(p)+'  '+p.name for p in (source_zip,desktop_zip))+'\n',encoding='ascii')
    print(json.dumps({'source':str(source_zip),'desktop':str(desktop_zip),'checksums':str(sums),'source_files':len(src)}))

if __name__=='__main__':package(sys.argv[1] if len(sys.argv)>1 else None)
