from pathlib import Path
import os
import subprocess
import sys

root=Path(__file__).resolve().parent
(root/"work").mkdir(exist_ok=True)
env=os.environ.copy()
env['PYINSTALLER_CONFIG_DIR']=str(root/'work/pyinstaller-cache')
kind='extractor'
source=root
name='Cohort_Data_Extractor'
extra=['--collect-all','openpyxl','--collect-all','et_xmlfile','--exclude-module','numpy','--exclude-module','pandas','--exclude-module','PIL','--exclude-module','lxml']
dist_root=root/'build_output'
assert (dist_root/name).resolve().is_relative_to((root/'build_output').resolve())
version_file=root/'work'/f'{kind}_windows_version.txt'
product='Cohort Data Extractor'
version_file.write_text(f'''VSVersionInfo(
  ffi=FixedFileInfo(filevers=(1,1,0,0), prodvers=(1,1,0,0), mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0,0)),
  kids=[StringFileInfo([StringTable('040904B0',[
    StringStruct('CompanyName','Hang Xu'), StringStruct('FileDescription','{product} offline desktop GUI'),
    StringStruct('FileVersion','1.1.0'), StringStruct('InternalName','{name}'),
    StringStruct('LegalCopyright','Copyright (c) 2026 Hang Xu'), StringStruct('OriginalFilename','{name}.exe'),
    StringStruct('ProductName','{product}'), StringStruct('ProductVersion','1.1.0')])]),
    VarFileInfo([VarStruct('Translation',[1033,1200])])])
''',encoding='utf-8')
command=[sys.executable,'-m','PyInstaller','--noconfirm','--onedir','--windowed','--noupx','--name',name,
         '--version-file',str(version_file),
         '--distpath',str(root/'build_output'),'--workpath',str(root/'work/desktop-build'/kind),
         '--specpath',str(root/'work/desktop-specs'),'--paths',str(source),*extra,str(source/'desktop_app.py')]
log=root/'work'/f'{kind}_desktop_build.log'
with log.open('w',encoding='utf-8') as handle:
    result=subprocess.run(command,env=env,stdout=handle,stderr=subprocess.STDOUT)
print(f'{kind} desktop build exited {result.returncode}. Log: {log}')
raise SystemExit(result.returncode)
