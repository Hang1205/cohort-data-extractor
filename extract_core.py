"""Offline, streaming id-subset extraction. No source records leave the computer."""
from __future__ import annotations
from collections import Counter
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
import csv
import fnmatch
import gzip
import json
from itertools import islice
from pathlib import Path
import re
import sys
import threading
import time
import zipfile
from uuid import uuid4

VERSION = '1.0.0'
BASE = Path(__file__).resolve().parent
APP_HOME = Path(sys.executable).resolve().parent if getattr(sys,'frozen',False) else BASE
if (BASE / 'vendor').is_dir():
    sys.path.insert(0, str(BASE / 'vendor'))
elif (BASE / 'vendor.zip').is_file():
    sys.path.insert(0,str(BASE/'vendor.zip'))
# Source notes can exceed csv's 128 KiB default. Memory use is bounded by a
# record plus the selected ID set and coverage counts, not the full export.
csv.field_size_limit(512 * 1024 * 1024)
DELIMITERS = {'Auto':None, 'Comma':',', 'Tab':'\t', 'Semicolon':';', 'Pipe':'|'}
ID_MODES=('Exact text','Ignore case','Remove integer .0 suffix','Ignore case + remove integer .0 suffix')


def match_id(value, mode='Exact text'):
    if mode not in ID_MODES:raise ExtractionError('Choose a supported ID matching mode.')
    value=str(value).strip()
    if 'remove integer .0 suffix' in mode.casefold():
        match=re.fullmatch(r'([+-]?\d+)\.0+',value)
        if match:value=match.group(1)
    if 'ignore case' in mode.casefold():value=value.casefold()
    return value


class ExtractionError(Exception):
    pass


class Cancelled(ExtractionError):
    pass


@dataclass
class Selection:
    ids: set[str]
    records: int
    duplicates: int
    blanks: int
    column: str
    sheet: str
    header_row: int = 1
    filtered_out: int = 0


def open_text(path, encoding='utf-8-sig', mode='rt'):
    return (gzip.open if str(path).lower().endswith('.gz') else open)(path, mode, encoding=encoding, newline='')


def supported(path):
    return str(path).lower().endswith(('.csv', '.tsv', '.csv.gz', '.tsv.gz'))


def delimiter_for(handle, path, choice='Auto'):
    if choice not in DELIMITERS:
        raise ExtractionError('Choose a supported delimiter setting.')
    if DELIMITERS[choice] is not None:
        return DELIMITERS[choice]
    if str(path).lower().endswith(('.tsv', '.tsv.gz')):
        return '\t'
    # Sniff the header only; notes may themselves contain tabs/commas/newlines.
    position = handle.tell()
    header = handle.readline(1024 * 1024)
    handle.seek(position)
    if not header:
        raise ExtractionError('A table is empty and has no header.')
    try:
        return csv.Sniffer().sniff(header, delimiters=',\t;|').delimiter
    except csv.Error:
        return ','  # A legitimate single-column id list needs no delimiter.


def validate_headers(headers):
    normalized = [str(h).strip().casefold() for h in headers]
    if not headers or any(not h for h in normalized) or len(set(normalized)) != len(normalized):
        raise ExtractionError('Headers must be nonempty and unique, ignoring case and surrounding spaces.')


def column_index(headers, column, strict=True):
    if not column.strip():raise ExtractionError('Enter a nonempty id ID / cohort column name.')
    if strict:validate_headers(headers)
    matches = [i for i,h in enumerate(headers) if str(h).strip().casefold() == column.strip().casefold()]
    if len(matches) != 1:
        raise ExtractionError('The configured id ID column is missing or ambiguous.')
    return matches[0]


def excel_library():
    try:
        import openpyxl
        return openpyxl
    except ImportError:
        raise ExtractionError('Excel support is missing. Extract the full package including vendor/, or save the id list as CSV.') from None


def open_excel(path):
    try:
        return excel_library().load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except FileNotFoundError:
        raise ExtractionError('ID-list file not found. Confirm the file exists and the mapped drive is connected, then choose the file again.') from None
    except PermissionError:
        raise ExtractionError('Cannot access the id-list file. Close it in Excel or copy it to an approved local folder and try again.') from None
    except (zipfile.BadZipFile,KeyError,ValueError):
        raise ExtractionError('This file cannot be opened as an .xlsx workbook. It may be damaged, encrypted, or renamed from another format. In Excel, save a new unencrypted .xlsx copy, or export the id-ID sheet as CSV UTF-8.') from None
    except OSError:
        raise ExtractionError('Cannot read the workbook from this location. Check drive access and try an approved local copy.') from None


def excel_sheets(path):
    workbook=open_excel(path)
    try:return workbook.sheetnames
    finally:workbook.close()


def header_number(value):
    if str(value).strip().casefold()=='auto':return None
    try:
        number=int(value)
        if number<1 or number>100000:raise ValueError
        return number
    except (ValueError,TypeError):
        raise ExtractionError('Header row must be Auto or a positive row number (1–100000).') from None


def find_header(rows, header_row, column, relaxed):
    number=header_number(header_row)
    for row_number,raw in enumerate(rows,1):
        if number is None and row_number>50:break
        if number is not None and row_number<number:continue
        headers=[str(v) if v is not None else '' for v in raw]
        if number is None:
            # Only an exact, unique ID-column heading establishes an automatic header.
            if sum(h.strip().casefold()==column.strip().casefold() for h in headers)!=1:continue
        if relaxed:
            if not any(h.strip() for h in headers):
                raise ExtractionError('The chosen header row is empty. Select another worksheet or choose Auto / the correct header row.')
        else:validate_headers(headers)
        return headers,row_number
    raise ExtractionError('Cannot locate the header row. Choose the worksheet containing the id IDs and set the correct header row (Auto searches the first 50 rows).')


def delimiter_for_header(handle,path,choice,header_row,column):
    if choice!='Auto' or header_number(header_row)==1:return delimiter_for(handle,path,choice)
    position=handle.tell()
    try:
        for delim in (',','\t',';','|'):
            handle.seek(position)
            try:
                headers,_=find_header(csv.reader(handle,delimiter=delim,strict=True),header_row,column,True)
                column_index(headers,column,strict=False)
                return delim
            except (ExtractionError,csv.Error):continue
        handle.seek(position)
        return delimiter_for(handle,path,choice)
    finally:handle.seek(position)


def selection_layout(path, sheet='', encoding='utf-8-sig', delimiter='Auto', id_list=False, header_row=1, column='record_id'):
    path = Path(path)
    if path.suffix.lower() == '.xlsx':
        workbook = open_excel(path)
        try:
            sheets = workbook.sheetnames
            if sheet and sheet not in sheets:
                raise ExtractionError('The selected Excel sheet does not exist.')
            candidates=[sheet] if sheet else (sheets if header_number(header_row) is None else sheets[:1])
            last_error=None
            for candidate in candidates:
                ws=workbook[candidate];ws.reset_dimensions()
                try:
                    headers,number=find_header(ws.iter_rows(values_only=True),header_row,column,id_list)
                    while headers and not headers[-1].strip():headers.pop()
                    return headers,(sheets if sheet else [candidate]+[s for s in sheets if s!=candidate]),number
                except ExtractionError as error:last_error=error
            raise last_error or ExtractionError('The workbook contains no worksheets.')
        finally:
            workbook.close()
    if not supported(path):
        raise ExtractionError('Use a CSV/TSV id list or .xlsx Excel workbook. Save legacy .xls as .xlsx or CSV.')
    with open_text(path, encoding) as handle:
        reader = csv.reader(handle, delimiter=delimiter_for_header(handle,path,delimiter,header_row,column), strict=True)
        headers,number=find_header(reader,header_row,column,id_list)
        return headers, [],number


def inspect_selection(path, sheet='', encoding='utf-8-sig', delimiter='Auto', id_list=False, header_row=1, column='record_id'):
    headers,sheets,_=selection_layout(path,sheet,encoding,delimiter,id_list,header_row,column)
    return headers,sheets


def excel_id(cell, row):
    value = cell.value
    if value is None:
        return ''
    if cell.data_type in ('f','e'):
        raise ExtractionError(f'ID-list row {row} contains a formula or Excel error in the ID column. Use literal text IDs.')
    if isinstance(value, bool):
        raise ExtractionError(f'ID-list row {row} contains a boolean ID.')
    if isinstance(value, (int,float)):
        if not isinstance(value,int) and not value.is_integer():
            raise ExtractionError(f'ID-list row {row} contains a non-integer numeric ID. Use text IDs.')
        integer = str(int(value))
        if len(integer.lstrip('-')) > 15:
            raise ExtractionError(f'ID-list row {row} has a numeric ID longer than Excel\'s 15-digit precision. Recover the original ID as text.')
        if int(value) >= 0 and re.fullmatch(r'0+',cell.number_format or ''):
            return integer.zfill(len(cell.number_format))
        return integer
    if not isinstance(value,str):
        raise ExtractionError(f'ID-list row {row} has an unsupported ID type. Use text IDs.')
    return value.strip()


def load_selection(path, column='record_id', sheet='', encoding='utf-8-sig', delimiter='Auto', expected=None,
                   header_row=1, filter_column='', filter_values=''):
    path = Path(path).resolve()
    if not path.is_file():
        raise ExtractionError('The selected id-list file does not exist.')
    headers, sheets,actual_header = selection_layout(path,sheet,encoding,delimiter,True,header_row,column)
    index = column_index(headers,column,strict=False)
    filter_index=column_index(headers,filter_column,strict=False) if filter_column.strip() else None
    allowed={v.strip() for v in filter_values.split(',') if v.strip()}
    if filter_index is not None and not allowed:raise ExtractionError('Enter one or more cohort-label values, separated by commas, or clear the filter column.')
    filtered_out=0
    ids, total, duplicates, blanks = set(), 0, 0, 0
    def add(value):
        nonlocal total,duplicates,blanks
        total += 1
        if not value:
            blanks += 1
        elif value in ids:
            duplicates += 1
        else:
            ids.add(value)
    if path.suffix.lower() == '.xlsx':
        workbook = open_excel(path)
        try:
            actual_sheet = sheet or sheets[0]
            ws = workbook[actual_sheet]
            ws.reset_dimensions()
            for number,row in enumerate(ws.iter_rows(min_row=actual_header+1),actual_header+1):
                if filter_index is not None:
                    cell=row[filter_index] if len(row)>filter_index else None
                    if cell is not None and cell.data_type in ('f','e'):
                        raise ExtractionError(f'Cohort-label row {number} contains a formula or Excel error. Use literal label values.')
                    value='' if cell is None or cell.value is None else str(cell.value).strip()
                    if value not in allowed:filtered_out+=1;continue
                add(excel_id(row[index],number) if len(row)>index else '')
        finally:
            workbook.close()
    else:
        actual_sheet = ''
        with open_text(path,encoding) as handle:
            reader = csv.reader(handle,delimiter=delimiter_for_header(handle,path,delimiter,actual_header,column),strict=True)
            for _ in range(actual_header):next(reader)
            for number,row in enumerate(reader,actual_header+1):
                if not row:
                    continue
                if len(row) != len(headers):
                    raise ExtractionError(f'ID-list record {number} has the wrong number of columns.')
                if filter_index is not None and row[filter_index].strip() not in allowed:filtered_out+=1;continue
                add(row[index].strip())
    if not ids:
        raise ExtractionError('The id list contains no nonempty IDs.')
    if expected is not None and len(ids) != expected:
        raise ExtractionError(f'The list has {len(ids):,} unique IDs; expected {expected:,}. Verify your list or update the expected count.')
    return Selection(ids,total,duplicates,blanks,headers[index],actual_sheet,actual_header,filtered_out)


def discover_tables(folder, selection_path=None, recursive=False):
    folder = Path(folder).resolve()
    if not folder.is_dir():
        raise ExtractionError('Choose the folder containing the source data exports.')
    selection_path = Path(selection_path).resolve() if selection_path else None
    # Explicit recursion avoids directory symlinks/junctions and cloud shortcuts.
    paths, ignored = [], []
    def visit(current):
        for path in sorted(current.iterdir()):
            is_junction = getattr(path,'is_junction',lambda:False)()
            if path.is_symlink() or is_junction:
                ignored.append((path.relative_to(folder).as_posix(),'link not followed'))
            elif path.is_dir():
                if recursive:
                    visit(path)
            elif path.is_file():
                if path.resolve() == selection_path:
                    ignored.append((path.relative_to(folder).as_posix(),'selected-id list'))
                elif supported(path) or path.suffix.lower()=='.xlsx':
                    paths.append(path)
                else:
                    ignored.append((path.relative_to(folder).as_posix(),'unsupported file type'))
    visit(folder)
    if not paths:
        raise ExtractionError('No CSV/TSV or .xlsx cohort files were found in the selected folder.')
    return paths,ignored


def file_chosen(relative, include_patterns='*', exclude_patterns=''):
    include=[p.strip().casefold() for p in include_patterns.split(';') if p.strip()]
    exclude=[p.strip().casefold() for p in exclude_patterns.split(';') if p.strip()]
    def matches(pattern):return fnmatch.fnmatchcase(relative.casefold(),pattern) or fnmatch.fnmatchcase(Path(relative).name.casefold(),pattern)
    return any(matches(p) for p in (include or ['*'])) and not any(matches(p) for p in exclude)


def file_settings(relative, file_options, id_column, header_row, encoding, delimiter):
    config={'include':True,'column':id_column,'header_row':header_row,'sheet':'','encoding':encoding,'delimiter':delimiter,'id_index':None}
    config.update((file_options or {}).get(relative,{}))
    header_number(config['header_row'])
    return config


def configured_index(headers, config, strict_headers=False):
    if config.get('id_index') is None:return column_index(headers,config['column'],strict_headers)
    if strict_headers:validate_headers(headers)
    try:index=int(config['id_index'])
    except (ValueError,TypeError):raise ExtractionError('The configured ID column position is invalid.') from None
    if index<0 or index>=len(headers):raise ExtractionError('The configured ID column position is outside this file.')
    if str(headers[index]).strip().casefold()!=config['column'].strip().casefold():
        raise ExtractionError('The selected ID column heading changed. Read this file’s headers and select its ID column again.')
    return index


@contextmanager
def source_rows(path, config, strict_headers=False):
    """Stream original values; XLSX exports become CSV with literal cell values/formulas."""
    headers,sheets,number=selection_layout(path,config['sheet'],config['encoding'],config['delimiter'],not strict_headers,config['header_row'],config['column'])
    index=configured_index(headers,config,strict_headers)
    if path.suffix.lower()=='.xlsx':
        book=open_excel(path)
        try:
            actual_sheet=config['sheet'] or sheets[0];ws=book[actual_sheet];ws.reset_dimensions()
            def records():
                for rownum,cells in enumerate(ws.iter_rows(min_row=number+1),number+1):
                    if len(cells)>len(headers) and any(c.value is not None for c in cells[len(headers):]):
                        raise ExtractionError(f'Excel data row {rownum} has values beyond the header columns. Choose the correct header row.')
                    values=[]
                    for i in range(len(headers)):
                        cell=cells[i] if i<len(cells) else None
                        value=None if cell is None else cell.value
                        values.append('' if value is None else (value.isoformat() if hasattr(value,'isoformat') else str(value)))
                    id=excel_id(cells[index],rownum) if len(cells)>index else ''
                    yield values,id
            yield headers,index,records(),',',actual_sheet,number
        finally:book.close()
    else:
        with open_text(path,config['encoding']) as handle:
            delim=delimiter_for_header(handle,path,config['delimiter'],number,config['column'])
            reader=csv.reader(handle,delimiter=delim,strict=True)
            for _ in range(number):next(reader)
            yield headers,index,((row,row[index] if len(row)>index else '') for row in reader),delim,'',number


def inspect_tables(folder, selection_path=None, id_column='record_id', encoding='utf-8-sig', delimiter='Auto', recursive=False,
                   header_row=1, strict_headers=False, include_patterns='*', exclude_patterns='', file_options=None):
    files,ignored = discover_tables(folder,selection_path,recursive)
    results = []
    for path in files:
        result = {'file':path.relative_to(Path(folder).resolve()).as_posix(), 'bytes':path.stat().st_size}
        config=file_settings(result['file'],file_options,id_column,header_row,encoding,delimiter)
        result['config']=config
        if not config['include'] or not file_chosen(result['file'],include_patterns,exclude_patterns):
            result.update(status='excluded by file choices',columns=0);results.append(result);continue
        try:
            headers,sheets,number = selection_layout(path,config['sheet'],config['encoding'],config['delimiter'],not strict_headers,config['header_row'],config['column'])
            result.update(headers=headers,sheets=sheets)
            configured_index(headers,config,strict_headers)
            result.update(status='ready',columns=len(headers))
            result['header_row']=number
        except (ExtractionError,UnicodeError,csv.Error,OSError):
            result.update(status='check ID column / header / encoding',columns=0)
        results.append(result)
    return results,ignored


def sample_matches(folder, selection, *, file_options=None, id_column='record_id', header_row=1, encoding='utf-8-sig',
                   delimiter='Auto', recursive=False, selection_path=None, include_patterns='*', exclude_patterns='', id_mode='Exact text', limit=10000):
    if not isinstance(limit,int) or limit<1:raise ExtractionError('Sample size must be a positive integer.')
    results,ignored=inspect_tables(folder,selection_path,id_column,encoding,delimiter,recursive,header_row,False,include_patterns,exclude_patterns,file_options)
    memberships={mode:{match_id(id,mode) for id in selection.ids} for mode in ID_MODES}
    for result in results:
        if result['status']!='ready':continue
        counts={mode:0 for mode in ID_MODES};sampled=0
        try:
            with source_rows(Path(folder)/result['file'],result['config']) as (headers,index,records,_,_,_):
                for row,id in islice(records,limit):
                    if not row:continue
                    if len(row)!=len(headers):raise ExtractionError('Malformed sampled record; check the file delimiter and header row.')
                    sampled+=1
                    for mode in ID_MODES:counts[mode]+=int(bool(id.strip()) and match_id(id,mode) in memberships[mode])
            result.update(sampled_rows=sampled,sample_matches=counts[id_mode],mode_counts=counts,status=f'sample: {counts[id_mode]}/{sampled} matching rows')
        except (ExtractionError,OSError,UnicodeError,csv.Error):result.update(status='sample failed — check file settings')
    return results,ignored


def extract(folder, selection_path, output, *, selection_column='record_id', id_column='record_id',
            sheet='', encoding='utf-8-sig', selection_encoding=None, delimiter='Auto', selection_delimiter='Auto',
            recursive=False, expected=None, cancel=None, progress=None, selection_header_row=1, source_header_row=1,
            filter_column='', filter_values='', strict_headers=False, include_patterns='*', exclude_patterns='', skip_invalid_files=False,
            file_options=None, id_mode='Exact text'):
    cancel = cancel or threading.Event()
    source = Path(folder).resolve()
    selection_path = Path(selection_path).resolve()
    output = Path(output).resolve()
    if output.exists():
        raise ExtractionError('Choose a NEW output folder. Existing output is never overwritten.')
    if output == source or source in output.parents or output in source.parents:
        raise ExtractionError('Choose an output folder outside the source source folder, without containing the source.')
    if output == selection_path or output in selection_path.parents:
        raise ExtractionError('The output folder cannot contain the selected-id list.')
    if cancel.is_set():
        raise Cancelled('Cancelled before extraction.')
    list_stamp = (selection_path.stat().st_size,selection_path.stat().st_mtime_ns)
    selection = load_selection(selection_path,selection_column,sheet,selection_encoding or encoding,selection_delimiter,expected,
                               selection_header_row,filter_column,filter_values)
    selected_lookup={match_id(id,id_mode):id for id in selection.ids}
    if len(selected_lookup)!=len(selection.ids):
        raise ExtractionError('This matching mode merges different selected id IDs. Use Exact text or correct the id list; no output was created.')
    diagnostic_sets={mode:{match_id(id,mode) for id in selection.ids} for mode in ID_MODES[:3]}
    all_files,ignored = discover_tables(source,selection_path,recursive)
    files=[];configs={}
    for path in all_files:
        relative=path.relative_to(source).as_posix()
        configs[path]=file_settings(relative,file_options,id_column,source_header_row,encoding,delimiter)
        if configs[path]['include'] and file_chosen(relative,include_patterns,exclude_patterns):files.append(path)
        else:ignored.append((relative,'excluded by file choices'))
    stamps = {p:(p.stat().st_size,p.stat().st_mtime_ns) for p in files}
    ready=[];layouts={}
    # Check every header before creating any partial output.
    for path in files:
        if cancel.is_set():
            raise Cancelled('Cancelled before extraction.')
        try:
            cfg=configs[path]
            headers,sheets,number = selection_layout(path,cfg['sheet'],cfg['encoding'],cfg['delimiter'],not strict_headers,cfg['header_row'],cfg['column'])
            configured_index(headers,cfg,strict_headers)
            if path.suffix.lower()=='.xlsx':cfg['sheet']=cfg['sheet'] or sheets[0]
            layouts[path]=number;ready.append(path)
        except (ExtractionError,csv.Error,UnicodeError,OSError):
            if skip_invalid_files:ignored.append((path.relative_to(source).as_posix(),'explicitly skipped: unreadable header / missing or ambiguous ID / encoding'))
            else:raise ExtractionError(f'Cannot read the configured ID column in {path.name}. Check its header row, delimiter and encoding. Exclude that file explicitly or use Skip files with invalid headers; no output was created.') from None
    files=ready
    if not files:raise ExtractionError('No readable source files remain after your file choices and header checks. No output was created.')
    destinations=[(p.relative_to(source).with_suffix('.csv') if p.suffix.lower()=='.xlsx' else p.relative_to(source)).as_posix().casefold() for p in files]
    if len(set(destinations))!=len(destinations):raise ExtractionError('Selected files would produce the same output filename (for example table.csv and table.xlsx). Exclude one or rename a source copy.')
    stamps={p:stamps[p] for p in files}
    output.parent.mkdir(parents=True,exist_ok=True)
    partial = output.with_name(output.name+'.incomplete-'+uuid4().hex[:8])
    partial.mkdir()
    (partial/'data').mkdir()
    started = time.monotonic()
    audit = {'app_version':VERSION,'started_utc':datetime.now(timezone.utc).isoformat(),'status':'running',
             'method':'Exact string ID membership; surrounding ID whitespace trimmed for matching only',
             'data_id_column':id_column,'selected_list_column':selection.column,
             'selected_unique_ids':len(selection.ids),'selected_list_records':selection.records,
             'duplicate_selected_ids':selection.duplicates,'blank_selected_ids':selection.blanks,
             'encoding':encoding,'expected_selected_count':expected,'files':[],
             'selected_sheet':selection.sheet,'selection_header_row':selection.header_row,'selection_rows_filtered_out':selection.filtered_out,
             'filter_column':filter_column,'filter_values':filter_values,'strict_headers':strict_headers,'source_header_row':source_header_row,
             'include_patterns':include_patterns,'exclude_patterns':exclude_patterns,'skip_invalid_files':skip_invalid_files,
             'id_matching_mode':id_mode,'per_file_options':{p.relative_to(source).as_posix():configs[p] for p in files},
             'ignored_files':[{'file':name,'reason':reason} for name,reason in ignored],
             'outputs':'Original columns and matching row values retained; no de-identification performed.'}
    coverage = Counter()
    try:
        (partial/'audit.json').write_text(json.dumps(audit,indent=2),encoding='utf-8')
        for file_index,path in enumerate(files,1):
            if cancel.is_set():
                raise Cancelled('Cancelled. The incomplete folder is not a completed extraction.')
            relative = path.relative_to(source)
            destination = partial/'data'/(relative.with_suffix('.csv') if path.suffix.lower()=='.xlsx' else relative)
            destination.parent.mkdir(parents=True,exist_ok=True)
            scanned,matched,blank_ids,blank_records = 0,0,0,0
            file_ids = set()
            last_update = 0.0
            report = {'file':relative.as_posix(),'input_bytes':stamps[path][0],'rows_scanned':0,'rows_selected':0,'ids_found':0,'blank_id_rows':0,'blank_records':0,'status':'running'}
            cfg=configs[path]
            report.update(id_column=cfg['column'],id_column_position=cfg.get('id_index'),sheet=cfg['sheet'],header_row=layouts[path],output_file=destination.relative_to(partial).as_posix(),
                          exact_id_rows=0,case_insensitive_id_rows=0,integer_suffix_id_rows=0)
            audit['files'].append(report)
            with source_rows(path,cfg,strict_headers) as (headers,index,reader,delim,_,_), open_text(destination,'utf-8-sig','wt') as target:
                writer = csv.writer(target,delimiter=delim,lineterminator='\r\n')
                writer.writerow(headers)
                for row,raw_id in reader:
                    if cancel.is_set():
                        raise Cancelled('Cancelled. The incomplete folder is not a completed extraction.')
                    if not row:
                        blank_records += 1
                        continue
                    scanned += 1
                    if len(row) != len(headers):
                        raise ExtractionError(f'A malformed record was found in {path.name}, data record {scanned}. No completed output was published.')
                    id = match_id(raw_id,id_mode)
                    blank_ids += int(not id)
                    for mode,key in zip(ID_MODES[:3],('exact_id_rows','case_insensitive_id_rows','integer_suffix_id_rows')):
                        report[key]+=int(bool(raw_id.strip()) and match_id(raw_id,mode) in diagnostic_sets[mode])
                    if id in selected_lookup:
                        writer.writerow(row)
                        matched += 1
                        original_id=selected_lookup[id]
                        file_ids.add(original_id)
                        coverage[original_id] += 1
                    now = time.monotonic()
                    if now-last_update >= 0.3:
                        report.update(rows_scanned=scanned,rows_selected=matched,ids_found=len(file_ids),blank_id_rows=blank_ids,blank_records=blank_records)
                        if progress:
                            progress({'file':relative.as_posix(),'file_index':file_index,'file_count':len(files),'rows_scanned':scanned,'rows_selected':matched,'elapsed_seconds':round(now-started,1)})
                        last_update = now
            report.update(rows_scanned=scanned,rows_selected=matched,ids_found=len(file_ids),blank_id_rows=blank_ids,blank_records=blank_records,status='complete',output_bytes=destination.stat().st_size)
            if (path.stat().st_size,path.stat().st_mtime_ns) != stamps[path]:
                raise ExtractionError('A source file changed during extraction. Retry with stable input files.')
            if progress:
                progress({'file':relative.as_posix(),'file_index':file_index,'file_count':len(files),'rows_scanned':scanned,'rows_selected':matched,'file_finished':True,'elapsed_seconds':round(time.monotonic()-started,1)})
        if cancel.is_set():
            raise Cancelled('Cancelled. The incomplete folder is not a completed extraction.')
        if any((p.stat().st_size,p.stat().st_mtime_ns)!=stamp for p,stamp in stamps.items()) or list_stamp != (selection_path.stat().st_size,selection_path.stat().st_mtime_ns):
            raise ExtractionError('Input files or the selected-id list changed during extraction.')
        current_files,_ = discover_tables(source,selection_path,recursive)
        if current_files != all_files:
            raise ExtractionError('The source folder gained or lost a source table during extraction. Retry with a stable folder.')
        # The explicit selected list and missing IDs remain local with the extracted source data.
        with (partial/'id_coverage.csv').open('w',encoding='utf-8-sig',newline='') as handle:
            writer = csv.writer(handle)
            writer.writerow(['record_id','total_selected_rows','found_in_any_file'])
            for id in sorted(selection.ids):
                writer.writerow([id,coverage[id],int(id in coverage)])
        with (partial/'missing_ids.csv').open('w',encoding='utf-8-sig',newline='') as handle:
            writer = csv.writer(handle)
            writer.writerow(['record_id'])
            writer.writerows([id] for id in sorted(selection.ids-coverage.keys()))
        with (partial/'file_summary.csv').open('w',encoding='utf-8-sig',newline='') as handle:
            fields = ['file','rows_scanned','rows_selected','ids_found','blank_id_rows','blank_records','input_bytes','output_bytes','status']
            writer = csv.DictWriter(handle,fieldnames=fields,extrasaction='ignore')
            writer.writeheader()
            writer.writerows(audit['files'])
        with (partial/'matching_diagnostics.csv').open('w',encoding='utf-8-sig',newline='') as handle:
            fields=['file','id_column','sheet','header_row','exact_id_rows','case_insensitive_id_rows','integer_suffix_id_rows','rows_selected','ids_found']
            writer=csv.DictWriter(handle,fieldnames=fields,extrasaction='ignore');writer.writeheader();writer.writerows(audit['files'])
        audit.update(status='complete',finished_utc=datetime.now(timezone.utc).isoformat(),elapsed_seconds=round(time.monotonic()-started,2),
                     ids_found_anywhere=len(coverage),ids_missing_everywhere=len(selection.ids-coverage.keys()),
                     total_rows_scanned=sum(f['rows_scanned'] for f in audit['files']),total_rows_selected=sum(f['rows_selected'] for f in audit['files']))
        audit['zero_match_files']=[f['file'] for f in audit['files'] if f['rows_selected']==0]
        (partial/'audit.json').write_text(json.dumps(audit,indent=2),encoding='utf-8')
        if output.exists():
            raise ExtractionError('The output folder appeared during processing. Choose a new folder.')
        partial.rename(output)
        return audit
    except Exception as error:
        audit.update(status='cancelled' if isinstance(error,Cancelled) else 'failed',finished_utc=datetime.now(timezone.utc).isoformat())
        try:
            (partial/'audit.json').write_text(json.dumps(audit,indent=2),encoding='utf-8')
        except OSError:
            pass  # Preserve the original failure when storage is unavailable.
        if isinstance(error,ExtractionError):
            raise
        raise ExtractionError('Extraction failed. Check encoding, CSV formatting, permissions and available space. The incomplete folder must not be treated as completed output.') from None
