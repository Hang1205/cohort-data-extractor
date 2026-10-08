"""Command-line interface to the same streaming extraction engine as the GUI."""
import argparse
import json
import sys
from extract_core import ExtractionError,extract,ID_MODES


def main():
    parser=argparse.ArgumentParser(description='Extract selected ids from every source CSV/TSV export, offline.')
    parser.add_argument('--source-file',action='append',default=[],help='Additional file from any folder; repeat for multiple files')
    parser.add_argument('--source',required=True,help='source export folder')
    parser.add_argument('--ids',required=True,help='CSV/TSV or XLSX selected-id list')
    parser.add_argument('--output',required=True,help='New output folder outside the source')
    parser.add_argument('--id-column',default='record_id')
    parser.add_argument('--selected-id-column',default='record_id')
    parser.add_argument('--sheet',default='')
    parser.add_argument('--expected',type=int,default=0,help='Expected unique ids; 0 disables count check')
    parser.add_argument('--encoding',default='utf-8-sig')
    parser.add_argument('--selection-encoding',default='utf-8-sig')
    parser.add_argument('--delimiter',choices=['Auto','Comma','Tab','Semicolon','Pipe'],default='Auto')
    parser.add_argument('--selection-delimiter',choices=['Auto','Comma','Tab','Semicolon','Pipe'],default='Auto')
    parser.add_argument('--recursive',action='store_true')
    parser.add_argument('--selection-header-row',default='Auto',help='Auto (first 50 rows) or row number')
    parser.add_argument('--source-header-row',default='1',help='Auto or row number')
    parser.add_argument('--filter-column',default='',help='Optional cohort/cohort label column in id list')
    parser.add_argument('--filter-values',default='',help='Explicit exact label values, comma separated')
    parser.add_argument('--include',default='*',help='Include filename patterns, separated by semicolons')
    parser.add_argument('--exclude',default='',help='Exclude filename patterns, separated by semicolons')
    parser.add_argument('--strict-headers',action='store_true')
    parser.add_argument('--skip-invalid-files',action='store_true',help='Explicitly skip files whose headers cannot be used; audit lists them')
    parser.add_argument('--file-options',help='JSON object keyed by relative filename with per-file include, column, id_index, sheet, header_row, encoding and delimiter')
    parser.add_argument('--id-mode',choices=ID_MODES,default='Exact text')
    args=parser.parse_args()
    if args.expected<0:parser.error('--expected must be nonnegative')
    try:
        file_options=None
        if args.file_options:
            try:
                with open(args.file_options,encoding='utf-8') as handle:file_options=json.load(handle)
                if not isinstance(file_options,dict) or any(not isinstance(v,dict) for v in file_options.values()):raise ValueError
            except (OSError,ValueError):raise ExtractionError('Cannot read per-file options. Use a JSON object keyed by relative filename, with one settings object per file.') from None
        result=extract(args.source,args.ids,args.output,id_column=args.id_column,selection_column=args.selected_id_column,
                       sheet=args.sheet,expected=args.expected or None,encoding=args.encoding,selection_encoding=args.selection_encoding,
                       delimiter=args.delimiter,selection_delimiter=args.selection_delimiter,recursive=args.recursive,
                       selection_header_row=args.selection_header_row,source_header_row=args.source_header_row,
                       filter_column=args.filter_column,filter_values=args.filter_values,include_patterns=args.include,exclude_patterns=args.exclude,
                       strict_headers=args.strict_headers,skip_invalid_files=args.skip_invalid_files,file_options=file_options,id_mode=args.id_mode,source_files=args.source_file)
        print(json.dumps({key:result[key] for key in ['status','selected_unique_ids','ids_found_anywhere','ids_missing_everywhere','total_rows_scanned','total_rows_selected']},indent=2))
        return 0
    except (ExtractionError,OSError,UnicodeError) as error:
        print(str(error) if isinstance(error,ExtractionError) else 'Check input files, encoding and permissions.',file=sys.stderr)
        return 1


if __name__=='__main__':
    raise SystemExit(main())
