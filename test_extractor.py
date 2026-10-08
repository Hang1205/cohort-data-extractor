"""Synthetic tests only. No hospital records are used."""
import csv
import gzip
import hashlib
import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

from extract_core import ExtractionError,Cancelled,extract,load_selection,inspect_selection,inspect_tables,sample_matches,excel_library


def write_table(path,rows,delimiter=',',encoding='utf-8-sig'):
    path=Path(path)
    path.parent.mkdir(parents=True,exist_ok=True)
    opener=gzip.open if str(path).endswith('.gz') else open
    with opener(path,'wt',encoding=encoding,newline='') as h:
        csv.writer(h,delimiter=delimiter).writerows(rows)


def read_table(path,delimiter=','):
    opener=gzip.open if str(path).endswith('.gz') else open
    with opener(path,'rt',encoding='utf-8-sig',newline='') as h:
        return list(csv.reader(h,delimiter=delimiter))


class Tests(unittest.TestCase):
    def test_general_default_has_no_fixed_cohort_size(self):
        write_table(self.source/'Entities.csv',[['record_id','category'],['001','fictional'],['B2','other']])
        audit=extract(self.source,self.selected,self.output)
        self.assertEqual(audit['selected_unique_ids'],2)
        self.assertIsNone(audit['expected_selected_count'])

    def test_general_default_column_and_multiple_entity_tables(self):
        write_table(self.selected,[['record_id'],['C0001']])
        write_table(self.source/'Orders.csv',[['CustomerKey','amount'],['C0001','12.50'],['C0002','8.00']])
        audit=extract(self.source,self.selected,self.output,file_options={'Orders.csv':{'column':'CustomerKey'}})
        self.assertEqual(audit['total_rows_selected'],1)
        self.assertEqual(audit['ids_missing_everywhere'],0)
    def test_matching_sample_identifies_format_difference_without_output(self):
        write_table(self.source/'notes.csv',[['EntityKey','note'],['001.0','fictional'],['B2','other']])
        results,_=sample_matches(self.source,load_selection(self.selected,expected=2),file_options={'notes.csv':{'column':'EntityKey'}},selection_path=self.selected)
        self.assertEqual(results[0]['sample_matches'],1)
        self.assertEqual(results[0]['mode_counts']['Remove integer .0 suffix'],2)
        self.assertFalse(self.output.exists())

    def test_matching_sample_respects_individual_exclusions(self):
        write_table(self.source/'notes.csv',[['EntityKey','note'],['001','fictional']])
        results,_=sample_matches(self.source,load_selection(self.selected,expected=2),file_options={'notes.csv':{'include':False}})
        self.assertEqual(results[0]['status'],'excluded by file choices')
        self.assertNotIn('sample_matches',results[0])

    def test_individual_file_columns_and_explicit_file_exclusion(self):
        write_table(self.source/'notes.csv',[['MRN','note'],['001','fictional'],['other','exclude']])
        write_table(self.source/'labs.csv',[['EntityKey','value'],['B2','fictional']])
        write_table(self.source/'unusable.csv',[['not_id_data'],['fictional']])
        options={'notes.csv':{'column':'MRN'},'labs.csv':{'column':'EntityKey'},'unusable.csv':{'include':False}}
        audit=self.run_extract(file_options=options)
        self.assertEqual(audit['total_rows_selected'],2);self.assertEqual(audit['ids_missing_everywhere'],0)
        self.assertFalse((self.output/'data/unusable.csv').exists())
        self.assertEqual({f['id_column'] for f in audit['files']},{'MRN','EntityKey'})

    def test_excel_source_sheet_and_numeric_zero_format(self):
        lib=excel_library();book=lib.Workbook();book.active.title='Cover';book.active.append(['Fictional'])
        ws=book.create_sheet('Results');ws.append(['Fictional title']);ws.append(['Source_ID','value'])
        ws.append([1,'fictional lab']);ws['A3'].number_format='000';ws.append(['B2','other'])
        path=self.source/'Labs.xlsx';book.save(path);book.close()
        audit=self.run_extract(file_options={'Labs.xlsx':{'column':'Source_ID','sheet':'Results','header_row':'2'}})
        self.assertEqual(audit['total_rows_selected'],2)
        self.assertEqual(read_table(self.output/'data/Labs.csv'),[['Source_ID','value'],['1','fictional lab'],['B2','other']])
        self.assertEqual(audit['files'][0]['sheet'],'Results')

    def test_explicit_column_position_can_disambiguate_duplicate_names(self):
        write_table(self.source/'notes.csv',[['ID','ID','note'],['other','001','fictional'],['other','B2','second']])
        audit=self.run_extract(file_options={'notes.csv':{'column':'ID','id_index':1}})
        self.assertEqual(audit['total_rows_selected'],2)

    def test_decimal_suffix_mode_preserves_values_and_original_selected_ids(self):
        write_table(self.source/'notes.csv',[['record_id','note'],['001.0','fictional'],['B2','other']])
        audit=self.run_extract(id_mode='Remove integer .0 suffix')
        self.assertEqual(audit['ids_missing_everywhere'],0)
        self.assertEqual(read_table(self.output/'data/notes.csv')[1][0],'001.0')
        self.assertEqual(read_table(self.output/'id_coverage.csv')[1][0],'001')

    def test_zero_match_diagnostics_show_decimal_difference_without_auto_matching(self):
        write_table(self.source/'notes.csv',[['record_id','note'],['001.0','fictional'],['B2.0','other']])
        audit=self.run_extract()
        self.assertEqual(audit['total_rows_selected'],0)
        self.assertEqual(audit['files'][0]['integer_suffix_id_rows'],1)
        self.assertEqual(audit['files'][0]['exact_id_rows'],0)
        self.assertTrue((self.output/'matching_diagnostics.csv').is_file())

    def test_normalization_collision_is_rejected(self):
        write_table(self.selected,[['record_id'],['001'],['001.0']])
        write_table(self.source/'notes.csv',[['record_id','note'],['001','fictional']])
        with self.assertRaisesRegex(ExtractionError,'merges different'):self.run_extract(id_mode='Remove integer .0 suffix')
        self.assertFalse(self.output.exists())

    def test_excel_csv_output_name_collision_is_rejected(self):
        write_table(self.source/'Labs.csv',[['record_id','value'],['001','fictional']])
        book=excel_library().Workbook();book.active.append(['record_id','value']);book.active.append(['001','fictional']);book.save(self.source/'Labs.xlsx');book.close()
        with self.assertRaisesRegex(ExtractionError,'same output filename'):self.run_extract()
        self.assertFalse(self.output.exists())

    def test_per_file_encoding_and_delimiter(self):
        write_table(self.source/'notes.csv',[['EntityKey','note'],['001','fictional café'],['B2','other']],delimiter='|',encoding='cp1252')
        audit=self.run_extract(file_options={'notes.csv':{'column':'EntityKey','encoding':'cp1252','delimiter':'Pipe'}})
        self.assertEqual(audit['total_rows_selected'],2)
        self.assertEqual(read_table(self.output/'data/notes.csv','|')[1][1],'fictional café')

    def test_auto_sheet_and_header_with_cohort_label_filter(self):
        lib=excel_library();book=lib.Workbook();book.active.title='Cover';book.active.append(['Fictional cover'])
        ws=book.create_sheet('Ground truth');ws.append(['Fictional cohort']);ws.append([None]);ws.append(['record_id',None,'cohort_flag','comment','comment'])
        ws.append(['001',None,1,'fictional','x']);ws.append(['other',None,0,'fictional','x']);ws.append(['B2',None,1,'fictional','x'])
        path=self.root/'truth.xlsx';book.save(path);book.close()
        selected=load_selection(path,expected=2,header_row='Auto',filter_column='cohort_flag',filter_values='1')
        self.assertEqual(selected.ids,{'001','B2'})
        self.assertEqual((selected.sheet,selected.header_row,selected.filtered_out),('Ground truth',3,1))

    def test_csv_preamble_auto_delimiter_and_filtered_count(self):
        self.selected.write_text('Fictional cohort\nrecord_id;label;\n001;cohort;\nB2;cohort;\nother;source;\n',encoding='utf-8-sig')
        selected=load_selection(self.selected,expected=2,header_row='Auto',filter_column='label',filter_values='cohort')
        self.assertEqual(selected.ids,{'001','B2'});self.assertEqual(selected.header_row,2)
        with self.assertRaisesRegex(ExtractionError,'unique IDs'):load_selection(self.selected,expected=3,header_row='Auto',filter_column='label',filter_values='cohort')

    def test_source_unrelated_headings_preserved_and_strict_choice(self):
        rows=[['record_id','','note','note'],['001','','fictional','x'],['B2','','other','y']]
        write_table(self.source/'notes.csv',rows)
        with self.assertRaises(ExtractionError):self.run_extract(strict_headers=True)
        audit=self.run_extract()
        self.assertEqual(read_table(self.output/'data/notes.csv'),rows)
        self.assertEqual(audit['total_rows_selected'],2)

    def test_source_header_row_auto_with_preamble(self):
        (self.source/'notes.csv').write_text('Fictional table\nrecord_id;note\n001;fictional\nB2;other\n',encoding='utf-8-sig')
        self.run_extract(source_header_row='Auto')
        self.assertEqual(read_table(self.output/'data/notes.csv',delimiter=';'),[['record_id','note'],['001','fictional'],['B2','other']])

    def test_filename_choices_and_skip_invalid_headers_are_audited(self):
        write_table(self.source/'notes.csv',[['record_id','note'],['001','fictional']])
        write_table(self.source/'Labs.csv',[['record_id','value'],['B2','fictional']])
        write_table(self.source/'metadata.csv',[['unrelated'],['fictional']])
        audit=self.run_extract(include_patterns='notes*;metadata*',skip_invalid_files=True)
        self.assertEqual(audit['total_rows_selected'],1)
        self.assertEqual(audit['ids_missing_everywhere'],1)
        ignored={f['file']:f['reason'] for f in audit['ignored_files']}
        self.assertIn('excluded',ignored['Labs.csv']);self.assertIn('skipped',ignored['metadata.csv'])
        self.assertFalse((self.output/'data/metadata.csv').exists())

    def test_skip_header_policy_does_not_skip_malformed_data(self):
        write_table(self.source/'notes.csv',[['record_id','note'],['001','fictional'],['B2','bad','extra']])
        with self.assertRaisesRegex(ExtractionError,'malformed'):self.run_extract(skip_invalid_files=True)
        self.assertFalse(self.output.exists())

    def test_no_files_after_choices_does_not_publish_empty_output(self):
        write_table(self.source/'notes.csv',[['record_id','note'],['001','fictional']])
        with self.assertRaisesRegex(ExtractionError,'No readable'):self.run_extract(exclude_patterns='*')
        self.assertFalse(self.output.exists())

    def test_automatic_header_scan_has_a_limit_and_manual_row_can_override(self):
        rows=[['Fictional title'] for _ in range(51)]+[['record_id'],['001'],['B2']]
        write_table(self.selected,rows)
        with self.assertRaisesRegex(ExtractionError,'header row'):load_selection(self.selected,header_row='Auto',expected=2)
        self.assertEqual(load_selection(self.selected,header_row='52',expected=2).ids,{'001','B2'})

    def test_formula_cohort_labels_rejected_before_selection(self):
        lib=excel_library();book=lib.Workbook();ws=book.active;ws.append(['record_id','label']);ws.append(['001','=1'])
        path=self.root/'formula-label.xlsx';book.save(path);book.close()
        with self.assertRaisesRegex(ExtractionError,'formula'):load_selection(path,expected=1,filter_column='label',filter_values='1')

    def test_excel_id_list_allows_unrelated_blank_and_duplicate_headers(self):
        lib=excel_library();book=lib.Workbook();ws=book.active
        ws.append(['record_id',None,'comment','comment'])
        ws.append(['001',None,'fictional','value']);ws.append(['B2',None,'other','value'])
        path=self.root/'id-list.xlsx';book.save(path);book.close()
        self.assertEqual(load_selection(path,expected=2).ids,{'001','B2'})
        with self.assertRaises(ExtractionError):inspect_selection(path)

    def test_excel_duplicate_id_id_headers_still_rejected(self):
        lib=excel_library();book=lib.Workbook();ws=book.active
        ws.append(['record_id','record_id']);ws.append(['001','other'])
        path=self.root/'ambiguous.xlsx';book.save(path);book.close()
        with self.assertRaisesRegex(ExtractionError,'missing or ambiguous'):load_selection(path,expected=1)

    def test_workbook_errors_are_actionable_and_do_not_echo_contents(self):
        path=self.root/'bad.xlsx';path.write_text('Fictional private payload',encoding='utf-8')
        with self.assertRaisesRegex(ExtractionError,'CSV UTF-8') as caught:load_selection(path)
        self.assertNotIn('Fictional private payload',str(caught.exception))
        with self.assertRaisesRegex(ExtractionError,'mapped drive'):inspect_selection(self.root/'missing.xlsx',id_list=True)

    def test_empty_first_sheet_can_be_bypassed_with_selected_sheet(self):
        lib=excel_library();book=lib.Workbook();book.active.title='Cover'
        ws=book.create_sheet('IDs');ws.append(['record_id']);ws.append(['001'])
        path=self.root/'sheets.xlsx';book.save(path);book.close()
        with self.assertRaisesRegex(ExtractionError,'header row'):inspect_selection(path,id_list=True)
        self.assertEqual(load_selection(path,sheet='IDs',expected=1).ids,{'001'})

    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.root=Path(self.temp.name)
        self.source=self.root/'source'
        self.source.mkdir()
        self.selected=self.root/'cohort.csv'
        self.output=self.root/'output'
        write_table(self.selected,[['record_id'],['001'],['B2']])

    def tearDown(self):self.temp.cleanup()

    def run_extract(self,**options):
        return extract(self.source,self.selected,self.output,expected=2,**options)

    def test_all_files_all_columns_multiline_duplicates_and_originals(self):
        note='First line, comma\nSecond line "quoted"\n'+('x'*200_000)
        rows=[['record_id','text','extra'],['001',note,'exact'],['not-selected','other','x'],['B2','second','y'],['001',note,'exact']]
        write_table(self.source/'Provider_Notes.csv',rows)
        write_table(self.source/'Labs.csv',[['record_id','lab'],['B2','value'],['unselected','skip']])
        before={p:hashlib.sha256(p.read_bytes()).hexdigest() for p in [self.selected,*self.source.iterdir()]}
        audit=self.run_extract()
        self.assertEqual(read_table(self.output/'data/Provider_Notes.csv'),[rows[0],rows[1],rows[3],rows[4]])
        self.assertEqual(audit['total_rows_selected'],4)
        self.assertEqual(audit['ids_found_anywhere'],2)
        self.assertEqual(len(audit['files']),2)
        self.assertEqual(read_table(self.output/'missing_ids.csv'),[['record_id']])
        self.assertEqual(before,{p:hashlib.sha256(p.read_bytes()).hexdigest() for p in before})

    def test_string_ids_leading_zeros_whitespace_and_case(self):
        write_table(self.source/'ID_Demographics.csv',[[' record_id ','value'],['001','keep'],['1','exclude'],['b2','exclude'],[' B2 ','keep']])
        self.run_extract()
        self.assertEqual(read_table(self.output/'data/ID_Demographics.csv')[1:],[['001','keep'],[' B2 ','keep']])

    def test_missing_column_stops_before_any_output(self):
        write_table(self.source/'bad.csv',[['id_key','note'],['001','x']])
        with self.assertRaises(ExtractionError):self.run_extract()
        self.assertFalse(self.output.exists())
        self.assertFalse(list(self.root.glob('output.incomplete-*')))

    def test_malformed_record_retains_failed_partial_only(self):
        write_table(self.source/'notes.csv',[['record_id','text'],['001','good'],['B2','bad','extra']])
        with self.assertRaises(ExtractionError):self.run_extract()
        self.assertFalse(self.output.exists())
        partial=next(self.root.glob('output.incomplete-*'))
        self.assertEqual(json.loads((partial/'audit.json').read_text())['status'],'failed')

    def test_cancellation_does_not_publish_complete(self):
        write_table(self.source/'notes.csv',[['record_id','text'],*([['001','value']]*2000)])
        cancel=threading.Event()
        with self.assertRaises(Cancelled):self.run_extract(cancel=cancel,progress=lambda event:cancel.set())
        self.assertFalse(self.output.exists())
        self.assertEqual(json.loads((next(self.root.glob('output.incomplete-*'))/'audit.json').read_text())['status'],'cancelled')

    def test_gzip_tsv_semicolon_recursive_headers_only_and_missing_ids(self):
        write_table(self.source/'folder/Vital_Signs.tsv.gz',[['record_id','result'],['001','5\n6'],['excluded','3']],delimiter='\t')
        write_table(self.source/'Labs.csv',[['record_id','result']],delimiter=';')
        audit=self.run_extract(recursive=True)
        self.assertEqual(read_table(self.output/'data/folder/Vital_Signs.tsv.gz','\t'),[['record_id','result'],['001','5\n6']])
        self.assertEqual(read_table(self.output/'data/Labs.csv',';'),[['record_id','result']])
        self.assertEqual(read_table(self.output/'missing_ids.csv'),[['record_id'],['B2']])
        self.assertEqual(audit['ids_missing_everywhere'],1)

    def test_no_overwrite_or_output_inside_input(self):
        write_table(self.source/'a.csv',[['record_id'],['001']])
        self.output.mkdir()
        with self.assertRaises(ExtractionError):self.run_extract()
        with self.assertRaises(ExtractionError):extract(self.source,self.selected,self.source/'out',expected=2)

    def test_duplicate_blank_ids_count_and_expected_guard(self):
        write_table(self.selected,[['record_id','comment'],['001','a'],['001','dup'],[' B2 ','b'],['','blank']])
        selected=load_selection(self.selected,expected=2)
        self.assertEqual((selected.duplicates,selected.blanks,len(selected.ids)),(1,1,2))
        with self.assertRaises(ExtractionError):load_selection(self.selected,expected=842)

    def test_excel_sheet_and_zero_formatted_ids(self):
        lib=excel_library()
        wb=lib.Workbook()
        wb.active.title='Other';wb.active.append(['ignored'])
        ws=wb.create_sheet('Selected cohort')
        ws.append(['record_id','group'])
        ws.append([1,'cohort']);ws['A2'].number_format='000'
        ws.append(['B2','cohort'])
        path=self.root/'selected.xlsx';wb.save(path);wb.close()
        self.assertEqual(inspect_selection(path,'Selected cohort')[1],['Other','Selected cohort'])
        self.assertEqual(load_selection(path,sheet='Selected cohort',expected=2).ids,{'001','B2'})
        write_table(self.source/'a.csv',[['record_id','x'],['001','1'],['B2','2'],['skip','0']])
        audit=extract(self.source,path,self.output,sheet='Selected cohort',expected=2)
        self.assertEqual(audit['total_rows_selected'],2)

    def test_excel_formulas_and_long_numeric_ids_rejected(self):
        lib=excel_library()
        for value in ['=1+1',1234567890123456]:
            wb=lib.Workbook();ws=wb.active;ws.append(['record_id']);ws.append([value])
            path=self.root/'bad.xlsx';wb.save(path);wb.close()
            with self.assertRaises(ExtractionError):load_selection(path,expected=None)

    def test_encoding_and_unsupported_file_audit(self):
        write_table(self.source/'a.csv',[['record_id','note'],['001','café']],encoding='cp1252')
        (self.source/'readme.txt').write_text('Synthetic instructions only')
        audit=self.run_extract(encoding='cp1252',selection_encoding='utf-8-sig')
        self.assertEqual(read_table(self.output/'data/a.csv')[1],['001','café'])
        self.assertEqual(audit['ignored_files'][0]['file'],'readme.txt')

    def test_ambiguous_header_and_inspection(self):
        write_table(self.source/'bad.csv',[['record_id',' record_id '],['001','001']])
        results,_=inspect_tables(self.source,self.selected)
        self.assertNotEqual(results[0]['status'],'ready')
        with self.assertRaises(ExtractionError):self.run_extract()

    def test_source_change_detected(self):
        path=self.source/'a.csv'
        write_table(path,[['record_id','text'],['001','a']])
        changed=False
        def progress(event):
            nonlocal changed
            if event.get('file_finished') and not changed:
                with path.open('a',encoding='utf-8') as handle:handle.write('\n')
                changed=True
        with self.assertRaises(ExtractionError):self.run_extract(progress=progress)
        self.assertFalse(self.output.exists())

    def test_new_source_file_during_run_prevents_completion(self):
        write_table(self.source/'a.csv',[['record_id','text'],['001','value']])
        def progress(event):
            if event.get('file_finished'):
                write_table(self.source/'new.csv',[['record_id'],['B2']])
        with self.assertRaises(ExtractionError):self.run_extract(progress=progress)
        self.assertFalse(self.output.exists())

    def test_storage_failure_does_not_publish_or_leak_payload(self):
        import extract_core
        original_open=extract_core.open_text
        original_write=Path.write_text
        def open_text(path,encoding='utf-8-sig',mode='rt'):
            if mode=='wt':raise OSError('DO_NOT_LOG_PATIENT_TEXT')
            return original_open(path,encoding,mode)
        def write(path,value,*args,**kwargs):
            if path.name=='audit.json' and json.loads(value).get('status')=='failed':raise OSError('Synthetic disk full')
            return original_write(path,value,*args,**kwargs)
        write_table(self.source/'a.csv',[['record_id','text'],['001','Fictional']])
        with patch.object(extract_core,'open_text',open_text),patch.object(Path,'write_text',write):
            with self.assertRaises(ExtractionError) as caught:self.run_extract()
        self.assertNotIn('DO_NOT_LOG_PATIENT_TEXT',str(caught.exception))
        self.assertFalse(self.output.exists())


if __name__=='__main__':unittest.main()
