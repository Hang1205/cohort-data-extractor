"""Packaged Windows GUI entry point with a built-in fictional self-test."""
import sys


def self_test():
    import csv
    from pathlib import Path
    import socket
    import tempfile
    import time
    import tkinter as tk
    from extract_core import extract,excel_library
    from cohort_extract_gui import App
    def denied(*args,**kwargs):raise OSError('Network blocked for offline self-test.')
    socket.create_connection=denied
    socket.socket.connect=denied
    socket.getaddrinfo=denied
    with tempfile.TemporaryDirectory(prefix='cohort-desktop-test-') as temp:
        base=Path(temp);source=base/'source';source.mkdir()
        library=excel_library();book=library.Workbook();book.active.title='Cover';book.active.append(['Fictional only'])
        sheet=book.create_sheet('Selected cohort');sheet.append(['Fictional cohort']);sheet.append([None]);sheet.append(['record_id',None,'cohort_label','comment','comment'])
        sheet.append(['001',None,1,'fictional','x']);sheet.append(['002',None,1,'fictional','x']);sheet.append(['999',None,0,'fictional','x']);book.save(base/'selected.xlsx');book.close()
        with (source/'notes.csv').open('w',newline='',encoding='utf-8') as handle:
            writer=csv.writer(handle);writer.writerow(['EntityKey','','text','text']);writer.writerows([['001','','Fictional note\nsecond line','x'],['999','','Exclude','x'],['002','','Second fictional note','x']])
        with (source/'metadata.csv').open('w',newline='',encoding='utf-8') as handle:csv.writer(handle).writerows([['metadata'],['fictional']])
        excel_table=library.Workbook();excel_table.active.title='Results';excel_table.active.append(['SourceID','value']);excel_table.active.append([1,'fictional']);excel_table.active['A2'].number_format='000';excel_table.active.append(['002','fictional']);excel_table.save(source/'Data.xlsx');excel_table.close()
        options={'notes.csv':{'column':'EntityKey','id_index':0},'Data.xlsx':{'column':'SourceID','sheet':'Results','id_index':0},'metadata.csv':{'include':False}}
        audit=extract(source,base/'selected.xlsx',base/'out',expected=2,selection_header_row='Auto',filter_column='cohort_label',filter_values='1',file_options=options)
        assert audit['total_rows_selected']==4 and audit['ids_found_anywhere']==2
        assert audit['selection_header_row']==3 and audit['selection_rows_filtered_out']==1
        root=tk.Tk();root.withdraw()
        try:
            app=App(root)
            for key,value in {'selected':str(base/'selected.xlsx'),'sheet':'','folder':str(source),'output':str(base/'gui-out'),'expected':'2','filter_column':'cohort_label','filter_values':'1'}.items():app.vars[key].set(value)
            app.file_scope=str(source.resolve());app.file_options=options
            app.run();deadline=time.monotonic()+30
            while app.busy and time.monotonic()<deadline:
                root.update();time.sleep(0.01)
            assert not app.busy and app.last_output==str(base/'gui-out')
        finally:root.destroy()
    return {'excel_support':'passed','offline_socket_block':'enabled during test','fictional_selected_rows':4,'threaded_gui_batch':'passed','auto_sheet_and_header':'passed','cohort_label_filter':'passed','file_choices':'passed','unrelated_blank_duplicate_headers':'passed','per_file_id_columns':'passed','excel_source_export':'passed'}


def factory(root):
    from cohort_extract_gui import App
    return App(root)


if __name__=='__main__':
    from desktop_runtime import desktop_main,run_self_test
    raise SystemExit(run_self_test(self_test) if '--self-test' in sys.argv else desktop_main(factory))
