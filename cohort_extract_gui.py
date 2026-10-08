"""Desktop GUI: extract all rows for a selected cohort id cohort."""
from datetime import datetime
import os
import copy
from pathlib import Path
import queue
import subprocess
import sys
import threading
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from extract_core import APP_HOME, VERSION, ID_MODES, ExtractionError, Cancelled, DELIMITERS, inspect_selection, selection_layout, header_number, load_selection, inspect_tables, sample_matches, extract, excel_sheets


class App:
    def __init__(self,root):
        self.root=root
        self.bus=queue.Queue()
        self.cancel=threading.Event()
        self.busy=False
        self.last_output=None
        self.controls=[]
        self.file_options={};self.file_scope='';self.file_info={}
        self.vars={k:tk.StringVar(value=v) for k,v in {
            'selected':'','folder':'','output':'','selected_column':'record_id','id_column':'record_id',
            'sheet':'','expected':'','encoding':'utf-8-sig','selection_encoding':'utf-8-sig','delimiter':'Auto','selection_delimiter':'Auto',
            'selection_header_row':'Auto','source_header_row':'1','filter_column':'','filter_values':'',
            'include_patterns':'*','exclude_patterns':'','header_policy':'Allow unrelated blank/duplicate headings','invalid_policy':'Stop on invalid file headers',
            'id_mode':'Exact text',
        }.items()}
        self.recursive=tk.BooleanVar(value=False)
        self.status=tk.StringVar(value='Ready. Choose your selected-id list and source export folder.')
        self.list_status=tk.StringVar(value='CSV/TSV or Excel .xlsx list; IDs are matched as text, preserving leading zeros.')
        self.summary=tk.StringVar(value='No files inspected yet.')
        root.title('Cohort Data Extractor — Offline — v'+VERSION)
        root.geometry('1190x850')
        root.minsize(1000,760)
        style=ttk.Style(root)
        style.theme_use('clam')
        font='Helvetica Neue' if sys.platform=='darwin' else 'Segoe UI'
        style.configure('.',font=(font,10),background='#eef3f7',foreground='#163047')
        style.configure('TButton',padding=(12,7))
        style.configure('Primary.TButton',background='#17666b',foreground='white',font=(font,10,'bold'))
        style.map('Primary.TButton',background=[('active','#125459'),('disabled','#a5b7b7')])
        outer=ttk.Frame(root,padding=16)
        outer.pack(fill='both',expand=True)
        ttk.Label(outer,text='Cohort Data Extractor',font=(font,24,'bold')).pack(anchor='w')
        ttk.Label(outer,text='Your selected IDs → matching rows from each chosen source table',foreground='#496477').pack(anchor='w',pady=(4,8))
        tk.Label(outer,text='Local extraction retains original values and identifiers. It does not anonymize data.',bg='#fff1d7',fg='#704c15',anchor='w',padx=10,pady=8).pack(fill='x',pady=(0,12))
        self.tabs=ttk.Notebook(outer)
        self.tabs.pack(fill='both',expand=True)
        self.main=ttk.Frame(self.tabs,padding=10)
        self.settings=ttk.Frame(self.tabs,padding=16)
        self.tabs.add(self.main,text='  Select & extract  ')
        self.tabs.add(self.settings,text='  Matching settings  ')
        self.build_main()
        self.build_settings()
        self.build_choices()
        self.build_files()
        footer=ttk.Frame(outer)
        footer.pack(side='bottom',fill='x',before=self.tabs,pady=(12,0))
        ttk.Label(footer,textvariable=self.status,wraplength=900).pack(anchor='w',pady=(0,5))
        self.progress=ttk.Progressbar(footer,mode='indeterminate')
        self.progress.pack(fill='x')
        root.protocol('WM_DELETE_WINDOW',self.close)
        root.after(100,self.poll)

    def entry(self,parent,key):
        item=ttk.Entry(parent,textvariable=self.vars[key])
        self.controls.append(item)
        return item

    def button(self,parent,text,fn,primary=False):
        item=ttk.Button(parent,text=text,command=fn,style='Primary.TButton' if primary else 'TButton')
        self.controls.append(item)
        return item

    def combo(self,parent,key,values=()):
        item=ttk.Combobox(parent,textvariable=self.vars[key],values=list(values),state='readonly')
        self.controls.append(item)
        return item

    def build_main(self):
        frame=self.main
        frame.columnconfigure(1,weight=1)
        ttk.Label(frame,text='1  Selected cohort ids',font=('Segoe UI',12,'bold')).grid(row=0,column=0,columnspan=4,sticky='w',pady=(0,4))
        ttk.Label(frame,text='ID list').grid(row=1,column=0,sticky='w',padx=(0,12))
        self.entry(frame,'selected').grid(row=1,column=1,columnspan=2,sticky='ew')
        self.button(frame,'Choose list…',self.choose_list).grid(row=1,column=3,padx=(8,0))
        row=ttk.Frame(frame)
        row.grid(row=2,column=1,columnspan=3,sticky='ew',pady=3)
        ttk.Label(row,text='Excel sheet').pack(side='left')
        self.sheet_combo=self.combo(row,'sheet')
        self.sheet_combo.configure(width=18)
        self.sheet_combo.pack(side='left',padx=(5,14))
        self.sheet_combo.bind('<<ComboboxSelected>>',lambda e:self.load_headers())
        ttk.Label(row,text='ID column').pack(side='left')
        self.id_combo=self.combo(row,'selected_column')
        self.id_combo.configure(width=22)
        self.id_combo.pack(side='left',padx=5)
        self.button(row,'Read headers',self.load_headers).pack(side='left',padx=8)
        ttk.Label(frame,textvariable=self.list_status,wraplength=930,foreground='#496477').grid(row=3,column=0,columnspan=4,sticky='w',pady=(0,6))
        ttk.Label(frame,text='2  source data exports',font=('Segoe UI',12,'bold')).grid(row=4,column=0,columnspan=4,sticky='w',pady=(0,4))
        ttk.Label(frame,text='Source folder').grid(row=5,column=0,sticky='w')
        self.entry(frame,'folder').grid(row=5,column=1,columnspan=2,sticky='ew')
        self.button(frame,'Choose folder…',self.choose_folder).grid(row=5,column=3,padx=(8,0))
        self.button(frame,'Check id list & all file headers',self.inspect).grid(row=6,column=1,columnspan=3,sticky='w',pady=(3,4))
        columns=('file','size','state','scanned','selected','ids')
        self.table=ttk.Treeview(frame,columns=columns,show='headings',height=7)
        for key,label,width in [('file','File',265),('size','Size',90),('state','Status',190),('scanned','Rows scanned',105),('selected','Rows selected',105),('ids','IDs found',100)]:
            self.table.heading(key,text=label)
            self.table.column(key,width=width,minwidth=65,stretch=key in ('file','state'),anchor='w' if key in ('file','state') else 'e')
        self.table.grid(row=7,column=0,columnspan=4,sticky='nsew')
        frame.rowconfigure(7,weight=1)
        scroll=ttk.Scrollbar(frame,orient='vertical',command=self.table.yview)
        scroll.grid(row=7,column=4,sticky='ns')
        self.table.configure(yscrollcommand=scroll.set)
        self.rows={}
        ttk.Label(frame,textvariable=self.summary,wraplength=950,foreground='#496477').grid(row=8,column=0,columnspan=4,sticky='w',pady=4)
        ttk.Label(frame,text='3  Extract a new copy',font=('Segoe UI',12,'bold')).grid(row=9,column=0,columnspan=4,sticky='w',pady=(0,4))
        ttk.Label(frame,text='New output folder').grid(row=10,column=0,sticky='w')
        self.entry(frame,'output').grid(row=10,column=1,columnspan=2,sticky='ew')
        self.button(frame,'Choose parent…',self.choose_output).grid(row=10,column=3,padx=(8,0))
        actions=ttk.Frame(frame)
        actions.grid(row=11,column=0,columnspan=4,sticky='w',pady=(8,0))
        self.button(actions,'Extract selected IDs',self.run,True).pack(side='left')
        self.cancel_button=ttk.Button(actions,text='Cancel',command=self.cancel.set,state='disabled')
        self.cancel_button.pack(side='left',padx=8)
        ttk.Button(actions,text='Open completed output',command=self.open_output).pack(side='left')
        self.button(actions,'Load fictional demo',self.load_demo).pack(side='left',padx=8)

    def build_settings(self):
        frame=self.settings
        frame.columnconfigure(1,weight=1)
        options=[('id_column','Default ID column for new / unconfigured files'),('expected','Expected unique ids (blank = any count)'),('encoding','Default source-file encoding'),('selection_encoding','CSV id-list encoding')]
        for row,(key,label) in enumerate(options):
            ttk.Label(frame,text=label).grid(row=row,column=0,sticky='w',padx=(0,12),pady=8)
            self.entry(frame,key).grid(row=row,column=1,sticky='ew',pady=8)
        for row,key,label in [(4,'delimiter','Source-file delimiter'),(5,'selection_delimiter','CSV id-list delimiter')]:
            ttk.Label(frame,text=label).grid(row=row,column=0,sticky='w',pady=8)
            self.combo(frame,key,DELIMITERS).grid(row=row,column=1,sticky='ew',pady=8)
        check=ttk.Checkbutton(frame,text='Include CSV/TSV files in subfolders',variable=self.recursive)
        check.grid(row=6,column=0,columnspan=2,sticky='w',pady=(12,14))
        self.controls.append(check)
        text=('Exact id IDs are matched after trimming surrounding whitespace. Leading zeros and letter case in IDs are preserved. '
              'For example, 00123 and 123 are different IDs. Duplicate selected IDs are counted once.\n\n'
              'The expected count is optional. This checks unique IDs before extraction. Change it if your selected cohort changes. '
              'Use an already selected id list, or explicitly choose the cohort-label column and accepted values in Import & cohort choices. cohort membership is never inferred.\n\n'
              'CSV/TSV, .csv.gz and .tsv.gz files are supported. Use include/exclude filename patterns to choose tables. '
              'By default, an unusable ID column or malformed record stops the run. Explicit skipping applies only to invalid headers; skipped files are audited.\n\n'
              'Output data/ retains each source filename and relative folder, all original columns, duplicate records and matching row order. '
              'Values are preserved; CSV quoting and encoding can change. Output uses UTF-8 with a BOM for Excel.\n\n'
              'file_summary.csv reports each file. id_coverage.csv lists selected IDs and their total matched rows. '
              'missing_id_ids.csv lists selected IDs absent from every source file. audit.json records totals and completion status.\n\n'
              'Close other tools that write to the source files. Incomplete folders from failure or cancellation are not completed subsets. '
              'The desktop EXE includes Python, Tkinter and Excel support and makes no network requests.')
        details=ScrolledText(frame,wrap='word',height=8,font=('Segoe UI',10),bg='#f6f9fc',fg='#496477',relief='flat')
        details.insert('1.0',text)
        details.configure(state='disabled')
        details.grid(row=7,column=0,columnspan=2,sticky='nsew')
        frame.rowconfigure(7,weight=1)

    def build_files(self):
        self.files_page=ttk.Frame(self.tabs,padding=12);frame=self.files_page
        self.tabs.add(frame,text='  Files & ID columns  ')
        ttk.Label(frame,text='Include only the source files you want; configure the id-ID column separately for each file.',wraplength=900).pack(anchor='w')
        actions=ttk.Frame(frame);actions.pack(fill='x',pady=10)
        for text,fn in [('Read files',self.scan_files),('Include all',lambda:self.set_file_inclusion(True)),('Exclude all',lambda:self.set_file_inclusion(False)),
                        ('Toggle selected',lambda:self.set_file_inclusion(None)),('Choose ID column / sheet',self.configure_file)]:self.button(actions,text,fn).pack(side='left',padx=(0,5))
        holder=ttk.Frame(frame);holder.pack(fill='both',expand=True)
        columns=('include','file','column','sheet','header','state')
        self.file_tree=ttk.Treeview(holder,columns=columns,show='headings',selectmode='extended')
        for key,label,width in [('include','Use',50),('file','CSV / Excel file',230),('column','ID ID column',180),('sheet','Excel worksheet',130),('header','Header row',85),('state','Status',210)]:
            self.file_tree.heading(key,text=label);self.file_tree.column(key,width=width,minwidth=50,stretch=key in ('file','state'))
        self.file_tree.grid(row=0,column=0,sticky='nsew');holder.rowconfigure(0,weight=1);holder.columnconfigure(0,weight=1)
        vertical=ttk.Scrollbar(holder,orient='vertical',command=self.file_tree.yview);vertical.grid(row=0,column=1,sticky='ns');self.file_tree.configure(yscrollcommand=vertical.set)
        horizontal=ttk.Scrollbar(holder,orient='horizontal',command=self.file_tree.xview);horizontal.grid(row=1,column=0,sticky='ew')
        self.file_tree.configure(xscrollcommand=horizontal.set)
        self.file_tree.bind('<Double-1>',lambda e:self.configure_file())
        self.file_tree.bind('<space>',lambda e:self.set_file_inclusion(None))
        ttk.Label(frame,text='Double-click a file to select its ID column, worksheet, header row, encoding and delimiter. Unticked files are not extracted.\nExcel .xlsx source files are exported as CSV; cell formatting and workbook layout are not retained.\nIf matches are zero, review matching_diagnostics.csv. Matching modes are under Import & cohort choices.',wraplength=950).pack(anchor='w',pady=10)
        self.button(frame,'Test ID matching on first 10,000 rows per included file',lambda:self.launch('sample')).pack(anchor='w')
        self.button(frame,'Reset per-file settings to current defaults',self.reset_file_choices).pack(anchor='w',pady=(5,0))

    def reset_file_choices(self):
        if self.busy:return
        self.file_options={};self.scan_files()

    def refresh_file_choices(self, files):
        self.file_tree.delete(*self.file_tree.get_children());self.file_info={}
        for position,item in enumerate(files):
            relative=item['file'];config=copy.deepcopy(item['config'])
            self.file_options[relative]=config;self.file_info[str(position)]=item
            self.file_tree.insert('', 'end',iid=str(position),values=('Yes' if config['include'] else 'No',relative,config['column'],config['sheet'],config['header_row'],item['status']))

    def set_file_inclusion(self,include):
        if self.busy:return
        if include is True:self.vars['include_patterns'].set('*');self.vars['exclude_patterns'].set('')
        keys=self.file_tree.get_children() if include is not None else self.file_tree.selection()
        for key in keys:
            relative=self.file_info[key]['file'];cfg=self.file_options[relative]
            cfg['include']=not cfg['include'] if include is None else include
            values=list(self.file_tree.item(key,'values'));values[0]='Yes' if cfg['include'] else 'No';self.file_tree.item(key,values=values)

    def configure_file(self):
        if self.busy:return
        selected=self.file_tree.selection()
        if str(Path(self.vars['folder'].get()).resolve())!=self.file_scope:
            messagebox.showinfo('Source folder changed','Click Read files to load the current source folder first.');return
        if not selected:messagebox.showinfo('Choose a file','Read files, then select a source file.');return
        key=selected[0];relative=self.file_info[key]['file'];cfg=copy.deepcopy(self.file_options[relative]);path=Path(self.vars['folder'].get())/relative
        dialog=tk.Toplevel(self.root);dialog.title('Configure source file — '+relative);dialog.geometry('720x420');dialog.transient(self.root);dialog.grab_set()
        body=ttk.Frame(dialog,padding=16);body.pack(fill='both',expand=True);body.columnconfigure(1,weight=1)
        variables={name:tk.StringVar(value=str(cfg[name])) for name in ('sheet','header_row','encoding','delimiter')}
        included=tk.BooleanVar(value=cfg['include']);headers=[];read_signature=None;selected_column=tk.StringVar()
        for row,(name,label) in enumerate([('sheet','Excel worksheet'),('header_row','Header row (Auto or number)'),('encoding','CSV encoding'),('delimiter','CSV delimiter')]):
            ttk.Label(body,text=label).grid(row=row,column=0,sticky='w',padx=(0,12),pady=6)
            widget=ttk.Combobox(body,textvariable=variables[name],values=list(DELIMITERS),state='readonly') if name=='delimiter' else ttk.Entry(body,textvariable=variables[name])
            if name=='sheet':
                widget=ttk.Combobox(body,textvariable=variables[name],state='readonly')
                if path.suffix.lower()=='.xlsx':
                    try:widget.configure(values=['']+excel_sheets(path))
                    except Exception:pass
            widget.grid(row=row,column=1,sticky='ew',pady=6)
        ttk.Label(body,text='ID-ID column position').grid(row=4,column=0,sticky='w',pady=6)
        column_combo=ttk.Combobox(body,textvariable=selected_column,state='readonly');column_combo.grid(row=4,column=1,sticky='ew',pady=6)
        status=tk.StringVar(value='Read headers, select the actual id-ID column, then Save.')
        def read():
            nonlocal headers,read_signature
            headers=[];read_signature=None
            try:
                headers,sheets,number=selection_layout(path,variables['sheet'].get(),variables['encoding'].get(),variables['delimiter'].get(),True,variables['header_row'].get(),cfg['column'])
                values=[f'{i+1}: {name or "(blank header)"}' for i,name in enumerate(headers)]
                column_combo.configure(values=values)
                index=cfg.get('id_index')
                if index is None or index>=len(headers):index=next((i for i,h in enumerate(headers) if h.strip().casefold()==cfg['column'].strip().casefold()),0)
                selected_column.set(values[index]);variables['header_row'].set(str(number))
                if sheets and not variables['sheet'].get():variables['sheet'].set(sheets[0])
                read_signature={name:var.get() for name,var in variables.items()}
                status.set('Headers read. Select the id-ID column; column positions distinguish duplicate names.')
            except Exception as error:status.set(str(error) if isinstance(error,ExtractionError) else 'Cannot read this file. Check sheet, header row, encoding and delimiter.')
        def save():
            try:
                if not headers or column_combo.current()<0:raise ExtractionError('Read the headers and select a id-ID column first.')
                if read_signature!={name:var.get() for name,var in variables.items()}:raise ExtractionError('Settings changed. Click Read headers again before saving the column choice.')
                index=column_combo.current();header_number(variables['header_row'].get())
                cfg.update({name:var.get() for name,var in variables.items()});cfg.update(column=headers[index],id_index=index,include=included.get())
                self.file_options[relative]=cfg
                self.file_tree.item(key,values=('Yes' if cfg['include'] else 'No',relative,cfg['column'],cfg['sheet'],cfg['header_row'],'configured — run Check'))
                dialog.destroy()
            except ExtractionError as error:status.set(str(error))
        ttk.Checkbutton(body,text='Include this file in extraction',variable=included).grid(row=5,column=0,columnspan=2,sticky='w',pady=6)
        ttk.Label(body,textvariable=status,wraplength=660).grid(row=6,column=0,columnspan=2,sticky='w',pady=8)
        ttk.Button(body,text='Read headers',command=read).grid(row=7,column=0,sticky='w');ttk.Button(body,text='Save file choices',command=save).grid(row=7,column=1,sticky='e')
        read()

    def scan_files(self):
        self.tabs.select(self.files_page);self.launch('files')

    def choose_list(self):
        path=filedialog.askopenfilename(filetypes=[('ID lists','*.csv *.tsv *.gz *.xlsx'),('All files','*')])
        if path:
            self.vars['selected'].set(path)
            self.vars['sheet'].set('')
            self.vars['filter_column'].set('');self.vars['filter_values'].set('')
            self.load_headers()

    def build_choices(self):
        frame=ttk.Frame(self.tabs,padding=16);self.tabs.add(frame,text='  Import & cohort choices  ')
        frame.columnconfigure(1,weight=1)
        ttk.Label(frame,text='Import and cohort choices',font=('Segoe UI',12,'bold')).grid(row=0,column=0,columnspan=2,sticky='w',pady=(0,10))
        options=[('selection_header_row','ID-list header row (Auto or row number)'),('source_header_row','Source-file header row (Auto or row number)'),
                 ('filter_values','Keep cohort-label values (comma separated)'),('include_patterns','Include filenames (* = all; separate patterns with ;)'),('exclude_patterns','Exclude filenames (optional; separate patterns with ;)')]
        rows={'selection_header_row':1,'source_header_row':2,'filter_values':4,'include_patterns':5,'exclude_patterns':6}
        for key,label in options:
            ttk.Label(frame,text=label).grid(row=rows[key],column=0,sticky='w',padx=(0,12),pady=7)
            self.entry(frame,key).grid(row=rows[key],column=1,sticky='ew',pady=7)
        ttk.Label(frame,text='Optional cohort-label column (blank = every listed ID)').grid(row=3,column=0,sticky='w',padx=(0,12),pady=7)
        self.filter_combo=self.combo(frame,'filter_column',['']);self.filter_combo.grid(row=3,column=1,sticky='ew',pady=7)
        for row,key,label,choices in [(7,'header_policy','Unrelated column headings',['Allow unrelated blank/duplicate headings','Require all headings unique and nonempty']),
                                      (8,'invalid_policy','Files with invalid headers',['Stop on invalid file headers','Skip files with invalid headers']),
                                      (9,'id_mode','ID-ID matching',ID_MODES)]:
            ttk.Label(frame,text=label).grid(row=row,column=0,sticky='w',pady=7)
            self.combo(frame,key,choices).grid(row=row,column=1,sticky='ew',pady=7)
        text=('Read headers after changing the list header row or worksheet. Auto searches the first 50 rows and, when no sheet is selected, each worksheet for your ID heading.\n\n'
              'If your selection file contains a larger source cohort, select its cohort label column and explicitly enter the value identifying cohort (for example 1 or selected). No label is inferred. The expected count checks the filtered unique IDs.\n\n'
              'File examples: Provider_Notes*;Labs* includes only those tables. Exclude *backup* to omit backups. Check files before extraction; excluded and skipped files are listed in the audit.\n\n'
              'The id-ID heading must always be unique. All original source columns and headings are retained. Malformed data records always stop extraction; rows are never silently discarded. IDs retain leading zeros and case.')
        details=ScrolledText(frame,wrap='word',height=8,font=('Segoe UI',10));details.insert('1.0',text);details.configure(state='disabled')
        details.grid(row=10,column=0,columnspan=2,sticky='nsew',pady=(10,0));frame.rowconfigure(10,weight=1)

    def load_headers(self):
        if self.busy:return
        try:
            path=self.vars['selected'].get()
            if Path(path).suffix.lower()=='.xlsx':
                sheets=excel_sheets(path)
                self.sheet_combo.configure(values=sheets)
                if self.vars['sheet'].get() not in sheets:self.vars['sheet'].set('')
            headers,sheets,number=selection_layout(path,self.vars['sheet'].get(),self.vars['selection_encoding'].get(),self.vars['selection_delimiter'].get(),True,self.vars['selection_header_row'].get(),self.vars['selected_column'].get())
            self.sheet_combo.configure(values=sheets)
            if sheets and not self.vars['sheet'].get():self.vars['sheet'].set(sheets[0])
            headers=[h for h in headers if h.strip()]
            self.id_combo.configure(values=headers)
            self.filter_combo.configure(values=['']+headers)
            configured=self.vars['id_column'].get().strip().casefold()
            chosen=next((h for h in headers if h.strip().casefold()==configured),headers[0])
            self.vars['selected_column'].set(chosen)
            self.list_status.set(f'Headers loaded from row {number}. Confirm the sheet and ID column; optional cohort filtering is in Import & cohort choices.')
        except ExtractionError as error:
            self.list_status.set('ID list could not be read. Correct the sheet or file and click Read headers again.')
            messagebox.showerror('Cannot read id list',str(error))
        except Exception as error:
            messagebox.showerror('Cannot read id list',f'ID-list reader failed ({type(error).__name__}). Check mapped-drive access and save the correct worksheet as CSV UTF-8, with record_id on row 1. Choose that CSV and click Read headers. No id data is included in this message.')

    def choose_folder(self):
        path=filedialog.askdirectory(title='Choose the folder containing all source source exports')
        if path:self.vars['folder'].set(path)

    def load_demo(self):
        self.file_options={};self.file_scope='';self.vars['id_mode'].set('Exact text')
        base=APP_HOME
        if not (base/'examples/selected_ids.csv').is_file() and (base/'examples.zip').is_file():
            with zipfile.ZipFile(base/'examples.zip') as archive:
                for name in ('selected_ids.csv','source/Entities.csv','source/Events.csv','source/Transactions.csv'):
                    target=base/'examples'/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(archive.read(name))
        values={'selected':str(base/'examples/selected_ids.csv'),'folder':str(base/'examples/source'),
                'output':str(base/('cohort_demo_output_'+datetime.now().strftime('%Y%m%d_%H%M%S'))),
                'selected_column':'record_id','id_column':'record_id','sheet':'','expected':'5',
                'encoding':'utf-8-sig','selection_encoding':'utf-8-sig','delimiter':'Auto','selection_delimiter':'Auto'}
        for key,value in values.items():self.vars[key].set(value)
        self.recursive.set(False)
        for key,value in {'selection_header_row':'Auto','source_header_row':'1','filter_column':'','filter_values':'','include_patterns':'*','exclude_patterns':'',
                          'header_policy':'Allow unrelated blank/duplicate headings','invalid_policy':'Stop on invalid file headers'}.items():self.vars[key].set(value)
        self.load_headers()
        self.list_status.set('Fictional demo only: 5 selected IDs and three source tables; expected output is 9 rows. Click Check, then Extract.')

    def choose_output(self):
        path=filedialog.askdirectory(title='Choose a parent for a NEW cohort output folder')
        if path:self.vars['output'].set(str(Path(path)/('cohort_selected_'+datetime.now().strftime('%Y%m%d_%H%M%S'))))

    def options(self, job='extract'):
        result={k:v.get().strip() for k,v in self.vars.items()}
        if not result['folder'] or (job!='files' and not result['selected']):
            raise ExtractionError('Choose both the selected-id list and the source data folder.')
        try:result['expected']=int(result['expected']) if result['expected'] else None
        except ValueError:raise ExtractionError('Expected id count must be a positive integer or blank.') from None
        if result['expected'] is not None and result['expected']<1:
            raise ExtractionError('Expected id count must be positive, or leave it blank.')
        result['recursive']=self.recursive.get()
        header_number(result['selection_header_row']);header_number(result['source_header_row'])
        result['strict_headers']=result['header_policy']=='Require all headings unique and nonempty'
        result['skip_invalid_files']=result['invalid_policy']=='Skip files with invalid headers'
        scope=str(Path(result['folder']).resolve())
        if scope!=self.file_scope:self.file_options={};self.file_scope=scope
        result['file_options']=copy.deepcopy(self.file_options)
        return result

    def launch(self,job):
        if self.busy:return
        try:
            values=self.options(job)
            if job=='extract' and not values['output']:raise ExtractionError('Choose a new output folder.')
        except ExtractionError as error:
            messagebox.showerror('Check settings',str(error));return
        self.busy=True
        self.cancel.clear()
        for widget in self.controls:widget.configure(state='disabled')
        self.cancel_button.configure(state='normal' if job=='extract' else 'disabled')
        self.progress.start(12)
        self.status.set('Checking the selected list and file headers locally…')
        threading.Thread(target=self.worker,args=(job,values),daemon=True).start()

    def worker(self,job,v):
        try:
            if job=='files':
                table_info,ignored=inspect_tables(v['folder'],v['selected'] or None,v['id_column'],v['encoding'],v['delimiter'],v['recursive'],v['source_header_row'],v['strict_headers'],v['include_patterns'],v['exclude_patterns'],v['file_options'])
                self.bus.put(('file_choices',table_info));self.bus.put(('done','Source files read. Choose which files to use and configure their ID columns, then Check and Extract.'));return
            selected=load_selection(v['selected'],v['selected_column'],v['sheet'],v['selection_encoding'],v['selection_delimiter'],v['expected'],v['selection_header_row'],v['filter_column'],v['filter_values'])
            if job=='sample':
                table_info,ignored=sample_matches(v['folder'],selected,file_options=v['file_options'],id_column=v['id_column'],header_row=v['source_header_row'],encoding=v['encoding'],delimiter=v['delimiter'],recursive=v['recursive'],selection_path=v['selected'],include_patterns=v['include_patterns'],exclude_patterns=v['exclude_patterns'],id_mode=v['id_mode'])
                self.bus.put(('matching_sample',table_info));self.bus.put(('done','Sample comparison complete. Zero sample matches do not prove that a id is absent from later rows.'));return
            table_info,ignored=inspect_tables(v['folder'],v['selected'],v['id_column'],v['encoding'],v['delimiter'],v['recursive'],v['source_header_row'],v['strict_headers'],v['include_patterns'],v['exclude_patterns'],v['file_options'])
            self.bus.put(('inspected',(selected,table_info,ignored)))
            if job=='inspect':
                self.bus.put(('done','Inspection complete. Files marked ready can be extracted.'));return
            audit=extract(v['folder'],v['selected'],v['output'],selection_column=v['selected_column'],id_column=v['id_column'],
                          sheet=v['sheet'],encoding=v['encoding'],selection_encoding=v['selection_encoding'],delimiter=v['delimiter'],
                          selection_delimiter=v['selection_delimiter'],recursive=v['recursive'],expected=v['expected'],cancel=self.cancel,
                          selection_header_row=v['selection_header_row'],source_header_row=v['source_header_row'],filter_column=v['filter_column'],filter_values=v['filter_values'],
                          strict_headers=v['strict_headers'],include_patterns=v['include_patterns'],exclude_patterns=v['exclude_patterns'],skip_invalid_files=v['skip_invalid_files'],
                          file_options=v['file_options'],id_mode=v['id_mode'],
                          progress=lambda event:self.bus.put(('progress',event)))
            self.bus.put(('complete',(v['output'],audit)))
        except ExtractionError as error:self.bus.put(('error',str(error)))
        except Exception:self.bus.put(('error','Cannot process the inputs. Check permissions, encoding and file format. No source data was included in this error message.'))

    def inspect(self):self.launch('inspect')
    def run(self):self.launch('extract')

    def finish(self):
        self.busy=False
        self.progress.stop()
        for widget in self.controls:widget.configure(state='readonly' if isinstance(widget,ttk.Combobox) else 'normal')
        self.cancel_button.configure(state='disabled')

    def poll(self):
        try:
            while True:
                kind,value=self.bus.get_nowait()
                if kind=='file_choices':self.refresh_file_choices(value)
                if kind=='matching_sample':
                    self.refresh_file_choices(value)
                    lines=[]
                    for item in value:
                        line=item['file']+': '+item['status']
                        if 'mode_counts' in item:line+='\n  '+', '.join(f'{mode}: {count}' for mode,count in item['mode_counts'].items())
                        lines.append(line)
                    messagebox.showinfo('ID matching sample', '\n\n'.join(lines[:30])+'\n\nFirst 10,000 rows per included file only. A zero sample result does not prove absence. No matching mode is changed automatically.')
                if kind=='inspected':
                    selected,files,ignored=value
                    self.refresh_file_choices(files)
                    self.list_status.set(f'{len(selected.ids):,} unique IDs; {selected.filtered_out:,} rows excluded by cohort filter; {selected.duplicates:,} duplicate IDs; {selected.blanks:,} blank IDs. Header row {selected.header_row}.')
                    self.table.delete(*self.table.get_children())
                    self.rows={}
                    for f in files:
                        self.rows[f['file']]=self.table.insert('', 'end',values=(f['file'],f"{f['bytes']/1024**3:.2f} GB",f['status'],'','',''))
                    errors=sum(f['status'] not in ('ready','excluded by file choices') for f in files)
                    self.summary.set(f'{len(files)} source tables • {sum(f["bytes"] for f in files)/1024**3:.2f} GB • {errors} header issues • {len(ignored)} excluded files (see audit on completion).')
                elif kind=='progress':
                    f=value
                    item=self.rows.get(f['file'])
                    if item:
                        old=list(self.table.item(item,'values'))
                        old[2]='complete' if f.get('file_finished') else 'processing'
                        old[3]=f"{f['rows_scanned']:,}";old[4]=f"{f['rows_selected']:,}"
                        self.table.item(item,values=old)
                        self.table.see(item)
                    self.status.set(f"File {f['file_index']}/{f['file_count']}: {f['file']} • {f['rows_scanned']:,} rows scanned • {f['rows_selected']:,} selected • {f['elapsed_seconds']:.0f} s")
                elif kind=='complete':
                    self.last_output,audit=value
                    for f in audit['files']:
                        item=self.rows[f['file']]
                        old=list(self.table.item(item,'values'));old[2]='complete';old[3]=f"{f['rows_scanned']:,}";old[4]=f"{f['rows_selected']:,}";old[5]=f"{f['ids_found']:,}"
                        self.table.item(item,values=old)
                    skipped=sum('skipped' in f['reason'] for f in audit['ignored_files'])
                    self.status.set(f"Complete: {audit['total_rows_selected']:,} rows; {audit['ids_found_anywhere']:,}/{audit['selected_unique_ids']:,} ids found; {audit['ids_missing_everywhere']:,} absent from processed files; {skipped} files skipped. Review output reports.")
                    self.finish()
                    if audit['total_rows_selected']==0:messagebox.showwarning('No IDs matched','No rows matched in any selected file. Check each file’s ID column in Files & ID columns. Open matching_diagnostics.csv to see whether case or integer .0 formatting explains the mismatch. Nothing is matched approximately or changed automatically.')
                elif kind=='done':self.status.set(value);self.finish()
                elif kind=='error':
                    self.status.set(value);self.finish();messagebox.showerror('Extraction stopped',value)
        except queue.Empty:pass
        self.root.after(100,self.poll)

    def open_output(self):
        if not self.last_output:
            messagebox.showinfo('Output','No completed extraction is available yet.');return
        if sys.platform=='win32':os.startfile(self.last_output)
        elif sys.platform=='darwin':subprocess.Popen(['open',self.last_output])
        else:subprocess.Popen(['xdg-open',self.last_output])

    def close(self):
        if self.busy:
            messagebox.showinfo('Extraction running','Cancel and wait for processing to stop before closing.');return
        self.root.destroy()


if __name__=='__main__':
    root=tk.Tk()
    App(root)
    root.mainloop()
