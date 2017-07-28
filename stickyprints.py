#!/usr/bin/env python

import configparser
import copy
import openpyxl
import os
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile

from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror, showinfo

########################################################################
#
#  Sticky processing code
#
########################################################################

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
BODY = WORD_NAMESPACE + 'body'
TR = WORD_NAMESPACE + 'tr'
TC = WORD_NAMESPACE + 'tc'
TBLBORDERS = WORD_NAMESPACE + 'tblBorders'
#TRBORDERS = WORD_NAMESPACE + 'trBorders'
TCBORDERS = WORD_NAMESPACE + 'tcBorders'

def read_tasks_from_excel(excel_filename):
    tasks = []
    wb = openpyxl.load_workbook(excel_filename)
    header = [str(cell.value) for cell in wb.worksheets[0].rows[0]]
    for datarow in wb.worksheets[0].rows[1:]:
        data = [str(cell.value) for cell in datarow]
        task = dict(zip(header, data))
        tasks.append(task)
    return tasks

def extract_zipfile_to_dir(filename, dir_to_extract_to):
    with zipfile.ZipFile(filename, 'r') as zf:
        zf.extractall(dir_to_extract_to)

def make_zipfile_from_dir(source_dir, output_filename):
    relroot = os.path.abspath(source_dir)

    with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(source_dir):
            for filename in files:
                filepath = os.path.join(root, filename)
                if os.path.isfile(filepath):
                    arcname = os.path.join(os.path.relpath(root, relroot), filename)
                    zf.write(filepath, arcname)

def replace_placeholders_with_task_data(cell, task):
        stickytext = ET.tostring(cell, encoding='unicode')
        for key, val in task.items():
            stickytext = re.sub('&lt;%s&gt;' % key, val, stickytext, flags = re.IGNORECASE)
        return ET.fromstring(stickytext)

def remove_all_table_borders(element):
    for borders in element.findall('.//%s' % TBLBORDERS):
        borders.clear()
    for borders in element.findall('.//%s' % TCBORDERS):
        borders.clear()

def create_stickies_from_template(template_filename, tasks, stickies_filename):

    with tempfile.TemporaryDirectory(prefix='stickiestmp') as tempdir:

        extract_zipfile_to_dir(template_filename, tempdir)

        docpath = os.path.join(tempdir, 'word', 'document.xml')
        with open(docpath, 'r') as f:
            xml_content = f.read()

        nr_of_pages = int((len(tasks) - 1) / 6) + 1

        tree = ET.fromstring(xml_content)
        body = tree.find(BODY)
        remove_all_table_borders(body)
        page = copy.deepcopy(list(body))

        # Add as many pages with the same content as the template as we need
        for i in range(nr_of_pages - 1):
            body.extend(page)

        # In each table remove all existing cells and add a copy of the first cell
        template_cell = tree.find('.//%s' % TC)

        tasks.reverse()
        for row in tree.iter(TR):
            for cell in row.iter(TC):
                row.remove(cell)

            if tasks:
                taskcell = replace_placeholders_with_task_data(template_cell, tasks.pop())
                row.append(taskcell)

        ET.ElementTree(tree).write(docpath)

        make_zipfile_from_dir(tempdir, stickies_filename)

def generate_stickies(template_filename, tasks_filename, stickies_filename):
    tasks = read_tasks_from_excel('tasks.xlsx')
    if tasks:
        create_stickies_from_template('template.docx', tasks, 'stickies.docx')
    else:
        print('No tasks, so no stickies')

########################################################################
#
#  User interface code
#
########################################################################

def get_dirname_from_filename(filename):
    return '.' if not filename else os.path.dirname(filename)

def get_settings_file_path():
    return os.path.expanduser(os.path.join('~', '.stickyprints.conf'))

class MyFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.master.resizable(True, False)
        self.master.minsize(800, 0)
        self.master.title("Stickyprints v0.1")
        self.master.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(0, pad=10)
        self.columnconfigure(1, pad=10)
        self.columnconfigure(2, pad=10)
        self.rowconfigure(0, pad=5)
        self.rowconfigure(1, pad=5)
        self.rowconfigure(2, pad=5)
        self.grid(sticky='nsew')

        self.template_label = Label(self, text = 'Template:')
        self.template_entry = Entry(self)
        self.template_button = Button(self, text = 'Change...', command = self.on_change_template_button_clicked)
        self.tasklist_label = Label(self, text = 'Task list (xlsx):')
        self.tasklist_entry = Entry(self)
        self.tasklist_button = Button(self, text = 'Change...', command = self.on_change_tasks_button_clicked)
        self.last_used_stickies_filename = None
        self.generate_button = Button(self, text = 'Generate stickies...', command = self.on_generate_stickies_button_clicked)
        self.quit_button = Button(self, text = 'Quit', command = self.on_quit_button_clicked)

        self.template_label.grid(row = 0, column = 0)
        self.template_entry.grid(row = 0, column = 1, sticky='ew')
        self.template_button.grid(row = 0, column = 2)
        self.tasklist_label.grid(row = 1, column = 0)
        self.tasklist_entry.grid(row = 1, column = 1, sticky='ew')
        self.tasklist_button.grid(row = 1, column = 2)
        self.generate_button.grid(row = 2, column = 1)
        self.quit_button.grid(row = 2, column = 2)

        self.config = configparser.ConfigParser()
        self.load_filenames_from_config()

    def load_filenames_from_config(self):
        self.config.read(get_settings_file_path())

        template_filename = self.config.get('Settings', 'template_file', fallback='')
        self.template_entry.delete(0, END)
        self.template_entry.insert(0, template_filename)

        tasklist_filename = self.config.get('Settings', 'tasklist_file', fallback='')
        self.tasklist_entry.delete(0, END)
        self.tasklist_entry.insert(0, tasklist_filename)

    def save_filenames_to_config(self):
        try:
            self.config.add_section('Settings')
        except configparser.DuplicateSectionError:
            pass
        self.config.set('Settings', 'template_file', self.template_entry.get())
        self.config.set('Settings', 'tasklist_file', self.tasklist_entry.get())
        with open(get_settings_file_path(), 'w') as configfile:
            self.config.write(configfile)

    def on_change_template_button_clicked(self):
        filename = askopenfilename(initialdir=get_dirname_from_filename(self.template_entry.get()),
                                   filetypes=[('Template file', '*.docx')])
        if filename:
            self.template_entry.delete(0, END)
            self.template_entry.insert(0, filename)
            self.save_filenames_to_config()

    def on_change_tasks_button_clicked(self):
        filename = askopenfilename(initialdir=get_dirname_from_filename(self.tasklist_entry.get()),
                                   filetypes=[('Tasks file', '*.xlsx')])
        if filename:
            self.tasklist_entry.delete(0, END)
            self.tasklist_entry.insert(0, filename)
            self.save_filenames_to_config()

    def on_generate_stickies_button_clicked(self):
        if not os.path.exists(self.template_entry.get()):
            showerror('Error', 'The specified template file does not exist. Please select a valid template file.')
            return

        if not os.path.exists(self.tasklist_entry.get()):
            showerror('Error', 'The specified tasklist file does not exist. Please select a valid tasklist file.')
            return

        self.save_filenames_to_config()

        filename = askopenfilename(initialdir=get_dirname_from_filename(self.last_used_stickies_filename),
                                   filetypes=[('Stickies file', '*.docx')])
        if filename:
            self.last_used_stickies_filename = filename
            generate_stickies(self.template_entry.get(),
                              self.tasklist_entry.get(),
                              filename)
            showinfo("Stickies generated",
                     "The stickies were generated successfully.")

    def on_quit_button_clicked(self):
        self.master.destroy()

if __name__ == '__main__':
    MyFrame().mainloop()
