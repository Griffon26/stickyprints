#!/usr/bin/env python

import copy
import openpyxl
import os
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile

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

def main():
    tasks = read_tasks_from_excel('tasks.xlsx')
    if tasks:
        create_stickies_from_template('template.docx', tasks, 'stickies.docx')
    else:
        print('No tasks, so no stickies')

if __name__ == '__main__':
    main()
