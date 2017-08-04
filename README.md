# stickyprints
Print your scrum task notes onto stickies using a docx template and an xlsx task list

## How it works

### Creating the sticky template

First you make a one page docx file containing the template for your stickies.

Make a table the size of a sticky (3x3 inch or 7.62 x 7.62 cm) and fill it with
the content you want on your stickies. Use \<placeholders\> for information that
is different for each sticky and that will be taken from the task list.

Now add some more tables of the same size to the page, but leave them empty.
These tables will tell you where to place the stickies when you are going to
feed them into the printer.

![The sticky template document](/images/template.png?raw=true)

### Creating the task list

Now create an xlsx file containing the tasks for which you want stickies to be
printed. Each column header should match with the name of a placeholder in the
template.

![An example task list](/images/tasks.png?raw=true)

### Run stickyprints.py

![The stickyprints UI](/images/stickyprints.png?raw=true)

Select the template and task list you created and click on generate. You'll be
asked where to save the results. The results will look like this:

![The document with generated stickies](/images/stickies.png?raw=true)

### Prepare the stickies for printing

*Important note*: Only use one-sided printing when working with stickies,
otherwise your stickies may get stuck inside your printer.

In order to print on stickies, the stickies will have to be attached to a piece
of A4 paper. Use print-outs of the template page so you can easily see where
the stickies should be placed.

### Print task information on the stickies

*Important note*: Only use one-sided printing when working with stickies,
otherwise your stickies may get stuck inside your printer.

Put the pages with stickies in the appropriate printer tray with the sticky
side towards the printer. Now you can print the generated stickies.docx file.

## Installation

This program requires [openpyxl](https://pypi.python.org/pypi/openpyxl).

If you want this script to be runnable from a central location without requiring
users to install openpyxl and its dependencies, then do something like this:

1. Download [openpyxl](https://pypi.python.org/pypi/openpyxl) and extract it somewhere
2. Move the openpyxl subdirectory that it contains to the directory that contains stickyprints.py
3. Download [et_xmlfile](https://pypi.python.org/pypi/et_xmlfile) and extract it somewhere
4. Move the et_xmlfile subdirectory that it contains to the directory that contains stickyprints.py
5. Download [jdcal](https://pypi.python.org/pypi/jdcal) and extract it somewhere
6. Move the jdcal.py file that it contains to the directory that contains stickyprints.py

Now you should be able to run stickyprints.py without installing openpyxl
