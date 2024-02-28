# python-course-project-2024
Python project for Advanced Programming course

# Project Aims and Motivation
For my project I would like to use Python to implement Excel functionality into LabVIEW, which I have already started looking at so I have a slight headstart here!

The native LabVIEW Excel code is very buggy and unreliable, causing Excel to crash randomly and inconsistently. However, I have used Python for Excel before with the openpyxl package.

LabVIEW allows "Python Nodes" to access Python functions, with arguments limited to specific types (e.g. float, int, string, array - not numpy arrays).

# Final Notes
I have made several simple read/write functions in XL.py to work with LabVIEW's Python nodes. They work consistently and will only cause an error if the Excel file is open on that device, which I have hopefully handled in Python.

I used the same docstring format as openpyxl but in the future I would like to stick to one (e.g. numpy) and then use Sphinx etc. to automate readthedocs documentation, should I upload a package to PyPI.