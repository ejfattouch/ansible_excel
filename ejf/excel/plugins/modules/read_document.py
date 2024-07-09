#!/usr/bin/python
# Copyright (c) 2024 Edward-Joseph Fattouch (ejfattouch@outlook.com)
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

DOCUMENTATION = r'''
module: read_document
author:
    - Edward-Joseph Fattouch (@ejfattouch)
short_description: Reads an entire excel document
description:
    - Reads an entire excel document using the openpyxl module.
requirements:
    - "openpyxl python module"
options:
  path:
    description:
      - Path to the excel document.
    type: str
'''

EXAMPLES = r"""
- name: Read an entire excel document
  ejf.read_excel_document:
    path: /your/path/excel/document.xlsx
  register: document
"""

RETURN = r"""
path:
    description: The path to the excel document.
    type: str
    returned: always
sheets:
    description: List containing the names of the sheets.
    type: list
    returned: always
content:
    description: The contents of each sheets of the document.
    type: dict
    sample: {'Sheet1': [...], 'Sheet2': [...], ...}
"""

from ansible.module_utils.basic import AnsibleModule
from openpyxl import *
import os


def main():
    module_args = dict(path=dict(type='path', required=True))
    results = {}

    module = AnsibleModule(
        argument_spec=module_args,
        supports_check_mode=True,
    )

    if not os.path.isfile(module.params['path']):
        module.fail_json(msg="The specified excel file does not exist at " + module.params['path'])

    xl_wb = load_workbook(filename=module.params['path'], data_only=True)
    sheetNames = xl_wb.sheetnames

    resultMap = {}
    for sheetName in sheetNames:
        worksheet = xl_wb[sheetName]
        ws_content = []
        rows = worksheet.rows
        for row in rows:
            row_arr = [cell.value for cell in row]
            ws_content.append(row_arr)
        resultMap[sheetName] = ws_content

    results['path'] = module.params['path']
    results['sheets'] = sheetNames
    results['content'] = resultMap

    module.exit_json(**results)


if __name__ == '__main__':
    main()
