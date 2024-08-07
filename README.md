# Ansible Collection - ejfattouch.excel

This ```ejfattouch.excel``` collection provides modules that allow for reading and writing data to and from Excel files.

## Requirements
- Ansible: >= 2.9.10
- openpyxl
- xlwings and installed instance of Excel (optional, required only on macOS or Windows, needed only for function evaluations)

## Install
Ansible must be installed
```shell
sudo pip install ansible
```
There are 2 ways to install this collection:
### 1- Install from Ansible Galaxy
```shell
ansible-galaxy collection install ejfattouch.excel
```
### 2- Install from source
Clone the repo.
```shell
git clone https://github.com/ejfattouch/ansible_excel.git
```
Go to the excel directory
```shell
cd ejfattouch/excel
```
Pull the latest main
```shell
git pull origin main
```
Build and Install a collection from source
```shell
ansible-galaxy collection build --force
ansible-galaxy collection install ejfattouch-excel-* --force
```


## Modules

| Name          | Description                                                              |
|---------------|--------------------------------------------------------------------------|
| read_document | Reads an entire Excel document and returns its contents.<br/>            |
| read_sheet    | Reads a specified sheet from an Excel document and returns its contents. |
| write_sheet   | Writes data to a specified sheet in an Excel document                    |

## Use
Once the collection is installed, it can be used in a playbook by specifying the full namespace path to the plugin.
```yaml
- hosts: localhost
  gather_facts: no

  tasks:
    - name: Read data in an Excel document
      ejfattouch.excel.read_document:
        path: /your/path/excel/document.xlsx
      register: document

    - name: Read sheet Sheet1 in an Excel document
      ejfattouch.excel.read_sheet:
        path: /your/path/excel/document.xlsx
        sheet: "Sheet1"
      register: sheet1

    - name: Write data to a single cell in an Excel document
      ejfattouch.excel.write_sheet:
        path: /your/path/excel/document.xlsx
        sheet: "Sheet1"
        cell: B10
        data: "your_data"
```