

## About

A script that reads an Action Item spreadsheet, detailing the tasks assigned to each officer, and emails each officer a reminder of the tasks they need to complete.

## Set-up for development

1. Clone this repository. If you don't want to do that, download it as a zip file.
2. You should have access to the client_secret files you need in the IEEE Officer Google Drive. Download those files and put them into this directory.
3. Run the script. This will create a .credentials sub-directory, containing files necessary to run the script.

```
python Mailer.py
```


## Using A Different spreadsheet

If you'd like to use a different spreadsheet, make sure to update the spreadsheet_id
variable.

```python
spreadsheet_id = "ID of spreadsheet you want to use"

```
