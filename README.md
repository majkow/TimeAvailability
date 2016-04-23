# TimeAvailability
A place for People to enter volunteers availability into a spreadsheet using Google forms.

To get working.

Download and upload the .xlsx to google drive/sheets and convert.

create a form in the Spreadsheet with 8 questions. 1st question should be a dropdown list. "give it a title of enter your name" mark as required.
next 7 questions should be checkbox's

get your id for the form by looking at the URL and extract the part id part eg. https://docs.google.com/forms/d/1fCpAR_WSlv_2epXzuPfjbQAug6eTLfPSmjIoUREfD5Y/edit?usp=drive_web would be 1fCpAR_WSlv_2epXzuPfjbQAug6eTLfPSmjIoUREfD5Y

open the .gs file in a text editor and copy/paste into the script editor.
find the function that deals with updating the form and paste in your form id.
chagne the menu name from Coniston to what you desire
delete the rawdata sheet and rename the form response 1 to rawdata

enjoy.
