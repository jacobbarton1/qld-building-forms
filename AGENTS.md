# AGENTS.md
 
## Purpose
- Build Python3 + tkinter application to assist an engineer with filling out inspections forms (see file ./template.pdf)

## Functionallity
- each text entry area in the form should be configurable with defualt values via a defaults.json file
- once the gui has been poulated, the user can click a "generate" button which will create a copy of ./template.pdf and prompt the user to save the new pdf under their prefered directory and file name
- the user can also click 'save' which will save the details to a json file but not generate the pdf.
- if the user tries to generate a pdf for a project that has either been loaded from and existing json file or recently saved under a new json name, the application should remember the filename and directory the file was loaded from / saved to and offer this as the suggested path and filename of the pdf once the user clicks 'generate'
- the application will automatically create a json file with the same name as the pdf which will reside in the same directory, allowing the user to load a previously filled out form
- only the first 2 pages of ./template.pdf shall be used for generating poplated documents (the last 2 pages are ronly to assist with comleteing the form)

## Extra Features
- maintain a global.json file inside the root directory to remember 'building certifier' (part 7) and 'appointed competent person' (part 8 details)
- the user can manually fill these out, or use '+' icon in the app to generate details of of either the 'building certifier' or 'appointed competent person' based on previous forms completed.
- both the 'building certifier' or 'appointed competent person' are added each time the 
approval number
- whenever new details for a building certifier' or 'appointed competent person' are detected during a 'save' or 'generate' event, the new details shall be added to the global.json file without further input from the user
- ensure no duplication of 'building certifier' or 'appointed competent person' when addeing new entries to global.json 