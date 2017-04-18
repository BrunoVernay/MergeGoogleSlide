# MergeGoogleSlide
Merge Google spreadsheet rows in a Google Slide


## What I did
I get the objectId of the textElements to replace

Loop:
- Duplicate the page, fixing the new Objects Id
- Changing the text in the first page

### Todo

- The Google Ids for the template slide and the spreadsheet are hardcoded in the sources.
- The Spreadsheet headers should be read from the Google Sheet and used by default as field names, with the {{column_header}}. So there would be no need to write a mapping

## Inspiration
Not that easy to create one slide per row in a spreadsheet.
The example gives one presentation (=file) per row !!! https://developers.google.com/slides/how-tos/merge
- Problem is that the replaceText act on everything (even on the layouts in the Master slide)
  There is no way to reduce the scope (or the context) of the "replace all text" action!
- I did not find a way to get the whole page to recreate it later.
- I did not find a way to get the ObjectId from the Web (GUI) interface
