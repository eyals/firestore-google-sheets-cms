# Firestore CMS with Google Sheets
A Google App Script for managing a Firebase Firestore database using Google Sheets.

## Key features
* Two way sync between Firestore and Sheets
* Manage multiple Firestore collections with a single spreadsheet
* Mark fields as mandatory to protect from publishing invalid content
* Mark fields as no-sync to prevent them from going to Firestore
* Supports array field types
* Sheet validation styles
* Built in setup for Firebase authentication

## How to install
You only need to follow this once:
1. Make a copy of this spreadsheet: [bit.ly/firescore-cms](https://bit.ly/firestore-cms).\
You should see a new Firestore menu added in your copy. It might take a few seconds t appear.\
This means that the script is loaded successfully.
1. From the 'Firestore' menu, select 'Configure service account', and follow the instructions to set up a connection to your Firestore project.\
You can also skip this step and you'll be asked to do it later when trying to sync.

That's it!

Alternatively, from a Google Sheet, click Extensions -> Apps Script, and paste the the code from [this file](dist/firestore-sheets-cms.gs) as is. No neet to change anything.

## How to use

### Collection setup
This file can manage all your Firestore collections.\
Follow these steps for every collection you want to manage.
1. Create a new sheet (bottom tab), and name it with the Firestore collection name
1. From the 'Firestore' menu, select 'Prepare sheet'.\
This will add two critical columns: _active and _id.\
**DO NOT rename these columns**.\
Note - you can run this on an existing sheet, and it will add the missing columns.
1. Add the columns headers to match the field names of the documents in the collection.\
Make sure to include any possible field.
1. If you have content in your collection, click the 'Sync' option from the menu to load it to your sheet

### Content operations
1. **Active checkbox**
   - When you add a new row - a checkbox will be added automatically.\
   - An unchecked item will be deleted from Firestore on the next sync., but it will remain in your sheet so you can always check it on to restore.\
   - This allows you to work on drafts before publishing, or to temporarily deactivate content.
1. **Item ID**
   - You should give each row a unique ID. \
   - This will be the document ID, and will be used to compare between sheets and Firestore.\
   - Avoid special characters or spaces. Try Camelcase like `theLostArc`, or understores like `the_lost_arc`.\
1. **Mandatory fields**
   - To mark a field as mandatory, suffix the column header with `*`. E.g `title*`.\
   You **don't** need to make that change in Firestore. The script will sync teh `title*` column with a `title` field in Firestore.
   - If a mandatory field is missing from an active row, it will highlight it as a problem.
   - Rows with missing content in a mandatory field will **not** be published to Firestore.
1. **No-Sync fields**
   - To mark a field as no-sync, suffix the column header with `~`. E.g `notes~`.\
   - No-sync fields are ignored when syncing to Firestore.
   - This is useful for fields that are just there to help you manage, like calculated fields, notes, or image previews.
1. **Array fields**
   - Other than strings, numbers, and boolean (true/false) fields, the sync script also supports **arrays** of strings and numbers.
   - This is useful for tagging content or for listing simple properties.
   - Arrays should be surrounded by `[` `]`. E.g: `[a,b]`.
   - **Do not** use quotation marks like `["a","b"]`.
   - Cells with square brackets will be colored in purple to help you identify this special content type. 

## What happens on sync?
1. Each row in your sheet is published to Firestore, replacing an existing document with the same ID, or creating a new one if it doesn'e exist.
1. Invalid rows, missing an ID or a mandatory field - will not be published.
1. If a row is unchecked - it will be removed from Firestore
1. Once all the content from your sheet is up, any additional docs in Firestore will be added to your sheet, so that your sheet will have the full picture.