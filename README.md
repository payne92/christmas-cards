This is a simple Python utility to typeset Christmas card envelopes from a list 
stored in a Google Docs spreadsheet.  This allows easy editing of the list
by multiple users (complete with version history).

To configure:

Create a Google Docs spreadsheet with the following columns:

* card ('yes' if to be included)
* name
* address
* address2
* address3
* city
* state
* zip
* country

Then, copy sample-config to ~/.christmascard and edit with your Google username, password, and mailing address.

Type "make" to build document.pdf

