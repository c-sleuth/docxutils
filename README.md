# docxutils

A small utility module for working with docx files without having to create any temp file or directory

This was cobbled together quickly, error handling is minimal and by minimal I mean next to none... sorry I just needed these features quickly

## Use Cases

The primary use case in mind when making this was to replace text within template style docx files

## Features

Currently this module supports the following actions:

- Replacing multiple strings with one call
- Read text of docx file
- Read xml of docx file
- Read image data from embedded images
- List embedded images
- Copy embedded images to a destination
- Check if file contains a string

## Examples and Docs

Docs have can be generated with nim using `nim doc docxutils.nim`

The docx file within the example directory was taken from (https://file-examples.com/index.php/sample-documents-download/sample-doc-download/) and is the 100kB docx file

## Dependencies

- Zippy (https://github.com/guzba/zippy) is used to unzip the docx files without having to create any temp files or directories
- Nim Standard Library

