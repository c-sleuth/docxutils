import zippy/ziparchives
import std/xmlparser
import std/xmltree
import std/tables
import std/strutils
import std/strformat


proc docxReadRaw*(docxFile: string): string =
  
  ## `docxFile` path to docx file
  ##
  ## Will return the raw xml of the document ("document.xml")
  ##
  ## Example:
  ##
  ##  `let content: string = docxReadRaw("file.docx")`

  let reader = openZipArchive(docxFile)
  try:
    let document = reader.extractFile("word/document.xml")
    return document
  finally:
    reader.close()


proc docxReplace*(docxFile: string, context: Table[string, string],
    outputFile: string) =

  ## `docxFile` path to docx file
  ##
  ## `context` Table containing formatted as {"text_to_replace": "replacement_text"}
  ##
  ## `outputFile` file path for new document with replacements
  ##
  ## This will read through the docx file and replace all occurences of the key of the Table with the assosiated value
  ##
  ## Example:
  ##
  ##  `var context = Table[string, string]`
  ##
  ##  `context["text_to_replace"] = "replacement_text"`
  ##
  ##  `docxReplace("file.docx", context, "new_file.docx")`

  let reader = openZipArchive(docxFile)
  try:
    let docxTemplate = reader.extractFile("word/document.xml")
    var updatedTemplate: string = docxTemplate
    for key, value in context:
      updatedTemplate = replace(updatedTemplate, key, value)

    var entries: Table[string, string]

    for path in reader.walkFiles:
      if path != "word/document.xml":
        entries[path] = reader.extractFile(path)

    entries["word/document.xml"] = updatedTemplate
    let archive = createZipArchive(entries)
    writeFile(outputFile, archive)
  finally:
    reader.close()


proc docxContains*(docxFile: string, str: string): bool =

  ## `docxFile` path to docx file
  ##
  ## `str` string to match
  ##
  ## This will return true if the string exists in the docx file and return false if not
  ##
  ## Example:
  ##
  ##  `let match: bool = docxContains("file.docx", "text_to_find")`

  let contents = docxReadRaw(docxFile)
  if find(contents, str) != -1:
    return true
  else:
    return false


proc docxReadText*(docxFile: string): string =

  ## `docxFile` path to docx file
  ##
  ## This will return all the text within the docx file
  ##
  ## Example:
  ##
  ##  `let content: string = docxReadText("file.docx")`

  let contents = parseXml(docxReadRaw(docxFile))
  var text: string
  for i in contents.findAll("w:t"):
    text.add(i.innerText)
  return text


proc docxImages*(docxFile: string): seq[string] =

  ## `docxFile` path to docx file
  ##
  ## This will return a sequence of all the image file names found within the docx file
  ##
  ## Example:
  ##
  ##  `let images: seq[string] = docxImages("file.docx")`

  let reader = openZipArchive(docxFile)
  try:
    var images: seq[string]
    for imagePath in reader.walkFiles:
      if imagePath.startsWith("word/media/"):
        var image = imagePath
        image.removePrefix("word/media/")
        images.add(image)

    return images
  finally:
    reader.close()

proc docxCopyImage*(docxFile: string, imgName: string, dst: string) =

  ## `docxFile` path to docx file
  ##
  ## `imgName` name of image to copy
  ##
  ## `dst` the destination for the image (including filename)
  ##
  ## Example:
  ##
  ##  `docxCopyImage("file.docx", "image1.png", "output/image1.png")`

  let reader = openZipArchive(docxFile)
  try:
    let image = reader.extractFile(fmt"word/media/{imgName}")
    writeFile(dst, image)
  finally:
    reader.close()

proc docxReadImage*(docxFile: string, imgName: string): string =

  ## `docxFile` path to docx file
  ##
  ## `imgName` name of image to copy
  ##
  ## This will return the raw data of a specified image file within a docx file
  ##
  ## Example:
  ##
  ##  `let imageData: string = docxReadImage("file.docx", "image1.png")`

  let reader = openZipArchive(docxFile)
  try:
    let image = reader.extractFile(fmt"word/media/{imgName}")
    return image
  finally:
    reader.close()
