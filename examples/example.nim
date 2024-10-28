import ../docxutils
import std/tables

proc main() =
  let docxFile = "./file.docx"

  ######################## docxContains ############################### 
  let check = docxContains(docxFile, "ipsum")
  if check:
    echo "ipsum does appear in ", docxFile
  else:
    echo "ipsum does not appear in ", docxFile

  let images = docxImages(docxFile)
  for image in images:
    echo "found image: ", image
  ##################################################################### 


  ######################## docxCopy Image #############################
  docxCopyImage(docxFile, images[0], "copiedImage.jpeg")
  echo "copied the image from the document"
  ######################## docxCopy Image #############################

  
  ######################## docxReadImage  #############################
  let imageData = docxReadImage(docxFile, images[0])
  echo "the first bytes of the image are: ", imageData[0..2]
  ##################################################################### 


  ######################## docxReadRaw  ############################### 
  let xmlContent = docxReadRaw(docxFile)
  echo "Start of the xml is: ", xmlContent[0..54]
  ##################################################################### 


  ######################## docxReadText ############################### 
  let docxText = docxReadText(docxFile)
  echo "Start of the docx file is: ", docxText[0..51]
  #####################################################################


  ######################## docxReplace ################################ 
  var context: Table[string, string]
  context["lorem"] = "foo"
  context["ipsum"] = "bar"
  docxReplace(docxFile, context, "replaced.docx")
  echo "replaced 'lorem' with 'foo' and 'ipsum' with 'bar'"
  #####################################################################


when isMainModule:
  main()
