/**
 * @NotOnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

function test() {
  var body = DocumentApp.getActiveDocument().getBody();
  body.editAsText().replaceText("<List>", "Item1\nItem2");
  // Append a new list item to the body.
  //var item1 = body.appendListItem('Item 1');
  //  item1.getListId();
  //  body.appendParagraph("Test");
  //  body.appendListItem('Item 2').setListId(item1);
  //

  Logger.log(item1.getListId());
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename)
    .evaluate()
    .getContent();
}
function onOpen() {
  DocumentApp.getUi()
    .createMenu("Runbook Creator")
    .addItem("Start", "showDialog")
    .addSeparator()
    .addItem("Help", "showHelp")
    .addToUi();
}

function compare(a, b) {
  if (a.name < b.name) {
    return -1;
  }
  if (a.name > b.name) {
    return 1;
  }
  return 0;
}

function getModules(modulesFolderId) {
  var listOfModules = DriveApp.getFolderById(modulesFolderId).getFiles();
  var result = [];
  while (listOfModules.hasNext() == true) {
    var module = listOfModules.next();

    result.push({
      name: module.getName(),
      id: module.getId(),
      type: module.getMimeType()
    });
  }
  return result.sort(compare);
}

function saveProgress(listOfAddedModules) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty("currentProgress", listOfAddedModules);
}

function loadProgress() {
  var props = PropertiesService.getDocumentProperties();
  return props.getProperty("currentProgress")
    ? props.getProperty("currentProgress")
    : "";
}

function getProp(key) {
  var props = PropertiesService.getDocumentProperties();

  Logger.log(props.getProperty("currentProgress"));
  return props.getProperty("currentProgress") ? props.getProperty(key) : "";
}

function createFieldsFromTemplates(id) {
  var template = DocumentApp.openById(id);
  var jspattern = /\${[^{]*}/g;
  var textResult;
  var allTemplatesStrings = template
    .getBody()
    .editAsText()
    .getText()
    .match(jspattern);
  var result = {};
  allTemplatesStrings.forEach(function(element) {
    result[element] = "";
  });
  allTemplatesStrings = Object.keys(result);
  return allTemplatesStrings;
}

function cursorTest() {
  var elem = DocumentApp.getActiveDocument()
    .getBody()
    .editAsText()
    .findText("TR.IUE.CHK_EXP_REG_SYS")
    .getElement();
  var position = DocumentApp.getActiveDocument().newPosition(elem, 7);
  DocumentApp.getActiveDocument().setCursor(position);
}

function createRunbook(rdata) {
  console.log(rdata);
  var destination = DocumentApp.getActiveDocument();
  for (var i = 0; i < rdata.length; i++) {
    var module = DriveApp.getFileById(rdata[i].id);
    Logger.log("Test");
    Logger.log(module.getId());
    var newFile = module.makeCopy();
    var replacements = rdata[i].templates;
    replacements.forEach(function(replacer) {
      Logger.log(replacer.template);
      DocumentApp.openById(newFile.getId())
        .getBody()
        .editAsText()
        .replaceText(
          "\\" + replacer.template,
          replacer.value.replace(";", "\n")
        );
    });
    var title = DocumentApp.getActiveDocument()
      .getBody()
      .appendParagraph(rdata[i].title);
    title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    DocumentApp.getActiveDocument()
      .getBody()
      .editAsText()
      .setFontFamily("Raleway");
    appendDocument(
      DocumentApp.openById(newFile.getId()),
      DocumentApp.getActiveDocument()
    );
    newFile.setTrashed(true);
  }
  DocumentApp.getActiveDocument()
    .getBody()
    .editAsText()
    .setFontFamily("Raleway");
}
function showDialog() {
  var html = HtmlService.createTemplateFromFile("page")
    .evaluate()
    .setWidth(1000)
    .setHeight(600);
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModalDialog(html, "Runbook Creator");
}
function showHelp() {
  var html = HtmlService.createTemplateFromFile("help")
    .evaluate()
    .setWidth(600)
    .setHeight(600);
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModalDialog(html, "Runbook Creator Help");
}

function appendDocument(source, destination) {
  //All images in the tables should be inserted as a Drawing, tobe copied to a new document
  //var source = DocumentApp.openById("1eWFx9D8aZpJWWwReRkz1sYS2P4m_ZFY2IQ_wFLMDdac");
  //var destination= DocumentApp.getActiveDocument();
  var sourceBody = source.getBody();
  var destinationBody = destination.getBody();
  var totalElements = sourceBody.getNumChildren();
  for (var i = 0; i < totalElements; i++) {
    var element = sourceBody.getChild(i).copy();
    Logger.log(element.getType());
    var type = element.getType();
    if (type == DocumentApp.ElementType.PARAGRAPH)
      destinationBody.appendParagraph(element);
    else if (type == DocumentApp.ElementType.TABLE)
      destinationBody.appendTable(element);
    else if (type == DocumentApp.ElementType.LIST_ITEM) {
      destinationBody
        .appendListItem(element)
        .setListId(element)
        .setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
    } else if (type == DocumentApp.ElementType.INLINE_IMAGE)
      destinationBody.appendImage(element);
    else if (type == DocumentApp.ElementType.PAGE_BREAK)
      destinationBody.appendPageBreak(element);
    else throw new Error("Unknown element type: " + type);
  }
}
