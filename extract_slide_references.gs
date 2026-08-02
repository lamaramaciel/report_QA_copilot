/**
 * Optional Google Slides extractor for Slide QA Copilot.
 *
 * Add this code to a script bound to a Google Slides presentation:
 * Extensions → Apps Script. Run onOpen once, authorize, then use the
 * "Slide QA Copilot" menu in the presentation.
 *
 * The script creates a Google Sheet containing one row per referenced paragraph.
 * Download that Sheet as CSV and upload it to the Slide QA Copilot.
 */

function onOpen() {
  SlidesApp.getUi()
    .createMenu('Slide QA Copilot')
    .addItem('Extract text and links', 'extractSlideTextAndLinks')
    .addToUi();
}

function extractSlideTextAndLinks() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const output = SpreadsheetApp.create(presentation.getName() + ' - Slide QA Extraction');
  const sheet = output.getSheets()[0];
  sheet.setName('Slide Claims');

  const headers = [
    'Slide', 'Slide Title', 'Element ID', 'Element Name', 'Element Type',
    'Paragraph', 'Raw Text', 'Clean Claim', 'Ref Markers', 'URLs',
    'Mapping Note', 'Surrounding Text'
  ];
  const rows = [headers];

  slides.forEach((slide, slideIndex) => {
    const title = getSlideTitle_(slide);
    walkElements_(slide.getPageElements(), (element) => {
      collectElementRows_(element, slideIndex + 1, title, rows);
    });
  });

  if (rows.length === 1) {
    SlidesApp.getUi().alert('No referenced text was found. Confirm that the references are true hyperlinks.');
    return;
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1F2937').setFontColor('#FFFFFF');
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(7, 430);
  sheet.setColumnWidth(8, 430);
  sheet.setColumnWidth(10, 500);
  sheet.setColumnWidth(12, 500);
  sheet.getDataRange().setWrap(true).setVerticalAlignment('top');

  SlidesApp.getUi().alert(
    'Done! The extraction was created in Google Sheets:\n\n' + output.getUrl() +
    '\n\nOpen it, review the claim-to-reference mapping, and download as CSV.'
  );
}

function walkElements_(elements, callback) {
  elements.forEach((element) => {
    if (element.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      walkElements_(element.asGroup().getChildren(), callback);
    } else {
      callback(element);
    }
  });
}

function collectElementRows_(element, slideNumber, slideTitle, rows) {
  const type = String(element.getPageElementType());
  const elementId = element.getObjectId ? element.getObjectId() : '';
  const elementName = element.getTitle ? (element.getTitle() || '') : '';

  if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
    const shape = element.asShape();
    const text = shape.getText();
    collectTextRangeRows_(text, slideNumber, slideTitle, elementId, elementName, type, rows);
  }

  if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
    const table = element.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      for (let c = 0; c < table.getNumColumns(); c++) {
        const text = table.getCell(r, c).getText();
        collectTextRangeRows_(
          text, slideNumber, slideTitle, elementId + '-R' + (r + 1) + 'C' + (c + 1),
          elementName + ' [R' + (r + 1) + 'C' + (c + 1) + ']', 'TABLE_CELL', rows
        );
      }
    }
  }
}

function collectTextRangeRows_(textRange, slideNumber, slideTitle, elementId, elementName, elementType, rows) {
  const surrounding = textRange.asRenderedString().trim();
  const paragraphs = textRange.getParagraphs();

  paragraphs.forEach((paragraphRange, paragraphIndex) => {
    const rawText = paragraphRange.asRenderedString().trim();
    if (!rawText) return;

    const urls = [];
    const linkRanges = paragraphRange.getLinks();
    linkRanges.forEach((linkRange) => {
      const link = linkRange.getTextStyle().getLink();
      if (link && link.getLinkType() === SlidesApp.LinkType.URL) {
        const url = link.getUrl();
        if (url && urls.indexOf(url) === -1) urls.push(url);
      }
    });

    const visibleUrls = rawText.match(/https?:\/\/[^\s|,\]>)"']+/g) || [];
    visibleUrls.forEach((url) => {
      url = url.replace(/[.,;)]$/, '');
      if (urls.indexOf(url) === -1) urls.push(url);
    });
    if (!urls.length) return;

    const markers = rawText.match(/\[Ref\d*\]/gi) || [];
    const cleanClaim = rawText
      .replace(/\[Ref\d*\]/gi, '')
      .replace(/https?:\/\/\S+/g, '')
      .replace(/\s+/g, ' ')
      .trim();

    let note = 'Hyperlinks detected without explicit [Ref] markers';
    if (markers.length === urls.length && markers.length > 0) {
      note = 'Direct marker-to-link mapping available';
    } else if (markers.length > urls.length) {
      note = 'Incomplete mapping: ' + markers.length + ' marker(s), ' + urls.length + ' URL(s)';
    } else if (urls.length > markers.length && markers.length > 0) {
      note = 'Extra URLs: ' + markers.length + ' marker(s), ' + urls.length + ' URL(s)';
    }
    if (!cleanClaim) note = 'Reference-only paragraph; claim mapping unclear';

    rows.push([
      slideNumber, slideTitle, elementId, elementName, elementType,
      paragraphIndex + 1, rawText, cleanClaim, markers.join(' | '),
      urls.join(' | '), note, surrounding
    ]);
  });
}

function getSlideTitle_(slide) {
  const shapes = slide.getShapes();
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    if (shape.getPlaceholderType && shape.getPlaceholderType() === SlidesApp.PlaceholderType.TITLE) {
      return shape.getText().asRenderedString().trim();
    }
  }
  for (let i = 0; i < shapes.length; i++) {
    const text = shapes[i].getText().asRenderedString().trim();
    if (text) return text.split('\n')[0].substring(0, 180);
  }
  return '';
}
