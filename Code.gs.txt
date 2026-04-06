/**
 * Layout Toolkit for Google Docs
 * A free, open-source Google Docs add-on that brings desktop publishing
 * features to your documents. Create brochures, newsletters, event programs,
 * business cards, and flyers with ready-made templates and brand customization.
 *
 * Author: Alyssa I. Agard
 * License: MIT
 */

// ============================================
// MENU & SIDEBAR
// ============================================

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Layout Toolkit')
    .addItem('Open Sidebar', 'showSidebar')
    .addSeparator()
    .addSubMenu(
      DocumentApp.getUi().createMenu('Quick Insert')
        .addItem('Tri-Fold Brochure', 'insertTriFoldBrochure')
        .addItem('Bi-Fold Brochure', 'insertBiFoldBrochure')
        .addItem('Classic Newsletter', 'insertNewsletter')
        .addItem('Event Program', 'insertEventProgram')
        .addItem('Business Card', 'insertBusinessCard')
        .addItem('Event Flyer', 'insertFlyer')
    )
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Layout Toolkit');
  DocumentApp.getUi().showSidebar(html);
}

// ============================================
// BRAND PROFILE STORAGE (PropertiesService)
// ============================================

function saveBrandProfile(profileData) {
  var props = PropertiesService.getUserProperties();
  var profiles = JSON.parse(props.getProperty('lt_brandProfiles') || '{}');
  profiles[profileData.name] = profileData;
  props.setProperty('lt_brandProfiles', JSON.stringify(profiles));
  return { success: true };
}

function loadBrandProfiles() {
  var props = PropertiesService.getUserProperties();
  return JSON.parse(props.getProperty('lt_brandProfiles') || '{}');
}

function deleteBrandProfile(name) {
  var props = PropertiesService.getUserProperties();
  var profiles = JSON.parse(props.getProperty('lt_brandProfiles') || '{}');
  delete profiles[name];
  props.setProperty('lt_brandProfiles', JSON.stringify(profiles));
  return { success: true };
}

function getActiveBrand() {
  var props = PropertiesService.getUserProperties();
  var activeName = props.getProperty('lt_activeBrand');
  if (!activeName) return getDefaultBrand();
  var profiles = JSON.parse(props.getProperty('lt_brandProfiles') || '{}');
  return profiles[activeName] || getDefaultBrand();
}

function setActiveBrand(name) {
  PropertiesService.getUserProperties().setProperty('lt_activeBrand', name);
  return { success: true };
}

function getDefaultBrand() {
  return {
    name: 'Default',
    orgName: 'Your Organization',
    primaryColor: '#1B3A5C',
    accentColor: '#C0392B',
    headingFont: 'Georgia',
    bodyFont: 'Arial'
  };
}

// ============================================
// PAGE SETUP UTILITIES
// ============================================

function setPageOrientation(orientation) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  if (orientation === 'landscape') {
    body.setPageWidth(792);  // 11 inches
    body.setPageHeight(612); // 8.5 inches
  } else {
    body.setPageWidth(612);  // 8.5 inches
    body.setPageHeight(792); // 11 inches
  }
  return { success: true };
}

function setPageSize(size) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var sizes = {
    'letter':  { w: 612, h: 792 },
    'legal':   { w: 612, h: 1008 },
    'tabloid': { w: 792, h: 1224 },
    'a4':      { w: 595, h: 842 },
    'a5':      { w: 420, h: 595 },
    'card':    { w: 252, h: 144 }  // 3.5 x 2 inches
  };
  var s = sizes[size] || sizes['letter'];
  body.setPageWidth(s.w);
  body.setPageHeight(s.h);
  return { success: true };
}

function setMargins(top, bottom, left, right) {
  var body = DocumentApp.getActiveDocument().getBody();
  body.setMarginTop(top);
  body.setMarginBottom(bottom);
  body.setMarginLeft(left);
  body.setMarginRight(right);
  return { success: true };
}

// ============================================
// INSERT ELEMENTS
// ============================================

function insertDivider() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();
  var table = body.appendTable([['']]);
  table.setBorderWidth(0);
  var cell = table.getRow(0).getCell(0);
  cell.clear();
  var para = cell.appendParagraph('');
  para.appendHorizontalRule();
  // Set top border as colored line
  var style = {};
  style[DocumentApp.Attribute.BORDER_WIDTH] = 0;
  table.setAttributes(style);
  return { success: true };
}

function insertPullQuote(text) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();

  var table = body.appendTable([[text || '"Your pull quote text here."']]);
  table.setBorderWidth(1);
  table.setBorderColor(brand.primaryColor);
  var cell = table.getRow(0).getCell(0);
  cell.setPaddingTop(12);
  cell.setPaddingBottom(12);
  cell.setPaddingLeft(16);
  cell.setPaddingRight(16);

  var para = cell.getChild(0).asParagraph();
  para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  para.setFontFamily(brand.headingFont);
  para.setFontSize(14);
  para.setItalic(true);
  para.setForegroundColor(brand.primaryColor);

  return { success: true };
}

function insertCalloutBox(text) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();

  var table = body.appendTable([[text || 'Callout box content goes here.']]);
  table.setBorderWidth(2);
  table.setBorderColor(brand.accentColor);
  var cell = table.getRow(0).getCell(0);
  cell.setBackgroundColor('#F8F9FA');
  cell.setPaddingTop(10);
  cell.setPaddingBottom(10);
  cell.setPaddingLeft(12);
  cell.setPaddingRight(12);

  var para = cell.getChild(0).asParagraph();
  para.setFontFamily(brand.bodyFont);
  para.setFontSize(10);
  para.setForegroundColor('#333333');

  return { success: true };
}

function insertTextPlaceholder() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();

  var para = body.appendParagraph('[Your text here. Replace this placeholder with your content.]');
  para.setFontFamily(brand.bodyFont);
  para.setFontSize(11);
  para.setForegroundColor('#999999');
  para.setItalic(true);

  return { success: true };
}

// ============================================
// HELPER: Build a borderless layout table
// ============================================

function createLayoutTable(body, rows, cols, brand) {
  var data = [];
  for (var r = 0; r < rows; r++) {
    var row = [];
    for (var c = 0; c < cols; c++) {
      row.push('');
    }
    data.push(row);
  }
  var table = body.appendTable(data);
  table.setBorderWidth(0);
  table.setBorderColor('#FFFFFF');
  return table;
}

function styleCell(cell, brand, options) {
  options = options || {};
  cell.setPaddingTop(options.paddingTop || 8);
  cell.setPaddingBottom(options.paddingBottom || 8);
  cell.setPaddingLeft(options.paddingLeft || 10);
  cell.setPaddingRight(options.paddingRight || 10);
  if (options.bgColor) cell.setBackgroundColor(options.bgColor);
  if (options.width) cell.setWidth(options.width);
  if (options.verticalAlignment) {
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment[options.verticalAlignment]);
  }
}

function addStyledParagraph(cell, text, options) {
  options = options || {};
  // Clear the default empty paragraph if this is the first content
  if (cell.getNumChildren() === 1 && cell.getChild(0).asParagraph().getText() === '') {
    var para = cell.getChild(0).asParagraph();
  } else {
    var para = cell.appendParagraph('');
  }
  para.setText(text);
  if (options.font) para.setFontFamily(options.font);
  if (options.size) para.setFontSize(options.size);
  if (options.color) para.setForegroundColor(options.color);
  if (options.bold) para.setBold(options.bold);
  if (options.italic) para.setItalic(options.italic);
  if (options.alignment) para.setAlignment(DocumentApp.HorizontalAlignment[options.alignment]);
  if (options.heading) para.setHeading(DocumentApp.ParagraphHeading[options.heading]);
  if (options.spacing) para.setSpacingAfter(options.spacing);
  return para;
}

// ============================================
// TEMPLATE BUILDERS
// ============================================

function insertTriFoldBrochure() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();

  // Clear document
  body.clear();

  // Landscape, letter
  body.setPageWidth(792);
  body.setPageHeight(612);
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(36);
  body.setMarginRight(36);

  var colWidth = 228; // (792 - 72 margins - 36 gutter) / 3

  // === OUTSIDE (Page 1): Back Panel | Front Cover | Fold-in Flap ===
  var headerPara = body.appendParagraph('OUTSIDE (Print Side 1)');
  headerPara.setFontFamily(brand.bodyFont);
  headerPara.setFontSize(8);
  headerPara.setForegroundColor('#999999');
  headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var outside = createLayoutTable(body, 1, 3, brand);

  // Back Panel
  var backCell = outside.getRow(0).getCell(0);
  styleCell(backCell, brand, { width: colWidth, paddingTop: 20 });
  addStyledParagraph(backCell, brand.orgName, { font: brand.headingFont, size: 14, bold: true, color: brand.primaryColor, alignment: 'CENTER' });
  addStyledParagraph(backCell, '123 Main Street\nCity, State 12345\nphone: (555) 123-4567\nemail: info@example.org\nwww.example.org', { font: brand.bodyFont, size: 9, color: '#555555', alignment: 'CENTER', spacing: 6 });

  // Front Cover
  var frontCell = outside.getRow(0).getCell(1);
  styleCell(frontCell, brand, { width: colWidth, bgColor: brand.primaryColor, paddingTop: 60 });
  addStyledParagraph(frontCell, brand.orgName, { font: brand.headingFont, size: 20, bold: true, color: '#FFFFFF', alignment: 'CENTER', spacing: 12 });
  addStyledParagraph(frontCell, 'Your Brochure Title Here', { font: brand.headingFont, size: 14, color: '#FFFFFF', alignment: 'CENTER', spacing: 8 });
  addStyledParagraph(frontCell, 'A brief tagline or subtitle for your organization goes in this space.', { font: brand.bodyFont, size: 10, color: '#E0E0E0', alignment: 'CENTER' });

  // Fold-in Flap
  var flapCell = outside.getRow(0).getCell(2);
  styleCell(flapCell, brand, { width: colWidth, paddingTop: 20 });
  addStyledParagraph(flapCell, 'Quick Facts', { font: brand.headingFont, size: 13, bold: true, color: brand.primaryColor, alignment: 'LEFT', spacing: 8 });
  addStyledParagraph(flapCell, 'This panel folds inward and is the first thing readers see when they begin to open the brochure. Use it for a summary, key statistics, or a compelling introduction.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 6 });
  addStyledParagraph(flapCell, 'Highlight 1: Brief point about your mission or service.\n\nHighlight 2: A key statistic or achievement.\n\nHighlight 3: Contact or next-step information.', { font: brand.bodyFont, size: 9, color: '#555555' });

  // Page break
  body.appendPageBreak();

  // === INSIDE (Page 2): Panel 1 | Panel 2 | Panel 3 ===
  var insideHeader = body.appendParagraph('INSIDE (Print Side 2)');
  insideHeader.setFontFamily(brand.bodyFont);
  insideHeader.setFontSize(8);
  insideHeader.setForegroundColor('#999999');
  insideHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var inside = createLayoutTable(body, 1, 3, brand);

  // Panel 1
  var p1 = inside.getRow(0).getCell(0);
  styleCell(p1, brand, { width: colWidth, paddingTop: 16 });
  addStyledParagraph(p1, 'About Us', { font: brand.headingFont, size: 13, bold: true, color: brand.primaryColor, spacing: 6 });
  addStyledParagraph(p1, 'Provide a brief overview of your organization, its history, and its core mission. This panel typically introduces readers to who you are and why your work matters.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 6 });
  addStyledParagraph(p1, 'Add a second paragraph with additional context, recent milestones, or a brief story that illustrates your impact.', { font: brand.bodyFont, size: 9, color: '#444444' });

  // Panel 2
  var p2 = inside.getRow(0).getCell(1);
  styleCell(p2, brand, { width: colWidth, paddingTop: 16 });
  addStyledParagraph(p2, 'Our Programs', { font: brand.headingFont, size: 13, bold: true, color: brand.primaryColor, spacing: 6 });
  addStyledParagraph(p2, 'Program 1: Describe the first major program or service area. Include who it serves and what outcomes it achieves.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 6 });
  addStyledParagraph(p2, 'Program 2: Describe a second program area. Keep descriptions concise so readers can scan quickly.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 6 });
  addStyledParagraph(p2, 'Program 3: A third program or initiative. Consider including a brief statistic or quote to add credibility.', { font: brand.bodyFont, size: 9, color: '#444444' });

  // Panel 3
  var p3 = inside.getRow(0).getCell(2);
  styleCell(p3, brand, { width: colWidth, paddingTop: 16 });
  addStyledParagraph(p3, 'Get Involved', { font: brand.headingFont, size: 13, bold: true, color: brand.primaryColor, spacing: 6 });
  addStyledParagraph(p3, 'Volunteer: Describe volunteer opportunities and how to sign up.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 6 });
  addStyledParagraph(p3, 'Donate: Explain how contributions support your mission and where funds are directed.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 6 });
  addStyledParagraph(p3, 'Partner: Outline partnership opportunities for businesses and other organizations.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 6 });
  addStyledParagraph(p3, 'Visit: www.example.org\nCall: (555) 123-4567\nEmail: info@example.org', { font: brand.bodyFont, size: 9, bold: true, color: brand.accentColor });

  return { success: true, message: 'Tri-fold brochure created.' };
}

function insertBiFoldBrochure() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();

  body.clear();
  body.setPageWidth(792);
  body.setPageHeight(612);
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(36);
  body.setMarginRight(36);

  var colWidth = 352; // (792 - 72 - 16 gutter) / 2

  // === OUTSIDE ===
  var headerPara = body.appendParagraph('OUTSIDE (Print Side 1)');
  headerPara.setFontFamily(brand.bodyFont);
  headerPara.setFontSize(8);
  headerPara.setForegroundColor('#999999');
  headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var outside = createLayoutTable(body, 1, 2, brand);

  // Back cover
  var backCell = outside.getRow(0).getCell(0);
  styleCell(backCell, brand, { width: colWidth, paddingTop: 40 });
  addStyledParagraph(backCell, brand.orgName, { font: brand.headingFont, size: 16, bold: true, color: brand.primaryColor, alignment: 'CENTER', spacing: 10 });
  addStyledParagraph(backCell, '123 Main Street\nCity, State 12345\n(555) 123-4567\nwww.example.org', { font: brand.bodyFont, size: 10, color: '#555555', alignment: 'CENTER' });

  // Front cover
  var frontCell = outside.getRow(0).getCell(1);
  styleCell(frontCell, brand, { width: colWidth, bgColor: brand.primaryColor, paddingTop: 80 });
  addStyledParagraph(frontCell, brand.orgName, { font: brand.headingFont, size: 24, bold: true, color: '#FFFFFF', alignment: 'CENTER', spacing: 14 });
  addStyledParagraph(frontCell, 'Brochure Title', { font: brand.headingFont, size: 16, color: '#FFFFFF', alignment: 'CENTER', spacing: 8 });
  addStyledParagraph(frontCell, 'A subtitle or tagline for your brochure.', { font: brand.bodyFont, size: 11, color: '#E0E0E0', alignment: 'CENTER' });

  body.appendPageBreak();

  // === INSIDE ===
  var insideHeader = body.appendParagraph('INSIDE (Print Side 2)');
  insideHeader.setFontFamily(brand.bodyFont);
  insideHeader.setFontSize(8);
  insideHeader.setForegroundColor('#999999');
  insideHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var inside = createLayoutTable(body, 1, 2, brand);

  // Left inside
  var left = inside.getRow(0).getCell(0);
  styleCell(left, brand, { width: colWidth, paddingTop: 16 });
  addStyledParagraph(left, 'Our Mission', { font: brand.headingFont, size: 16, bold: true, color: brand.primaryColor, spacing: 8 });
  addStyledParagraph(left, 'Describe your organization\'s mission in one or two compelling paragraphs. Explain the problem you address and the approach you take.', { font: brand.bodyFont, size: 10, color: '#444444', spacing: 8 });
  addStyledParagraph(left, 'Add supporting details, recent accomplishments, or a brief story that illustrates the difference your organization makes in the community.', { font: brand.bodyFont, size: 10, color: '#444444' });

  // Right inside
  var right = inside.getRow(0).getCell(1);
  styleCell(right, brand, { width: colWidth, paddingTop: 16 });
  addStyledParagraph(right, 'Programs & Services', { font: brand.headingFont, size: 16, bold: true, color: brand.primaryColor, spacing: 8 });
  addStyledParagraph(right, 'Program One: Brief description of the first service area.', { font: brand.bodyFont, size: 10, color: '#444444', spacing: 6 });
  addStyledParagraph(right, 'Program Two: Brief description of the second service area.', { font: brand.bodyFont, size: 10, color: '#444444', spacing: 6 });
  addStyledParagraph(right, 'Program Three: Brief description of the third service area.', { font: brand.bodyFont, size: 10, color: '#444444', spacing: 8 });
  addStyledParagraph(right, 'Contact us to learn more:\ninfo@example.org | (555) 123-4567', { font: brand.bodyFont, size: 10, bold: true, color: brand.accentColor });

  return { success: true, message: 'Bi-fold brochure created.' };
}

function insertNewsletter() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();

  body.clear();
  body.setPageWidth(612);
  body.setPageHeight(792);
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(36);
  body.setMarginRight(36);

  // Masthead
  var mastheadTable = body.appendTable([['']]);
  mastheadTable.setBorderWidth(0);
  var mastheadCell = mastheadTable.getRow(0).getCell(0);
  mastheadCell.setBackgroundColor(brand.primaryColor);
  styleCell(mastheadCell, brand, { paddingTop: 16, paddingBottom: 16 });
  addStyledParagraph(mastheadCell, brand.orgName + ' Newsletter', { font: brand.headingFont, size: 22, bold: true, color: '#FFFFFF', alignment: 'CENTER', spacing: 4 });
  addStyledParagraph(mastheadCell, 'Volume 1, Issue 1  |  ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy'), { font: brand.bodyFont, size: 9, color: '#E0E0E0', alignment: 'CENTER' });

  body.appendParagraph('').setSpacingAfter(4);

  // Lead article (full width)
  var leadTitle = body.appendParagraph('Lead Article Headline Goes Here');
  leadTitle.setFontFamily(brand.headingFont);
  leadTitle.setFontSize(16);
  leadTitle.setBold(true);
  leadTitle.setForegroundColor(brand.primaryColor);
  leadTitle.setSpacingAfter(6);

  var leadBody = body.appendParagraph('Write your lead article here. This section spans the full width of the page and should contain your most important story or announcement. Use two to three paragraphs to introduce the topic, provide key details, and explain its significance to your audience.');
  leadBody.setFontFamily(brand.bodyFont);
  leadBody.setFontSize(10);
  leadBody.setForegroundColor('#333333');
  leadBody.setSpacingAfter(6);

  var leadBody2 = body.appendParagraph('Continue with supporting details, quotes from stakeholders, or additional context. A strong lead article draws readers into the rest of the newsletter.');
  leadBody2.setFontFamily(brand.bodyFont);
  leadBody2.setFontSize(10);
  leadBody2.setForegroundColor('#333333');
  leadBody2.setSpacingAfter(10);

  // Divider
  body.appendHorizontalRule();
  body.appendParagraph('').setSpacingAfter(4);

  // Two-column section
  var colWidth = 254; // (612 - 72 - 16) / 2
  var columns = createLayoutTable(body, 1, 2, brand);

  // Left column
  var leftCol = columns.getRow(0).getCell(0);
  styleCell(leftCol, brand, { width: colWidth });
  addStyledParagraph(leftCol, 'Secondary Story', { font: brand.headingFont, size: 13, bold: true, color: brand.primaryColor, spacing: 4 });
  addStyledParagraph(leftCol, 'Write a secondary article or announcement here. This two-column layout lets you present multiple stories side by side, keeping the newsletter scannable and visually balanced.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 6 });
  addStyledParagraph(leftCol, 'Add a second paragraph with follow-up details, dates, or calls to action.', { font: brand.bodyFont, size: 9, color: '#444444' });

  // Right column (sidebar)
  var rightCol = columns.getRow(0).getCell(1);
  styleCell(rightCol, brand, { width: colWidth, bgColor: '#F4F6F8' });
  addStyledParagraph(rightCol, 'Upcoming Events', { font: brand.headingFont, size: 13, bold: true, color: brand.accentColor, spacing: 4 });
  addStyledParagraph(rightCol, 'Jan 15: Event name and brief description.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 4 });
  addStyledParagraph(rightCol, 'Feb 10: Another event with a one-line summary.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 4 });
  addStyledParagraph(rightCol, 'Mar 22: Third event listed here.', { font: brand.bodyFont, size: 9, color: '#444444', spacing: 8 });
  addStyledParagraph(rightCol, 'Quick Links', { font: brand.headingFont, size: 13, bold: true, color: brand.accentColor, spacing: 4 });
  addStyledParagraph(rightCol, 'www.example.org\ninfo@example.org\n(555) 123-4567', { font: brand.bodyFont, size: 9, color: brand.primaryColor });

  // Footer
  body.appendParagraph('').setSpacingAfter(4);
  var footerTable = body.appendTable([['']]);
  footerTable.setBorderWidth(0);
  var footerCell = footerTable.getRow(0).getCell(0);
  footerCell.setBackgroundColor(brand.primaryColor);
  styleCell(footerCell, brand, { paddingTop: 8, paddingBottom: 8 });
  addStyledParagraph(footerCell, brand.orgName + '  |  123 Main Street, City, State 12345  |  (555) 123-4567  |  www.example.org', { font: brand.bodyFont, size: 8, color: '#FFFFFF', alignment: 'CENTER' });

  return { success: true, message: 'Newsletter created.' };
}

function insertEventProgram() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();

  body.clear();
  body.setPageWidth(612);
  body.setPageHeight(792);
  body.setMarginTop(48);
  body.setMarginBottom(48);
  body.setMarginLeft(54);
  body.setMarginRight(54);

  // Header
  var orgLine = body.appendParagraph(brand.orgName);
  orgLine.setFontFamily(brand.headingFont);
  orgLine.setFontSize(12);
  orgLine.setForegroundColor(brand.primaryColor);
  orgLine.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  orgLine.setSpacingAfter(4);

  var titleLine = body.appendParagraph('Event Program');
  titleLine.setFontFamily(brand.headingFont);
  titleLine.setFontSize(24);
  titleLine.setBold(true);
  titleLine.setForegroundColor(brand.primaryColor);
  titleLine.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titleLine.setSpacingAfter(4);

  var dateLine = body.appendParagraph('Saturday, January 15, 2026  |  10:00 AM - 4:00 PM');
  dateLine.setFontFamily(brand.bodyFont);
  dateLine.setFontSize(11);
  dateLine.setForegroundColor('#555555');
  dateLine.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  dateLine.setSpacingAfter(4);

  var venueLine = body.appendParagraph('Venue Name  |  123 Main Street, City, State');
  venueLine.setFontFamily(brand.bodyFont);
  venueLine.setFontSize(10);
  venueLine.setForegroundColor('#777777');
  venueLine.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  venueLine.setSpacingAfter(12);

  body.appendHorizontalRule();
  body.appendParagraph('').setSpacingAfter(6);

  // Schedule table
  var scheduleLabel = body.appendParagraph('Schedule');
  scheduleLabel.setFontFamily(brand.headingFont);
  scheduleLabel.setFontSize(14);
  scheduleLabel.setBold(true);
  scheduleLabel.setForegroundColor(brand.primaryColor);
  scheduleLabel.setSpacingAfter(6);

  var schedule = [
    ['10:00 AM', 'Registration & Welcome Coffee', 'Main Lobby'],
    ['10:30 AM', 'Opening Remarks', 'Grand Hall'],
    ['11:00 AM', 'Keynote Address: Speaker Name\n"Presentation Title Goes Here"', 'Grand Hall'],
    ['12:00 PM', 'Lunch Break', 'Dining Room'],
    ['1:00 PM',  'Panel Discussion: Topic Title\nPanelists: Name 1, Name 2, Name 3', 'Room A'],
    ['2:00 PM',  'Breakout Session A: Workshop Title\nFacilitator: Name', 'Room B'],
    ['2:00 PM',  'Breakout Session B: Workshop Title\nFacilitator: Name', 'Room C'],
    ['3:00 PM',  'Networking Break', 'Main Lobby'],
    ['3:30 PM',  'Closing Remarks & Awards', 'Grand Hall'],
    ['4:00 PM',  'Adjourn', '']
  ];

  var table = body.appendTable();
  // Header row
  var headerRow = table.appendTableRow();
  var headers = ['Time', 'Session', 'Location'];
  for (var h = 0; h < headers.length; h++) {
    var hCell = headerRow.appendTableCell(headers[h]);
    hCell.setBackgroundColor(brand.primaryColor);
    hCell.setPaddingTop(6);
    hCell.setPaddingBottom(6);
    hCell.setPaddingLeft(8);
    hCell.setPaddingRight(8);
    var hPara = hCell.getChild(0).asParagraph();
    hPara.setFontFamily(brand.bodyFont);
    hPara.setFontSize(9);
    hPara.setBold(true);
    hPara.setForegroundColor('#FFFFFF');
  }

  // Data rows
  for (var i = 0; i < schedule.length; i++) {
    var row = table.appendTableRow();
    for (var j = 0; j < 3; j++) {
      var cell = row.appendTableCell(schedule[i][j]);
      cell.setBackgroundColor(i % 2 === 0 ? '#F8F9FA' : '#FFFFFF');
      cell.setPaddingTop(5);
      cell.setPaddingBottom(5);
      cell.setPaddingLeft(8);
      cell.setPaddingRight(8);
      var cellPara = cell.getChild(0).asParagraph();
      cellPara.setFontFamily(brand.bodyFont);
      cellPara.setFontSize(9);
      cellPara.setForegroundColor('#333333');
      if (j === 0) cellPara.setBold(true);
    }
  }

  // Remove the initial empty row that appendTable creates
  table.removeRow(0);
  table.setBorderColor('#DEE2E6');
  table.setBorderWidth(1);

  // Set column widths
  for (var r = 0; r < table.getNumRows(); r++) {
    table.getRow(r).getCell(0).setWidth(72);
    table.getRow(r).getCell(2).setWidth(90);
  }

  body.appendParagraph('').setSpacingAfter(10);

  // Notes section
  var notesTitle = body.appendParagraph('Notes');
  notesTitle.setFontFamily(brand.headingFont);
  notesTitle.setFontSize(12);
  notesTitle.setBold(true);
  notesTitle.setForegroundColor(brand.primaryColor);
  notesTitle.setSpacingAfter(4);

  var notes = body.appendParagraph('Wi-Fi Network: EventGuest  |  Password: welcome2026\nAccessibility: All event spaces are wheelchair accessible.\nParking: Free parking available in Lot B.');
  notes.setFontFamily(brand.bodyFont);
  notes.setFontSize(9);
  notes.setForegroundColor('#555555');

  return { success: true, message: 'Event program created.' };
}

function insertBusinessCard() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();

  body.clear();
  // Standard business card: 3.5 x 2 in
  body.setPageWidth(252);
  body.setPageHeight(144);
  body.setMarginTop(14);
  body.setMarginBottom(14);
  body.setMarginLeft(18);
  body.setMarginRight(18);

  var name = body.appendParagraph('Your Name');
  name.setFontFamily(brand.headingFont);
  name.setFontSize(13);
  name.setBold(true);
  name.setForegroundColor(brand.primaryColor);
  name.setSpacingAfter(2);

  var title = body.appendParagraph('Title / Position');
  title.setFontFamily(brand.bodyFont);
  title.setFontSize(8);
  title.setForegroundColor(brand.accentColor);
  title.setSpacingAfter(6);

  var org = body.appendParagraph(brand.orgName);
  org.setFontFamily(brand.bodyFont);
  org.setFontSize(8);
  org.setBold(true);
  org.setForegroundColor('#333333');
  org.setSpacingAfter(4);

  var contact = body.appendParagraph('(555) 123-4567\nemail@example.org\nwww.example.org');
  contact.setFontFamily(brand.bodyFont);
  contact.setFontSize(7);
  contact.setForegroundColor('#555555');

  return { success: true, message: 'Business card created.' };
}

function insertFlyer() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var brand = getActiveBrand();

  body.clear();
  body.setPageWidth(612);
  body.setPageHeight(792);
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(48);
  body.setMarginRight(48);

  // Hero banner
  var heroTable = body.appendTable([['']]);
  heroTable.setBorderWidth(0);
  var heroCell = heroTable.getRow(0).getCell(0);
  heroCell.setBackgroundColor(brand.primaryColor);
  styleCell(heroCell, brand, { paddingTop: 48, paddingBottom: 48 });
  addStyledParagraph(heroCell, 'EVENT TITLE', { font: brand.headingFont, size: 28, bold: true, color: '#FFFFFF', alignment: 'CENTER', spacing: 8 });
  addStyledParagraph(heroCell, 'A brief subtitle or tagline goes here', { font: brand.bodyFont, size: 13, color: '#E0E0E0', alignment: 'CENTER' });

  body.appendParagraph('').setSpacingAfter(8);

  // Details bar
  var detailsTable = body.appendTable([['']]);
  detailsTable.setBorderWidth(0);
  var detailsCell = detailsTable.getRow(0).getCell(0);
  detailsCell.setBackgroundColor(brand.accentColor);
  styleCell(detailsCell, brand, { paddingTop: 10, paddingBottom: 10 });
  addStyledParagraph(detailsCell, 'Saturday, January 15, 2026  |  10:00 AM - 2:00 PM  |  Venue Name', { font: brand.bodyFont, size: 11, bold: true, color: '#FFFFFF', alignment: 'CENTER' });

  body.appendParagraph('').setSpacingAfter(8);

  // Body text
  var bodyTitle = body.appendParagraph('About This Event');
  bodyTitle.setFontFamily(brand.headingFont);
  bodyTitle.setFontSize(16);
  bodyTitle.setBold(true);
  bodyTitle.setForegroundColor(brand.primaryColor);
  bodyTitle.setSpacingAfter(6);

  var bodyText = body.appendParagraph('Write a compelling description of your event here. Explain what attendees can expect, who the event is for, and why they should attend. Keep this section to two or three paragraphs at most.');
  bodyText.setFontFamily(brand.bodyFont);
  bodyText.setFontSize(11);
  bodyText.setForegroundColor('#333333');
  bodyText.setSpacingAfter(6);

  var bodyText2 = body.appendParagraph('Add supporting details such as featured speakers, activities, food, parking, or any special instructions for attendees.');
  bodyText2.setFontFamily(brand.bodyFont);
  bodyText2.setFontSize(11);
  bodyText2.setForegroundColor('#333333');
  bodyText2.setSpacingAfter(12);

  // Highlights (2-column)
  var highlights = createLayoutTable(body, 1, 2, brand);
  var hlLeft = highlights.getRow(0).getCell(0);
  styleCell(hlLeft, brand, { width: 240 });
  addStyledParagraph(hlLeft, 'What to Expect', { font: brand.headingFont, size: 12, bold: true, color: brand.primaryColor, spacing: 4 });
  addStyledParagraph(hlLeft, 'Feature or activity #1\nFeature or activity #2\nFeature or activity #3', { font: brand.bodyFont, size: 10, color: '#444444' });

  var hlRight = highlights.getRow(0).getCell(1);
  styleCell(hlRight, brand, { width: 240 });
  addStyledParagraph(hlRight, 'Details', { font: brand.headingFont, size: 12, bold: true, color: brand.primaryColor, spacing: 4 });
  addStyledParagraph(hlRight, 'Cost: Free / $XX\nDress Code: Casual\nParking: Available on site', { font: brand.bodyFont, size: 10, color: '#444444' });

  body.appendParagraph('').setSpacingAfter(8);

  // CTA footer
  var ctaTable = body.appendTable([['']]);
  ctaTable.setBorderWidth(0);
  var ctaCell = ctaTable.getRow(0).getCell(0);
  ctaCell.setBackgroundColor(brand.accentColor);
  styleCell(ctaCell, brand, { paddingTop: 14, paddingBottom: 14 });
  addStyledParagraph(ctaCell, 'REGISTER NOW', { font: brand.headingFont, size: 16, bold: true, color: '#FFFFFF', alignment: 'CENTER', spacing: 4 });
  addStyledParagraph(ctaCell, 'www.example.org/register  |  (555) 123-4567  |  info@example.org', { font: brand.bodyFont, size: 10, color: '#FFFFFF', alignment: 'CENTER' });

  return { success: true, message: 'Event flyer created.' };
}
