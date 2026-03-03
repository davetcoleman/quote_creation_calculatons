// Author: Dave Coleman and ChatGPT ;-) 

function runSOWSetup() {
  var ui = DocumentApp.getUi();

  // Set to true to skip all prompts and use test values
  var DEBUG_MODE = true;

  var customer, sowNumber, folderHint;

  if (DEBUG_MODE) {
    customer = 'CleanBotix';
    sowNumber = '99';
    folderHint = ' - Testing';
  } else {
    // Prompt for customer name
    var nameResponse = ui.prompt(
      'Automatic SOW Setup v3',
      'What is the EXACT name of the existing customer project this proposal is for?\n\nWe will use this exact spelling to find the existing folder in our file structure.',
      ui.ButtonSet.OK_CANCEL
    );
    if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
    customer = nameResponse.getResponseText().trim();
  }

  // Folder IDs to search in
  var opportunitiesFolderId = '1VpiL6v3dcLNaOUG882cMv3k-C-5FXEsn';
  var projectsFolderId = '0ALDgMGsigPHbUk9PVA';
  var folderSources = [
    { id: opportunitiesFolderId, label: "Opportunities" },
    { id: projectsFolderId, label: "Customers" }
  ];

  // Try to find customer folder in either location
  var customerFolder = null;
  for (var i = 0; i < folderSources.length; i++) {
    var parent = DriveApp.getFolderById(folderSources[i].id);
    var matches = parent.getFoldersByName(customer);
    if (matches.hasNext()) {
      customerFolder = matches.next();
      break;
    }
  }

  if (!customerFolder) {
    var searchedPaths = folderSources.map(function(source) {
      return source.label + " (ID: " + source.id + ")";
    }).join('\n');
    ui.alert(
      "Customer Folder Not Found",
      "Unable to find the existing customer name '" + customer + "' in GDrive.\n\nPaths searched:\n" + searchedPaths + "\n\nPlease check spelling and try again.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Look for "CUSTOMER_NAME Legal & Sales" subfolder
  var legalSalesFolderName = customer + ' Legal & Sales';
  var legalSalesFolders = customerFolder.getFoldersByName(legalSalesFolderName);
 if (!legalSalesFolders.hasNext()) {
    // Build full GDrive path by walking up the tree
    var pathParts = [];
    var current = customerFolder;
    while (true) {
      pathParts.unshift(current.getName());
      var parents = current.getParents();
      if (!parents.hasNext()) break;
      current = parents.next();
    }
    var fullPath = pathParts.join(" > ") + " > " + legalSalesFolderName;

    ui.alert(
      "Missing Subfolder",
      "We found the customer folder but could not locate the expected subfolder: '" + legalSalesFolderName + "'.\n\n" +
      "Full path searched:\n" + fullPath + "\n\nPlease ensure this subfolder exists before running this script.",
      ui.ButtonSet.OK
    );
    return;
  }
  var legalSalesFolder = legalSalesFolders.next();

  if (!DEBUG_MODE) {
    // Prompt for SOW number
    var sowResponse = ui.prompt(
      'SOW Number',
      'Enter the SOW number.\n\nThis should be an integer and the next SOW in sequence. If this is the first SOW, set it to 1. You can also add a letter for special circumstances, like "2b"',
      ui.ButtonSet.OK_CANCEL
    );
    if (sowResponse.getSelectedButton() !== ui.Button.OK) return;
    sowNumber = sowResponse.getResponseText().trim();
    //if (!/^\d+$/.test(sowNumber)) {
    //  ui.alert("Invalid SOW Number", "Please enter a valid integer (e.g., 1, 2, 3).", ui.ButtonSet.OK);
    //  return;
    //}

    // Prompt for SOW folder name hint
    var hintResponse = ui.prompt(
      'SOW Number Short Hint',
      'Enter a short description for the SOW folder to better explain each phase of the engagement. Examples: "Initial Quote", "Solution Study & Digital Twin", "Annual Renewal", etc:',
      ui.ButtonSet.OK_CANCEL
    );
    if (hintResponse.getSelectedButton() !== ui.Button.OK) return;
    folderHint = hintResponse.getResponseText().trim();
    if (folderHint.length > 0) {
      folderHint = " - " + folderHint;
    }
  }

  // Show "please wait" sidebar
  var logoHTML = '<img src="https://picknik.ai/assets/press-kit/picknik/PickNik_black_3_bullet_point.png" style="width:100%;max-width:280px;margin-bottom:10px;" />';
  var loadingHTML = logoHTML + '<h1>Sales Quote Setup</h1>Custom PickNik scripts <br /><br />THIS COULD TAKE A MINUTE<br /><br /><i>Hold on to your butts</i> :-)';
  ui.showSidebar(HtmlService.createHtmlOutput(loadingHTML).setTitle("Running...").setWidth(300));

  // Create SOW subfolder inside "Legal & Sales"
  var sowFolderName = "SOW " + sowNumber + folderHint;
  var sowFolder = legalSalesFolder.createFolder(sowFolderName);

  var titleHTML = logoHTML + '<h1>' + customer + ' SOW ' + sowNumber + ' Created</h1>';
  var summaryHTML = '<p><b>Created Folder:</b> <a href="' + sowFolder.getUrl() + '" target="_blank">' + sowFolderName + '</a></p>';

  // Template file IDs
  var templateFileIds = [
    '1UcLVEmYfpEy5-lJE41BiSODtrW34H3LgMCS2TuamUak', // Sales Quote GDoc
    '1fvaPH0peqbP5oomqK3sVp_S2S9EeDpmRbaPrgCYOrE4'  // Budget Spreadsheet
  ];

  // Copy and personalize files, store info for sidebar
  var fileLinksHTML = '<ul>';
  var closingTag = '</ul>';
  var sowDocId = null; // Will hold the GDoc copy's ID for the section selector
  templateFileIds.forEach(function(templateId) {
    var originalFile = DriveApp.getFileById(templateId);
    var originalName = originalFile.getName();

    // Replace placeholders in file name (handle both space and underscore variants)
    var newName = originalName
      .replace("COMPANY_NAME", customer)
      .replace("COMPANY NAME", customer)
      .replace("SOW X", "SOW " + sowNumber);

    var copiedFile = originalFile.makeCopy(newName, sowFolder);

    // Add file to HTML summary with a label based on file type
    var mimeType = copiedFile.getMimeType();
    var fileLabel = (mimeType === MimeType.GOOGLE_DOCS) ? 'Quote Document' : 'Budget Spreadsheet';
    fileLinksHTML += '<li><b>' + fileLabel + ':</b> <a href="' + copiedFile.getUrl() + '" target="_blank">' + newName + '</a></li>';

    // Show intermediate progress
    var html = HtmlService.createHtmlOutput(titleHTML + summaryHTML + fileLinksHTML + closingTag)
                      .setTitle("Progress")
                      .setWidth(300);
    ui.showSidebar(html);

    // Try find-replace inside Docs only
    if (mimeType === MimeType.GOOGLE_DOCS) {
      var doc = DocumentApp.openById(copiedFile.getId());
      var body = doc.getBody();

      // Replace text in main body
      body.replaceText('COMPANY_NAME', customer);
      body.replaceText('SOW X', 'SOW ' + sowNumber);

      // Add today's date
      var today = new Date();
      var todayString = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM d, yyyy");
      body.replaceText('\\[TODAYS_DATE\\]', todayString);

      // Add Quote Number
      body.replaceText('\\[QUOTE_NUM\\]', sowNumber);

      doc.saveAndClose();

      // Track the GDoc ID for the section selector sidebar
      sowDocId = copiedFile.getId();
    }
  });
  fileLinksHTML += '</ul>';

  // Show interactive section selector sidebar (or simple confirmation if no GDoc)
  if (sowDocId) {
    var sidebarContent = buildSectionSidebar(sowDocId, titleHTML, summaryHTML, fileLinksHTML);
    var html = HtmlService.createHtmlOutput(sidebarContent)
                          .setTitle("SOW Section Selector")
                          .setWidth(300);
    ui.showSidebar(html);
  } else {
    var html = HtmlService.createHtmlOutput(titleHTML + summaryHTML + fileLinksHTML)
                          .setTitle("Confirmation")
                          .setWidth(300);
    ui.showSidebar(html);
  }
}


/**
 * Returns an array of {index, text} for each HEADING1 paragraph in the given doc.
 */
function getH1Sections(docId) {
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  var numChildren = body.getNumChildren();
  var sections = [];

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = child.asParagraph();
      if (paragraph.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
        sections.push({
          index: i,
          text: paragraph.getText()
        });
      }
    }
  }

  return sections;
}

/**
 * Deletes sections from the document. Each "section" is everything from an H1
 * heading to just before the next H1 heading (or end of document).
 * Matches by heading text, re-scanning after each deletion to handle index shifts.
 */
function deleteSections(docId, sectionTexts) {
  if (!sectionTexts || sectionTexts.length === 0) return;

  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();

  for (var s = 0; s < sectionTexts.length; s++) {
    var targetText = sectionTexts[s];
    var numChildren = body.getNumChildren();

    // Find all H1 indices and locate the target
    var allH1Indices = [];
    var targetIndex = -1;
    for (var i = 0; i < numChildren; i++) {
      var child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        var paragraph = child.asParagraph();
        if (paragraph.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
          allH1Indices.push(i);
          if (paragraph.getText() === targetText && targetIndex === -1) {
            targetIndex = i;
          }
        }
      }
    }

    if (targetIndex === -1) continue; // heading not found, skip

    // Determine section end: next H1 after targetIndex, or end of doc
    var endIndex = numChildren;
    for (var h = 0; h < allH1Indices.length; h++) {
      if (allH1Indices[h] > targetIndex) {
        endIndex = allH1Indices[h];
        break;
      }
    }

    // Delete from endIndex-1 down to targetIndex (reverse to avoid shifting)
    for (var e = endIndex - 1; e >= targetIndex; e--) {
      if (body.getNumChildren() > 1) {
        body.removeChild(body.getChild(e));
      } else {
        body.getChild(0).asParagraph().clear();
      }
    }
  }

  doc.saveAndClose();
}

/**
 * Builds the interactive HTML sidebar with H1 section checkboxes.
 */
function buildSectionSidebar(docId, titleHTML, summaryHTML, fileLinksHTML) {
  return '<html><head>'
    + '<style>'
    + '  body { font-family: Arial, sans-serif; font-size: 13px; padding: 10px; }'
    + '  .section-list { list-style: none; padding: 0; }'
    + '  .section-list li { padding: 4px 0; }'
    + '  .section-list label { cursor: pointer; }'
    + '  .btn { padding: 8px 16px; margin: 4px; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; }'
    + '  .btn-primary { background-color: #4285f4; color: white; }'
    + '  .btn-secondary { background-color: #e0e0e0; color: #333; }'
    + '  .btn:disabled { opacity: 0.5; cursor: not-allowed; }'
    + '  #status { margin-top: 10px; font-style: italic; color: #666; }'
    + '  .select-controls { margin: 8px 0; font-size: 12px; }'
    + '  .select-controls a { color: #4285f4; cursor: pointer; margin-right: 12px; }'
    + '</style>'
    + '</head><body>'
    + titleHTML
    + summaryHTML
    + fileLinksHTML
    + '<hr/>'
    + '<h3>Select Sections to Keep</h3>'
    + '<p>Uncheck sections you want to <b>remove</b> from the SOW document:</p>'
    + '<div class="select-controls">'
    + '  <a onclick="toggleAll(true)">Select All</a>'
    + '  <a onclick="toggleAll(false)">Deselect All</a>'
    + '</div>'
    + '<div id="loading">Loading sections...</div>'
    + '<ul id="sectionList" class="section-list" style="display:none;"></ul>'
    + '<div id="buttons" style="display:none;">'
    + '  <button class="btn btn-primary" id="applyBtn" onclick="applySections()">Apply</button>'
    + '  <button class="btn btn-secondary" onclick="google.script.host.close()">Skip</button>'
    + '</div>'
    + '<div id="status"></div>'
    + '<script>'
    + 'var DOC_ID = "' + docId + '";'
    + 'google.script.run'
    + '  .withSuccessHandler(function(sections) {'
    + '    var list = document.getElementById("sectionList");'
    + '    var loading = document.getElementById("loading");'
    + '    if (!sections || sections.length === 0) {'
    + '      loading.textContent = "No H1 sections found in the document.";'
    + '      return;'
    + '    }'
    + '    loading.style.display = "none";'
    + '    list.style.display = "block";'
    + '    document.getElementById("buttons").style.display = "block";'
    + '    for (var i = 0; i < sections.length; i++) {'
    + '      var li = document.createElement("li");'
    + '      var label = document.createElement("label");'
    + '      var cb = document.createElement("input");'
    + '      cb.type = "checkbox";'
    + '      cb.checked = true;'
    + '      cb.value = sections[i].text;'
    + '      cb.name = "section";'
    + '      label.appendChild(cb);'
    + '      label.appendChild(document.createTextNode(" " + sections[i].text));'
    + '      li.appendChild(label);'
    + '      list.appendChild(li);'
    + '    }'
    + '  })'
    + '  .withFailureHandler(function(err) {'
    + '    document.getElementById("loading").textContent = "Error: " + err.message;'
    + '  })'
    + '  .getH1Sections(DOC_ID);'
    + 'function toggleAll(checked) {'
    + '  var cbs = document.querySelectorAll("input[name=section]");'
    + '  for (var i = 0; i < cbs.length; i++) cbs[i].checked = checked;'
    + '}'
    + 'function applySections() {'
    + '  var cbs = document.querySelectorAll("input[name=section]");'
    + '  var unchecked = [];'
    + '  for (var i = 0; i < cbs.length; i++) {'
    + '    if (!cbs[i].checked) unchecked.push(cbs[i].value);'
    + '  }'
    + '  if (unchecked.length === 0) {'
    + '    document.getElementById("status").textContent = "All sections kept. No changes made.";'
    + '    return;'
    + '  }'
    + '  document.getElementById("applyBtn").disabled = true;'
    + '  document.getElementById("status").textContent = "Removing " + unchecked.length + " section(s)...";'
    + '  google.script.run'
    + '    .withSuccessHandler(function() {'
    + '      document.getElementById("status").textContent = "Done! " + unchecked.length + " section(s) removed.";'
    + '      document.getElementById("applyBtn").disabled = false;'
    + '    })'
    + '    .withFailureHandler(function(err) {'
    + '      document.getElementById("status").textContent = "Error: " + err.message;'
    + '      document.getElementById("applyBtn").disabled = false;'
    + '    })'
    + '    .deleteSections(DOC_ID, unchecked);'
    + '}'
    + '</script>'
    + '</body></html>';
}


function totalUpTheEstimate() {

  var doc = DocumentApp.getActiveDocument();
  var ui = DocumentApp.getUi(); 
  var sidebarTitle = "Proposal Estimate Summary";
  ui.showSidebar(HtmlService.createHtmlOutput("<i>Calculating...</i><br/><br/>" + 
      "Checking every element in the document for EQUATION types and adding them up.")
      .setTitle(sidebarTitle).setWidth(300));

  // Get the body of the document
  var body = doc.getBody()
  
  var quoteTotal = 0;
  var mathCheckHTML = '<br />Here are the raw estimate numbers detected:<br />';

  // Define the search parameters.
  var searchType = DocumentApp.ElementType.EQUATION;
  var searchResult = null;

  // Search until an Equation is found.
  while (searchResult = body.findElement(searchType, searchResult)) {
    var element = searchResult.getElement();

    // Get the text 
    var equation_text = element.getText();
      
    // If found, add it to that list
    if (equation_text) {

      // Highlight text yellow, set font black, make bold
      element.asText().setForegroundColor('#000000');
      element.asText().setBold(true);
      element.asText().setFontSize(16);
      
      var estimate_value = parseFloat(equation_text);
      if (isNaN(estimate_value))
      {
        ui.alert("Error","Unable to convert estimate field #" + j + 
                          " of value '" + equation_text + "' to float. Estimate summary may be wrong",
                          ui.ButtonSet.OK);
          
        mathCheckHTML = mathCheckHTML + "FAILED Equation Value " + j + ": " + equation_text + " day(s)<br />";
        
        // Highlight element red
        element.asText().setBackgroundColor('#ff0000');
      } else {
        mathCheckHTML = mathCheckHTML + estimate_value + "<br />";
        quoteTotal += estimate_value;
        
        // Highlight element green
        element.asText().setBackgroundColor('#02e89d');
      }
    }
  }
  
  var hourlyRate = 235;
  var dollarEstimate = hourlyRate * 8 * quoteTotal;
  var formattedTotalEstimate = dollarEstimate.toLocaleString();
  var contengencyPercent = 1.3;
  var totalWithContingency = dollarEstimate * contengencyPercent;
  var formattedWithContingency = totalWithContingency.toLocaleString();

  // Show math total in side bar
  var logoHTML = '<img src="https://picknik.ai/assets/press-kit/picknik/PickNik_black_3_bullet_point.png" style="width:100%;max-width:280px;margin-bottom:10px;" />';
  var titleHTML = logoHTML + '<h1>Total Proposal Cost</h1>' +
                  '<i>Custom PickNik scripts</i> <br />';
  var fullHTML = titleHTML + mathCheckHTML +
                 "<br /><h3>Total Quote:</h3><h2>" + quoteTotal + " days of effort</h2>" +
                 "<b>For Quick Reference Only:</b></br>" +
                 "$" + hourlyRate + " hourly rate<br/>" +
                 "Cost estimate: $" + formattedTotalEstimate + "<br />" +
                 "With 30% Contingency: <b>$" + formattedWithContingency + "</b><br /><br />" +
                 "<i>Don't forget to manually update the tables with the new total.</i>";

  var html = HtmlService.createHtmlOutput(fullHTML).setTitle(sidebarTitle).setWidth(300);
  ui.showSidebar(html);
}

function hideEquations()
{
  enableEquations(false);
}

function showEquations()
{
 enableEquations(true);
}

function enableEquations(showThem)
{
  var body = DocumentApp.getActiveDocument().getBody();

  // Define the search parameters.
  var searchType = DocumentApp.ElementType.EQUATION;
  var searchResult = null;

  // Search until an Equation is found.
  while (searchResult = body.findElement(searchType, searchResult)) {
    var element = searchResult.getElement();

    // Get the text 
    var equation_text = element.getText();

    // If found
    if (equation_text) {    
      if (showThem) {
        // Highlight text yellow, set font black, make bold
        element.asText().setBackgroundColor('#ffff00');
        element.asText().setForegroundColor('#000000');
        element.asText().setBold(true);
        element.asText().setFontSize(16);
      } else {
        // Highlight text white, set font white, make not bold
        element.asText().setBackgroundColor('#ffffff');
        element.asText().setForegroundColor('#ffffff');
        element.asText().setBold(false);
        element.asText().setFontSize(1);
      }
    }
  }
}

/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 */
function onOpen(e) {
  // Add custom menus to the spreadsheet.
  var ui = DocumentApp.getUi(); 

  ui.createMenu('PickNik')
      //.addItem('Sum Equations to get Total Estimate', 'totalUpTheEstimate')
      .addItem('Hide All Equations (White & Small Font)', 'hideEquations')
      .addItem('Show All Equations & Add Them (Highlighted & Large Font)', 'totalUpTheEstimate')
      .addItem('Copy sales proposal to specific customer', 'runSOWSetup')
      .addToUi();

  //ui.alert("PickNik Automated GDoc","This is a fancy quote GDoc that uses App Scripts to sum up Equations to get a quote total. See the //'PickNik' menu bar for capabilities.", ui.ButtonSet.OK);
}
