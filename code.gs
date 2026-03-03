// Author: Dave Coleman and ChatGPT ;-) 

function runSOWSetup() {
  var ui = DocumentApp.getUi();

  // Prompt for customer name
  var nameResponse = ui.prompt(
    'Automatic SOW Setup v3',
    'What is the EXACT name of the existing customer project this proposal is for?\n\nWe will use this exact spelling to find the existing folder in our file structure.',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  var customer = nameResponse.getResponseText().trim();

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

  // Prompt for SOW number
  var sowResponse = ui.prompt(
    'SOW Number',
    'Enter the SOW number.\n\nThis should be an integer and the next SOW in sequence. If this is the first SOW, set it to 1. You can also add a letter for special circumstances, like "2b"',
    ui.ButtonSet.OK_CANCEL
  );
  if (sowResponse.getSelectedButton() !== ui.Button.OK) return;
  var sowNumber = sowResponse.getResponseText().trim();
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
  var folderHint = hintResponse.getResponseText().trim();
  if (folderHint.length > 0) {
    folderHint = " - " + folderHint;
  }

  // Show "please wait" sidebar
  var loadingHTML = '<h1>Sales Quote Setup</h1>Custom PickNik scripts <br /><br />THIS COULD TAKE A MINUTE<br /><br /><i>Hold on to your butts</i> :-)';
  ui.showSidebar(HtmlService.createHtmlOutput(loadingHTML).setTitle("Running...").setWidth(300));

  // Create SOW subfolder inside "Legal & Sales"
  var sowFolderName = "SOW " + sowNumber + folderHint;
  var sowFolder = legalSalesFolder.createFolder(sowFolderName);

  var titleHTML = '<h1>' + customer + ' SOW ' + sowNumber + ' Created</h1>';
  var summaryHTML = '<p><b>Created Folder:</b> <a href="' + sowFolder.getUrl() + '" target="_blank">' + sowFolderName + '</a></p>';

  // Template file IDs
  var templateFileIds = [
    '1UcLVEmYfpEy5-lJE41BiSODtrW34H3LgMCS2TuamUak', // Sales Quote GDoc
    '1fvaPH0peqbP5oomqK3sVp_S2S9EeDpmRbaPrgCYOrE4'  // Budget Spreadsheet
  ];

  // Copy and personalize files, store info for sidebar
  var fileLinksHTML = '<ul>';
  var closingTag = '</ul>';
  templateFileIds.forEach(function(templateId) {
    var originalFile = DriveApp.getFileById(templateId);
    var originalName = originalFile.getName();

    // Replace placeholders in file name
    var newName = originalName
      .replace("COMPANY_NAME", customer)
      .replace("SOW X", "SOW " + sowNumber);

    var copiedFile = originalFile.makeCopy(newName, sowFolder);

    // Add file to HTML summary
    fileLinksHTML += '<li><a href="' + copiedFile.getUrl() + '" target="_blank">' + newName + '</a></li>';

    // Show intermediate progress
    var html = HtmlService.createHtmlOutput(titleHTML + summaryHTML + fileLinksHTML + closingTag)
                      .setTitle("Progress")
                      .setWidth(300);
    ui.showSidebar(html);

    // Get MIME type
    var mimeType = copiedFile.getMimeType();

    // Try find-replace inside Docs only
    // application/vnd.google-apps.document
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
    }
  });
  fileLinksHTML += '</ul>';

  // Final sidebar confirmation
  
  var html = HtmlService.createHtmlOutput(titleHTML + summaryHTML + fileLinksHTML)
                        .setTitle("Confirmation")
                        .setWidth(300);
  ui.showSidebar(html);
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
  var titleHTML = '<h1>Total Proposal Cost</h1>' + 
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
