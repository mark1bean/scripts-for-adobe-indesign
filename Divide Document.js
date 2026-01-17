/**
 * @file Divide Document.js
 *
 * Outputs new Indesign documents comprising
 * sequences of pages of the active document.
 *
 * Makes attempt to cleanly cut text threads but
 * some complex cases may not work properly.
 *
 * Keywords:
 *   #Extract Pages
 *   #Split Document
 *
 * @author m1b
 * @version 2026-01-17
 * @discussion https://community.adobe.com/t5/indesign-discussions/split-id-document-can-anybody-change-this-script/m-p/15241392#M618671
 */
function main() {

    var settings = {

        /** divide by "pages" or "spreads" */
        divideBy: "pages",

        /** number of pages/spreads in each extracted document */
        divisionCount: 1,

        /** path inside the parent folder and file name with template tokens */
        fileNameTemplate: "Output/{docName}_{startPageName}-{endPageName}.indd",

        /** whether to keep the divided documents section numbering start */
        keepPageNumbering: true,

        /** whether to reveal the folder in the OS afterwards */
        revealFolder: true,

        /** whether to show the UI */
        showUI: true,

    };

    if (0 === app.documents.length)
        return alert("Please open a document and try again.");

    var doc = settings.doc = app.activeDocument;

    if (!doc.saved)
        return alert("Please save the document first.");

    if (settings.showUI) {

        var result = ui(settings);

        if (2 === result)
            // user cancelled
            return;

    }

    var allowPageShuffle = doc.documentPreferences.allowPageShuffle;
    var preserveLayoutWhenShuffling = doc.documentPreferences.preserveLayoutWhenShuffling;

    doc.documentPreferences.properties = {
        allowPageShuffle: false,
        preserveLayoutWhenShuffling: true,
    };

    var docName = doc.name.replace(/\.[^\.]+$/, '');
    var filePath = doc.filePath;
    var f;

    var divideBy = settings.divideBy === "pages" ? "pages" : "spreads";

    var totalTargets = doc[divideBy].length;
    var len = Math.ceil(totalTargets / settings.divisionCount);

    for (var i = 0; i < len; i++) {

        var start = i * settings.divisionCount;
        var end = Math.min(start + settings.divisionCount - 1, totalTargets - 1);

        // open a copy of the document
        var newDoc = app.open(doc.fullName, true, OpenOptions.OPEN_COPY);

        // the pages/spreads we want
        var targets = newDoc[divideBy].itemByRange(end, start).getElements();

        var firstTarget = targets[0];
        var lastTarget = targets[targets.length - 1];
        var firstPage = firstTarget instanceof Spread ? firstTarget.pages.firstItem() : firstTarget;
        var lastPage = lastTarget instanceof Spread ? lastTarget.pages.lastItem() : lastTarget;

        var firstPageOffset = firstPage.documentOffset;

        // cut threads to these pages/spreads
        cutTextThreads(firstPage, lastPage);

        // remove the unwanted pages/spreads
        deleteAllPagesExceptRange(newDoc, firstPage, lastPage);

        if (settings.keepPageNumbering)
            newDoc.pages[0].appliedSection.properties = {
                continueNumbering: false,
                pageNumberStart: firstPageOffset + 1,
            };

        // populate the file name using the template
        var fileName = settings.fileNameTemplate
            .replace(/^\/?/, "/") // add leading slash
            .replace(/\{docName\}/g, docName)
            .replace(/\{startPageName\}/g, firstPage.name) /* || newDoc.pages[start].pages[0].name)*/
            .replace(/\{endPageName\}/g, lastPage.name); /* || newDoc.pages[end].pages[-1].name);*/

        // the file to save
        f = new File(filePath + fileName);

        if (!f.parent.exists)
            f.parent.create();

        newDoc.save(f);
        newDoc.close(SaveOptions.NO);

    }

    //cleanup
    doc.documentPreferences.allowPageShuffle = allowPageShuffle;
    doc.documentPreferences.preserveLayoutWhenShuffling = preserveLayoutWhenShuffling;

    if (settings.revealFolder)
        // reveal output folder
        f.parent.execute();

};
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Divide Document');

/**
 * Cuts text frame threading.
 *   - provide `startPage` to cut text threads before.
 *   - provide `endPage` to cut text threads after.
 * @author m1b
 * @version 2025-04-01
 * @param {Page} [startPage] - the start page where the threads should be cut.
 * @param {Page} [endPage] - the end page where the threads should be cut.
 */
function cutTextThreads(startPage, endPage) {

    if (startPage) {

        // cut text before this page
        for (var i = 0; i < startPage.textFrames.length; i++)
            cutTextFrame(startPage.textFrames[i], false);

    }

    if (endPage) {

        // cut text after this page
        for (var i = 0; i < endPage.textFrames.length; i++)
            if (endPage.textFrames[i].nextTextFrame)
                cutTextFrame(endPage.textFrames[i].nextTextFrame, false);

    }

};

/**
  * Cuts the text frame's story thread, so that
  * the story starts in this text frame.
  * @author m1b
  * @version 2025-04-01
  * @param {TextFrame} tf - the textframe to cut.
  * @param {Boolean} [cutThreadOnSamePage] - whether to cut the thread even if the previous text frame is on the same page (default: false).
  */
function cutTextFrame(tf, cutThreadOnSamePage) {

    var previous = tf.previousTextFrame;

    if (!previous)
        return;

    if (
        !cutThreadOnSamePage
        && previous.parentPage === tf.parentPage
    )
        return;

    var story = tf.parentStory,
        index = tf.characters[0].index;

    tf.previousTextFrame = null;
    story.characters.itemByRange(index, story.length - 1).move(LocationOptions.AT_BEGINNING, tf);

};

/**
 * Deletes all pages of `doc` before `firstPage` and after `lastPage`.
 * @author m1b
 * @version 2026-01-17
 * @param {Document} doc - an Indesign document.
 * @param {Page} firstPage - the first page to keep.
 * @param {Page} lastPage - the last page to keep.
 */
function deleteAllPagesExceptRange(doc, firstPage, lastPage) {

    var min = doc.pages.firstItem().documentOffset;
    var max = doc.pages.lastItem().documentOffset;
    var start = firstPage.documentOffset;
    var end = lastPage.documentOffset;

    var before = start > min ? [min, start - 1] : null;
    var after = end < max ? [end + 1, max] : null;

    if (after)
        // remove pages after
        doc.pages.itemByRange(after[0], after[1]).remove();

    if (before)
        // remove pages before
        doc.pages.itemByRange(before[0], before[1]).remove();

};

/**
 * Shows UI for Divide Document Script
 * @author m1b
 * @version 2026-01-17
 * @param {Object} settings - the settings to adjust via UI.
 * @returns {1|2} - result code
 */
function ui(settings) {

    loadSettings();

    // make the dialog
    var w = new Window("dialog", 'Divide Document', undefined);

    var group0 = w.add("group {orientation:'row', alignment:['fill','top'], alignChildren:['left','top'], margins:[10,10,10,10] }");
    var group1 = group0.add("group {orientation:'column', alignment:['fill','top'], alignChildren:['left','top'], margins:[10,10,10,10] }");
    var group2 = group0.add("group {orientation:'column', alignment:['fill','top'], alignChildren:['left','top'], margins:[10,10,10,10] }");
    var group3 = group0.add("group {orientation:'column', alignment:['fill','top'], alignChildren:['left','top'], margins:[10,10,10,10] }");

    var label = group1.add('statictext {preferredSize: [80,-1], text:"Divide By:", justify: "left"}');
    var radios = group1.add("group {orientation:'column', alignChildren:['left','top'] }");
    var pagesRadio = radios.add("radiobutton { text: 'Pages', value: true }");
    var spreadsRadio = radios.add("radiobutton { text: 'Spreads', value: false }");

    var label1 = group2.add('statictext {preferredSize: [160,-1], text:"[Pages/Spreads] Per Division:", justify: "left"}');
    var group2a = group2.add("group {orientation:'row' }");
    var divisionCountField = group2a.add('edittext {preferredSize: [60,-1], text:""}');
    var multiplierLabel = group2a.add('statictext {preferredSize: [100,20], text:"", justify: "left"}');
    var keepPageNumberingCheckBox = group2.add("CheckBox { text: 'Keep Numbering', alignment:['left','top'] }");

    var label2 = group3.add('statictext {text:"Filename Template:", justify: "left"}');
    var fileNameTemplateField = group3.add('edittext {preferredSize: [450,-1], text:""}');
    var infoText = group3.add('statictext {text:"", preferredSize:[350,-1], alignment:["fill","bottom"], justify: "left"}');

    var bottom = w.add("group {orientation:'row', alignment:['fill','bottom'], alignChildren:['fill','bottom'] }");
    var infoGroup = bottom.add("group {orientation:'row', alignment:['fill','bottom'], margins:[0,0,0,6] }");
    var buttons = bottom.add("Group { orientation: 'row', alignment: ['right','bottom'], margins:[0,0,0,6] }");
    var hints = infoGroup.add('statictext {text:"", justify: "left"}');
    var cancelButton = buttons.add("Button { text:'Cancel', properties:{name:'cancel'} }");
    var divideButton = buttons.add("Button { text:'Divide', enabled: true, properties:{name: 'ok'} }");

    // populate UI
    keepPageNumberingCheckBox.value = (true === settings.keepPageNumbering);
    divisionCountField.text = settings.divisionCount;
    fileNameTemplateField.text = settings.fileNameTemplate;
    hints.text = 'Placeholders: {docName} {startPageName} {endPageName} {part}';
    updateUI();

    // event handlers
    spreadsRadio.onClick = updateUI
    pagesRadio.onClick = updateUI;
    divisionCountField.onChanging = updateUI;
    fileNameTemplateField.onChanging = updateUI;
    divideButton.onClick = done;

    w.center();
    return w.show();

    /** @returns {Number} - the number of divisions */
    function getPartCount() {

        var divideBy = pagesRadio.value ? "pages" : "spreads";
        var partCount = Math.ceil(settings.doc[divideBy].length / Number(divisionCountField.text));

        if (
            !divisionCountField.text
            || isNaN(partCount)
        )
            return;

        return partCount;

    };

    function updateUI() {

        var partCount = getPartCount();

        divideButton.enabled = undefined != partCount;

        if (!partCount)
            return;

        label1.text = (pagesRadio.value ? 'Pages' : 'Spreads') + ' Per Division:';
        multiplierLabel.text = '\u00D7 ' + partCount + ' documents';

        // show mock filename
        infoText.text = 'Example: ' + fileNameTemplateField.text
            .replace(/^\/?/, "/") // add leading slash
            .replace(/\{docName\}/g, settings.doc.name.replace(/\.[^\.]+$/, ''))
            .replace(/\{startPageName\}/g, 1)
            .replace(/\{endPageName\}/g, divisionCountField.text)
            .replace(/\{part\}/g, partCount);

    };

    function done() {

        // update settings
        settings.divideBy = pagesRadio.value ? "pages" : "spreads";
        settings.divisionCount = Number(divisionCountField.text);
        settings.fileNameTemplate = fileNameTemplateField.text;
        settings.keepPageNumbering = keepPageNumberingCheckBox.value;

        saveSettings();

        // close window with success code
        w.close(1);
    };

    function loadSettings() {

        // load last-used settings from document
        settings.divideBy = settings.doc.extractLabel('divideBy') || settings.divideBy;
        settings.divisionCount = Number(settings.doc.extractLabel('divisionCount') || settings.divisionCount);
        settings.fileNameTemplate = settings.doc.extractLabel('fileNameTemplate') || settings.fileNameTemplate;

        var keepPageNumbering = settings.doc.extractLabel('keepPageNumbering');
        settings.keepPageNumbering = '' == keepPageNumbering
            ? settings.keepPageNumbering
            : 'FALSE' === keepPageNumbering;

    };

    function saveSettings() {

        // store last-used languages in document
        settings.doc.insertLabel('divideBy', settings.divideBy);
        settings.doc.insertLabel('divisionCount', String(settings.divisionCount));
        settings.doc.insertLabel('fileNameTemplate', settings.fileNameTemplate);
        settings.doc.insertLabel('keepPageNumbering', (settings.keepPageNumbering ? "TRUE" : "FALSE"));

    };

};