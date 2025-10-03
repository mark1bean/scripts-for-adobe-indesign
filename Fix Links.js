/**
 * @file Fix Links.js
 * @author m1b
 * @version 2025-10-03
 * ----------------------------------------------------------------
 * Use this script when moving Indesign document(s) between
 * environments to re-link to the same files but in the
 * new location, for example moving
 *
 * from user "bob" at
 *     /Users/bob/Dropbox/
 *
 * to user "alice" at
 *    /Users/alice/Library/CloudStorage/Dropbox.
 * ----------------------------------------------------------------
 * Instructions:
 *   1. Run script
 *   2. Choose a link from the Broken Links list.
 *   3. Click "Find" button and manually relink to that correct file.
 *   4. Script will use the new link to try to fix broken links.
 *   5. Repeat step 2 and 3 until Broken Links list is empty.
 *   6. Click "Fix [N] Links" button to perform the changes to the open document(s).
 *
 * If desired, save the setting with the document, so
 * you won't have to establish the change every time.
 *
 * TIP: changes that are saved in the document work backwards, too, so if
 * Alice sends the indesign file back, when Bob uses the script it will be
 * ready to fix all the links with just a press of the "Fix [N] Links" button.
 *
 * Click Reset button to start again.
 */

function main() {

    var settings = {
        brokenLinks: [],
        processAllOpenDocuments: true,
        saveInDoc: true,
        showResults: true,
        swaps: [],
    };

    if (0 === app.documents.length)
        return alert('Please open a document and try again.');

    var result = ui(settings);

    if (
        2 === result
        || settings.swaps.length < 2
    )
        // user cancelled
        return;

    if (undefined != settings.cleanup) {
        settings.cleanup();
        delete settings.cleanup;
    }

    // perform the relinking
    var counter = 0;

    for (var i = 0, f; i < settings.brokenLinks.length; i++) {

        f = File(settings.brokenLinks[i].newPath);

        if (f.exists) {
            settings.brokenLinks[i].link.relink(f);
            counter++;
        }

    }

    if (settings.showResults)
        alert('Fixed ' + counter + ' links.');

};
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Swap Links");

/**
 * UI for Fix Links script.
 * @author m1b
 * @version 2025-10-03
 * @param {Object} settings - the script settings.
 * @returns {1|2} - ScriptUI result code (1 = good, 2 = user cancelled).
 */
function ui(settings) {

    const PATH_LIST_WIDTH = 1000;
    const PATH_LIST_HEIGHT = 800;
    const LIST_ROW_HEIGHT = 26;

    var key = 'FixLinks';
    var already = {};

    if (settings.saveInDoc)
        loadSwaps();

    var activeBrokenLink;
    var fixCount = 0;
    var brokenLinks = getBrokenLinks(settings.processAllOpenDocuments, true);

    for (var i = 0; i < settings.swaps.length; i++)
        fixLinks(settings.swaps[i]);

    var w = new Window("dialog", 'Fix Links'),

        relinkPanel = w.add('panel {margins: [20,20,20,20], text:"Relink this broken link:",alignment:["fill","fill"], }'),
        relinkGroup = relinkPanel.add('group { orientation:"row", alignment:["fill","fill"], alignChildren:["fill","top"] }'),
        currentPathText = relinkGroup.add('statictext { preferredSize: [-1,35], text:"", justify: "left", alignment:["fill","top"], properties: {multiline: true, scrolling: true } }'),
        relinkButton = relinkGroup.add('button { text:"Find", alignment:["right","top"]}'),

        brokenLinksListGroup = w.add('group { orientation:"row", alignChildren:["center","top"] }'),
        brokenLinksList = brokenLinksListGroup.add("ListBox { alignment:['fill','fill'], preferredSize:[-1,-1], properties:{ multiselect:false, showHeaders:true, numberOfColumns:1, columnTitles:['Broken Links'], columnWidths:[-1] } }"),

        bottomUI = w.add("group {orientation:'row', alignment:['fill','top'], margins: [0,20,0,0] }"),
        saveInDocCheckBox = bottomUI.add("CheckBox { text: 'Save settings in Document' }"),
        resetButton = bottomUI.add('button { text:"Reset" }'),
        buttons = bottomUI.add("group {orientation:'row', alignment:['right','top'], alignChildren:'right' }"),
        cancelButton = buttons.add('button { text:"Cancel", properties: { name: "cancel" } }'),
        fixLinksButton = buttons.add('button { text:"", preferredSize:[130,-1] }');

    brokenLinksList.preferredSize = [PATH_LIST_WIDTH, Math.min(PATH_LIST_HEIGHT, brokenLinks.length * LIST_ROW_HEIGHT)];
    saveInDocCheckBox.value = true === settings.saveInDoc;

    /** Removes all swaps and currently fixed paths. */
    resetButton.onClick = function () {

        settings.swaps = [];
        already = {};

        // clear fixed paths
        for (var i = 0; i < brokenLinks.length; i++)
            brokenLinks[i].newPath = undefined;

        buildBrokenLinksList();
        updateUI();

    };

    /**
     * Asks for a new path to relink the active broken link,
     * and, using this, tries to fix other broken links.
     */
    relinkButton.onClick = function () {

        // ask user for a new path
        var testPath = decodeURI((File.openDialog() || 0).relativeURI);

        // derive a find/change pair from the new and old paths
        var swap = getFindChangePair(currentPathText.text, testPath, true);

        // now try to fix the other broken links
        var counter = fixLinks(swap);

        if (0 === counter)
            return;

        if (!already[swap.findWhat + swap.changeTo]) {
            settings.swaps.push(swap);
            already[swap.findWhat + swap.changeTo] = true;
        }

        buildBrokenLinksList();
        updateUI();
    };

    brokenLinksList.onChange = function () {
        if (this.selection) {
            activeBrokenLink = brokenLinks[this.selection.linkIndex];
            currentPathText.text = activeBrokenLink.link.filePath;
        }
        updateUI();
    };

    buildBrokenLinksList();
    updateUI();

    /** Updates `settings` object and closes window. */
    fixLinksButton.onClick = function () {

        // update settings
        settings.saveInDoc = true == saveInDocCheckBox.value;
        settings.brokenLinks = brokenLinks;

        if (settings.saveInDoc) {
            // cleanup function runs after modal is closed
            settings.cleanup = function () {
                // save settings in each document
                var docs = settings.processAllOpenDocuments ? getOpenVisibleDocuments() : [app.activeDocument];
                for (var i = 0; i < docs.length; i++)
                    saveDocumentLabel(docs[i], key, ['swaps'], settings);
            };
        }

        w.close(1);
    };

    w.center();
    return w.show();

    /** Updates UI elements */
    function updateUI() {
        resetButton.enabled = 0 !== settings.swaps.length;
        fixLinksButton.enabled = 0 !== fixCount;
        relinkPanel.enabled = null !== brokenLinksList.selection;
    };

    /** Build or rebuild the broken links list items. */
    function buildBrokenLinksList() {

        fixCount = 0;

        brokenLinksList.removeAll();

        for (var i = 0, item; i < brokenLinks.length; i++) {
            if (brokenLinks[i].newPath) {
                fixCount++;
                continue;
            }
            item = brokenLinksList.add('item', brokenLinks[i].link.filePath);
            item.linkIndex = i;
        }

        fixLinksButton.text = 'Fix ' + fixCount + ' Links';

    };

    /**
     * Load swaps into settings from the document(s).
     */
    function loadSwaps() {

        var docs = settings.processAllOpenDocuments ? getOpenVisibleDocuments() : [app.activeDocument];

        // avoid loading same swap twice
        for (var i = 0; i < settings.swaps.length; i++)
            already[settings.swaps[i].findWhat + settings.swaps[i].changeTo] = true;

        // look in each document
        for (var i = 0, loaded; i < docs.length; i++) {

            loaded = {};
            loadDocumentLabel(docs[i], key, loaded);

            if (
                !loaded.swaps
                || 0 === loaded.swaps.length
            )
                continue;

            // add new loaded swaps
            for (var j = 0; j < loaded.swaps.length; j++) {
                var swap = loaded.swaps[j];
                if (!already[swap.findWhat + swap.changeTo]) {
                    settings.swaps.push(swap);
                    already[swap.findWhat + swap.changeTo] = true;
                }
            }

        }

    };

    /**
     * Repair broken links by find/change operation on the file paths.
     * @param {Object} swap - the find/change pair {findWhat: changeTo:}.
     * @returns {Number} - the number of links fixed.
     */
    function fixLinks(swap) {

        var counter = 0;

        for (var i = 0, f1, f2; i < brokenLinks.length; i++) {

            if (brokenLinks[i].newPath)
                continue;

            f1 = File(brokenLinks[i].link.filePath.replace(swap.findWhat, swap.changeTo));
            f2 = File(brokenLinks[i].link.filePath.replace(swap.changeTo, swap.findWhat));

            if (f1.exists) {
                brokenLinks[i].newPath = String(f1);
                counter++;
            }
            // try the other way!
            else if (f2.exists) {
                brokenLinks[i].newPath = String(f2);
                counter++;
            }

        }

        if (1 < counter)
            settings.swaps.push(swap);

        return counter;

    };

};

/**
 * Returns array of broken (missing) link objects.
 * @author m1b
 * @version 2025-10-03
 * @param {Boolean} fromAllDocuments - whether to collect broken links from all open documents (default: active document only).
 * @param {Boolean} sort - whether to sort the links by file path (default: false).
 * @returns {Array<Object>} - { link: Link, newPath: undefined }
 */
function getBrokenLinks(fromAllDocuments, sort) {

    var brokenLinks = [];
    var docs = fromAllDocuments ? getOpenVisibleDocuments() : [app.activeDocument];
    var allLinks = [];

    // collect links from every document
    for (var i = 0; i < docs.length; i++)
        allLinks = allLinks.concat(docs[i].links.everyItem().getElements());

    // collect the broken links
    for (var i = 0; i < allLinks.length; i++)
        if (LinkStatus.LINK_MISSING === allLinks[i].status)
            brokenLinks.push({ link: allLinks[i], newPath: '' });

    if (sort) {
        brokenLinks.sort(function (a, b) {
            if (a.link.filePath.toLowerCase() < b.link.filePath.toLowerCase())
                return -1;
            else if (a.link.filePath.toLowerCase() > b.link.filePath.toLowerCase())
                return 1;
            else
                return 0;
        });
    }

    return brokenLinks;

};

/**
 * Returns all open, visible documents.
 * @author m1b
 * @version 2025-03-29
 * @returns {Array<Document>}
 */
function getOpenVisibleDocuments() {

    var docs = app.documents.everyItem().getElements(),
        visibleDocs = [];

    for (var i = 0; i < docs.length; i++)
        if (docs[i].visible)
            visibleDocs.push(docs[i]);

    return visibleDocs;

};

/**
 * Returns {findWhat: changeTo:} suitable for changing `str1` into `str2`.
 * (In worse case findWhat == str1 and changeTo == str2.)
 * @author m1b
 * @version 2025-10-03
 * @param {String} str1 - the 'starting' string.
 * @param {String} str2 - the 'destination' string.
 * @param {Bool} [asPrefixes] - whether the returned strings always match the start of the strings.
 * @returns {Object} - {findWhat: changeTo: }
 */
function getFindChangePair(str1, str2, asPrefixes) {

    var len1 = str1.length,
        len2 = str2.length,
        i = 0;

    // find the first difference
    while (
        i < len1
        && i < len2
        && str1.charAt(i) === str2.charAt(i)
    ) {
        i++;
    }

    // find the part of str1 to be replaced
    var j = len1,
        k = len2;

    while (
        j > i
        && k > i
        && str1.charAt(j - 1) === str2.charAt(k - 1)
    ) {
        j--;
        k--;
    }

    if (asPrefixes)
        i = 0;

    return {
        findWhat: str1.substring(i, j),
        changeTo: str2.substring(i, k),
    };

};

/**
 * Stores keys in document label `key`.
 * in triplets of 'key|type|value'.
 *
 * Valid Types are 'Number', 'Boolean', and 'String'.
 * (Experimental type 'Specifier' is untested)
 *
 * Retrieve using `loadDocumentLabel` function.
 *
 * @author m1b
 * @version 2024-07-11
 * @param {Document} doc - an Indesign Document.
 * @param {String} key - the label key.
 * @param {Array<String>} keys - the keys to `obj`.
 * @param {Object} obj - the object of values.
 */
function saveDocumentLabel(doc, key, keys, obj) {

    var payload = [],
        delimiter = '|';

    for (var i = 0; i < keys.length; i++) {

        var k = keys[i],
            v = obj[keys[i]];

        if (undefined == v) {
            payload.push([k, 'undefined', undefined].join(delimiter));
        }

        else if ('Array' === v.constructor.name) {
            payload.push([k, 'Array', v.toSource()].join(delimiter));
        }

        else if ('Number' === v.constructor.name) {
            payload.push([k, 'Number', v].join(delimiter));
        }

        else if ('Boolean' === v.constructor.name) {
            payload.push([k, 'Boolean', v].join(delimiter));
        }

        else if (
            v.hasOwnProperty('isValid')
            && v.isValid
        ) {
            payload.push([k, 'Specifier', v.toSpecifier()].join(delimiter));
        }

        else {
            payload.push([k, 'String', String(v)].join(delimiter));
        }

    }

    if (payload.length)
        doc.insertLabel(key, payload.join(delimiter));

};


/**
 * Parses values stored in document label `key`
 * via the `saveDocumentLabel` function.
 *
 * Expects values to be triplets of 'key|type|value'.
 * Valid Types are 'Number', 'Boolean', and 'String'.
 *
 * Note that `loadDocumentLabel` mutates `obj` so
 * `obj` must be supplied, usually containing default
 * values will may be overwritten if found in document.
 *
 * @author m1b
 * @version 2024-07-11
 * @param {Document} doc - an Indesign Document.
 * @param {String} key - the key, usually the script id.
 * @param {Object} obj - an object containing default values.
 */
function loadDocumentLabel(doc, key, obj) {

    if (!obj)
        throw Error('Utility.loadDocumentLabel: bad `obj` supplied.');

    var values = doc.extractLabel(key).split('|');

    while (values.length > 2) {

        // expects key/type/value triplets
        var k = values.shift(),
            t = values.shift(),
            v = values.shift(),
            success = false;

        if ('Array' === t) {
            v = eval(v);
            success = 'Array' === v.constructor.name;
        }

        else if ('Number' === t) {
            v = Number(v)
            success = !isNaN(v)
        }

        else if ('Boolean' === t) {
            v = Boolean(v);
            success = true;
        }

        else if ('undefined' === t) {
            v = undefined;
            success = true;
        }

        else if ('Specifier' === t) {
            v = resolve(v);
            success = v && v.isValid;
        }

        else if ('String' === t) {
            success = true;
        }

        if (success)
            obj[k] = v;

    }

};