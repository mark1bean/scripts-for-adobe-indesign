/**
 * @file Numbered Markers.js
 *
 * How to set up document:
 *   1. Create an object style for your numbered markers, eg. a circle with the number in it.
 *   2. Create a marker object with the style applied from step 1.
 *   3. Create a paragraph style for your numbered listing paragraphs. This must use Indesign's
 *      bullets and numbering system.
 *   4. Type your numbered list and apply the style from step 3.
 *   5. Run this script, choosing styles from steps 1 and 2 and also make sure that "Create missing markers" is ON.
 *
 * The script will create all the markers you need. Each marker is linked to the numbered list because its
 * "script label" (Windows > Utilities > Script Label) matches the contents of the numbered list text.
 *
 * @author m1b
 * @version v0.1 (2025-05-14)
 * @discussion https://community.adobe.com/t5/indesign-discussions/any-suggestions-of-redesigning-this-map/m-p/15315340
 */
function main() {

    var settings = {
        markerObjectStyleName: 'Place Marker',
        numberedParagraphStyleName: 'Place Listing',
        createMissingMarkers: true,
        removeOrphanMarkers: true,
        showUI: true,
    };

    if (0 === app.documents.length)
        return alert('Please open a document and try again.');

    var doc = settings.doc = app.activeDocument;

    if (settings.showUI) {

        var result = ui(settings);

        if (2 === result)
            // user cancelled UI
            return;

    }

    var markerObjectStyle = doc.objectStyles.itemByName(settings.markerObjectStyleName);
    if (!markerObjectStyle.isValid)
        return alert('There is no object style "' + settings.markerObjectStyleName + '"');

    var numberedParagraphStyle = doc.paragraphStyles.itemByName(settings.numberedParagraphStyleName);
    if (!numberedParagraphStyle.isValid)
        return alert('There is no paragraph style "' + settings.numberedParagraphStyleName + '"');

    // these are the texts which have numbering applied, eg. a list of map locations, numbered 1, 2, 3 ...
    var numberedTexts = findTextsWithStyleAndNumbering(doc, numberedParagraphStyle);

    // these are the marker objects, eg. circular text frame with number contents
    var markers = findObjectsWithStyle(doc, markerObjectStyle, onlyThingsWithContents);

    var messages = [],
        createdMarkerNumbers = [],
        templateMarker = markers[0],
        counter = 0;

    for (var i = 0, marker, index; i < numberedTexts.length; i++) {

        // look for the marker with label matching the numbered text
        index = getIndex(markers, 'label', getContents(numberedTexts[i]));
        marker = markers[index];

        if (marker) {
            // marker is valid, so we'll clear it from the list, so it doesn't get removed later
            markers[index] = undefined;
        }

        else if (settings.createMissingMarkers) {

            if (
                !templateMarker
                || !templateMarker.hasOwnProperty('contents')
            )
                continue;

            // create the missing marker, and store the number for showing results
            createdMarkerNumbers.push(createMarker(templateMarker, numberedTexts[i]).contents);

            continue;

        }

        else {
            // didn't find this one
            messages.push('Did not find place marker for place listing ' + numberedTexts[i].paragraphs[0].numberingResultNumber + '.');
            continue;
        }

        // update contents
        setMarkerItemContents(marker, numberedTexts[i]);
        marker.recompose();
        counter++;

    }

    if (settings.removeOrphanMarkers) {

        var removalCount = 0

        // remove any markers that weren't matched
        for (var i = markers.length - 1; i >= 0; i--)
            if (markers[i]) {
                markers[i].remove()
                removalCount++;
            }

        if (removalCount > 0)
            messages.push('Removed ' + removalCount + ' orphan marker' + (removalCount > 1 ? 's' : '') + '.');

    }

    if (createdMarkerNumbers.length > 0)
        messages.push('Created markers: ' + createdMarkerNumbers.join(', ') + '.');


    alert('Updated ' + counter + ' place markers.' + (0 === messages.length ? '' : '\n' + messages.join('\n')));

};
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Numbered Markers');


/**
 * Find page items in the given object style.
 * @author m1b
 * @version 2025-05-12
 * @param {Document} doc - an Indesign Document.
 * @param {ObjectStyle} objectStyle - the style to search for.
 * @param {Function} filter - a filter function.
 * @returns {Array<PageItem>}
 */
function findObjectsWithStyle(doc, objectStyle, filter) {

    if ('ObjectStyle' !== objectStyle.constructor.name)
        throw new Error('findObjectsWithStyle: bad `objectStyle` supplied.');

    app.findObjectPreferences = null;
    app.findObjectPreferences.appliedObjectStyles = objectStyle;

    var found = doc.findObject(),
        items = [];

    for (var i = 0; i < found.length; i++)
        if (!filter || filter(found[i]))
            items.push(found[i]);

    return items;

};

function onlyThingsWithContents(item) {
    return item.hasOwnProperty('contents');
}

/**
 * Find texts - not paragraphs, due to ommitting carriage returns - in the given paragraph style.
 * @author m1b
 * @version 2025-05-12
 * @param {Document} doc - an Indesign Document.
 * @param {ParagraphStyle} paragraphStyle - the style to search for.
 * @returns {Array<Text>}
 */
function findTextsWithStyleAndNumbering(doc, paragraphStyle) {

    if ('ParagraphStyle' !== paragraphStyle.constructor.name)
        return;

    if (ListType.NUMBERED_LIST !== paragraphStyle.bulletsAndNumberingListType)
        return [];

    app.findGrepPreferences = null;
    app.findGrepPreferences.findWhat = '^[^\\r]+$';
    app.findGrepPreferences.appliedParagraphStyle = paragraphStyle;

    return doc.findGrep();

};

/**
 * Returns a thing with matching property.
 * If `key` is undefined, evaluate the object itself.
 * @author m1b
 * @version 2024-04-21
 * @param {Array|Collection} things - the things to look through.
 * @param {String} [key] - the property name (default: undefined).
 * @param {*} value - the value to match.
 * @returns {*?} - the thing, if found.
 */
function getThing(things, key, value) {

    for (var i = 0; i < things.length; i++)
        if ((undefined == key ? things[i] : things[i][key]) == value)
            return things[i];

};

/**
 * Returns the index of the thing with matching property.
 * If `key` is undefined, evaluate the object itself.
 * @author m1b
 * @version 2025-05-14
 * @param {Array|Collection} things - the things to look through.
 * @param {String} [key] - the property name (default: undefined).
 * @param {*} value - the value to match.
 * @returns {Number} - the thing's index, if found, or -1.
 */
function getIndex(things, key, value) {

    for (var i = 0; i < things.length; i++) {

        if (undefined == things[i])
            continue;

        if ((undefined == key ? things[i] : things[i][key]) == value)
            return i;

    }

    return -1;

};

/**
 * Returns an array of values, one for each thing,
 * where the value is access using `keyPath`,
 * for example:
 *
 *     var items = app.activeDocument.symbolItems;
 *     var symbolNames = getProperties(items, 'symbol/name');
 *
 * will collect every symbolItem's symbol's name.
 *
 * Note: will return an array of things.length,
 * even if some elements are undefined.
 *
 * @author m1b
 * @version 2024-08-29
 * @param {Array<*>|Collection} things - the things to access.
 * @param {String} keyPath - the key(s) separated by /.
 * @returns {Array<*>}
 */
function getProperties(things, keyPath) {

    var found = [],
        keys = keyPath.split('/');

    for (var i = 0; i < things.length; i++) {

        var value = things[i];

        // get the deepest value we can
        for (var j = 0; j < keys.length; j++) {

            if (
                null != value
                && value.hasOwnProperty(keys[j])
            ) {
                value = value[keys[j]];
            }

            else {
                value = undefined;
                break;
            }
        }

        found[i] = value;

    }

    return found;

};

/**
 * Returns index of `obj` within `array`,
 * or undefined if `obj` not found.
 * @param {Array<*>} array - the array to search in.
 * @param {*} obj - the object to look for.
 * @returns {Number?}
 */
function indexOfArray(array, obj) {

    for (var i = 0; i < array.length; i++)
        if (array[i] == obj)
            return i;

    return -1;

};

/**
 * Set a text frame's contents to the numbering result text of `text`,
 * minus (hopefully!) any extraneous punctuation or decorations
 * @author m1b
 * @version 2025-05-13
 * @param {TextFrame} textFrame - the target text frame.
 * @param {Text} text - the source text.
 * @returns {TextFrame}
 */
function setMarkerItemContents(textFrame, text) {

    textFrame.contents = text.paragraphs[0].bulletsAndNumberingResultText
        .replace(/(^[^A-Za-z0-9.-]*|[^A-Za-z0-9]*$)/g, '');
    // .replace(/[\s\.\u2013\u2014:-]*$/, '');

};

/**
 * Returns the contents of `text`, with
 * leading/trailing space removed.
 * @author m1b
 * @version 2025-05-13
 * @param {Text} text - the text to return.
 * @returns {Boolean} - whether the marker item matches the text.
 */
function getContents(text) {

    return text.contents.replace(/^\s*|\s*$/g, '');

};

/**
 * Creates the ordinal marker items, by duplicate a given
 * template item and an array of numbered texts.
 * Note: the template item will become the first marker item.
 * @author m1b
 * @version 2025-05-13
 * @param {TextFrame} templateMarker - the marker item upon which the other will be based.
 * @param {Array<Text>} numberedTexts - the numbered texts which to link to the markers.
 * @returns {Array<TextFrame>} - the marker items.
 */
function createMarkers(templateMarker, numberedTexts) {

    var markers = [];

    // create all the other markers
    for (var i = 0; i < numberedTexts.length; i++)
        markers.push(createMarker(templateMarker, numberedTexts[i]));

    return markers;

};

/**
 * Creates a new marker, based on `templateMarker`
 * but with the numbering of `numberedText`.
 * @author m1b
 * @version 2025-05-14
 * @param {TextFrame} templateMarker
 * @param {Text} numberedText
 * @returns {TextFrame} - the new marker.
 */
function createMarker(templateMarker, numberedText) {

    if (!templateMarker.hasOwnProperty('contents'))
        return;

    var newMarker = templateMarker.duplicate();
    setMarkerItemContents(newMarker, numberedText);
    newMarker.label = getContents(numberedText);
    newMarker.move([
        numberedText.insertionPoints[0].horizontalOffset - 5 - (templateMarker.geometricBounds[3] - templateMarker.geometricBounds[1]) * (numberedText.paragraphs[0].numberingResultNumber % 2 + 2),
        numberedText.baseline - numberedText.ascent
    ]);

    return newMarker;

};

/**
 * Shows UI for Update Linked Numbered List
 * @param {Object} settings - the settings to adjust via UI.
 * @returns {1|2} - result code
 */
function ui(settings) {

    var ExitCode = {
        PERFORM_ACTION: 1,
        CANCEL: 2,
    };

    var numberedTexts = [],
        markers = [],
        paragraphStyles = [],
        objectStyles = [];

    for (var i = 0; i < settings.doc.allParagraphStyles.length; i++)
        if (settings.doc.allParagraphStyles[i].name[0] !== '[')
            paragraphStyles.push(settings.doc.allParagraphStyles[i]);

    for (var i = 0; i < settings.doc.allObjectStyles.length; i++)
        if (settings.doc.allObjectStyles[i].name[0] !== '[')
            objectStyles.push(settings.doc.allObjectStyles[i]);

    loadSettingsFromDocument();

    var w = new Window("dialog", 'Numbered Markers', undefined),

        group = w.add('group {orientation:"row", alignment:["left","top"], alignChildren: ["left","top"], margins:[10,10,10,0], preferredSize: [120,-1] }'),

        paraGroup = group.add('group {orientation:"column", alignment:["fill","top"], margins:[0,0,20,0], preferredSize:[120,-1] }'),
        paraLabel = paraGroup.add("statictext { text:'Paragraph Style of Numbered List:', alignment:['left','center'], justify:'left' }"),
        paraFieldGroup = paraGroup.add('group {orientation:"row", alignment:["fill","top"], margins:[0,0,0,0], preferredSize:[120,-1] }'),
        paragraphStyleMenu = paraFieldGroup.add('dropDownList { title:"", preferredSize:[210,-1] }'),
        paraStatus = paraGroup.add("statictext { text:'', alignment:['fill','center'] }"),

        objectGroup = group.add('group {orientation:"column", alignment:["fill","top"], margins:[0,0,0,0], preferredSize:[120,-1] }'),
        objectLabel = objectGroup.add("statictext { text:'Object Style of Markers:', alignment:['left','center'], justify:'left' }"),
        objectFieldGroup = objectGroup.add('group {orientation:"row", alignment:["fill","fill"], margins:[0,0,0,0], preferredSize:[120,-1] }'),
        objectStyleMenu = objectFieldGroup.add('dropDownList { title:"", preferredSize:[210,-1] }'),
        objectStatusGroup = objectGroup.add('group {orientation:"column", alignment:["fill","fill"], margins:[0,0,0,0], preferredSize:[120,-1] }'),
        objectStatus = objectStatusGroup.add("statictext { text:'', alignment:['fill','center'] }"),
        objectCheckboxes = objectStatusGroup.add('group {orientation:"column", alignment:["fill","fill"], margins:[0,0,0,0], preferredSize:[120,-1] }'),

        bottom = w.add('group {orientation:"row", alignment:["fill","fill"], margins:[0,0,0,0] }'),

        extras = bottom.add('group {orientation:"column", alignment:["fill","fill"], margins:[10,10,10,10] }'),
        removeOrphanCheckBox = extras.add("Checkbox { text:'Remove orphan markers', alignment:['left','center'], margins:[0,0,0,0], value:false }"),
        createMissingMarkersCheckBox = extras.add("Checkbox { text:'Create Missing Markers', alignment:['left','center'], margins:[0,0,0,0], value:false }"),
        // createMarkersButton = extras.add("Button { text:'Create Markers', properties:{ enabled:false } }"),

        buttons = bottom.add("Group { orientation:'row', alignment:['right','bottom'], margins:[0,0,0,6] }"),
        cancelButton = buttons.add("Button { text:'Cancel', properties:{name:'cancel'} }"),
        performActionButton = buttons.add("Button { text:'Number', enabled:false, enabled:true, name:'ok' }");

    // populate the menu
    populateMenu(paragraphStyleMenu, paragraphStyles);
    populateMenu(objectStyleMenu, objectStyles);

    /** handle changes to the paragraph style menu */
    paragraphStyleMenu.onChange = function () {
        settings.numberedParagraphStyleName = paragraphStyleMenu.selection.text;
        numberedTexts = findTextsWithStyleAndNumbering(settings.doc, getThing(settings.doc.allParagraphStyles, 'name', settings.numberedParagraphStyleName));
        updateStatusText();
    };

    /** handle changes to the object style menu */
    objectStyleMenu.onChange = function () {
        settings.markerObjectStyleName = objectStyleMenu.selection.text;
        markers = findObjectsWithStyle(settings.doc, getThing(settings.doc.allObjectStyles, 'name', settings.markerObjectStyleName));
        updateStatusText();
    };

    // createMarkersButton.onClick = function () {
    //     settings.markers = markers;
    //     settings.numberedTexts = numberedTexts;
    //     saveSettingsInDocument();
    //     w.close(ExitCode.CREATE_MARKERS);
    // };

    performActionButton.onClick = function () {
        settings.createMissingMarkers = createMissingMarkersCheckBox.value;
        settings.removeOrphanMarkers = removeOrphanCheckBox.value;
        saveSettingsInDocument();
        w.close(ExitCode.PERFORM_ACTION);
    };

    updateUI();

    w.center();
    return w.show();

    /** populate a menu with the names of things */
    function populateMenu(menu, things) {
        for (var i = 0; i < things.length; i++)
            menu.add('item', things[i].name);
    };

    function updateUI() {
        paragraphStyleMenu.selection = Math.max(0, indexOfArray(getProperties(paragraphStyles, 'name'), settings.numberedParagraphStyleName));
        objectStyleMenu.selection = Math.max(0, indexOfArray(getProperties(objectStyles, 'name'), settings.markerObjectStyleName));
        removeOrphanCheckBox.value = settings.removeOrphanMarkers;
        createMissingMarkersCheckBox.value = settings.createMissingMarkers;
    };

    function updateStatusText() {

        paraStatus.text = numberedTexts.length + ' texts found  ';
        objectStatus.text = markers.length + ' markers found  ';
        performActionButton.enabled = numberedTexts.length && markers.length;

        // remove orphan markers when there are more markers than texts
        // removeOrphanCheckBox.visible = numberedTexts.length < markers.length;

        // remove orphan markers when there are more markers than texts
        // createMissingMarkersCheckBox.visible = numberedTexts.length > markers.length;

    };

    /** Saves user settings to the document. */
    function saveSettingsInDocument() {
        // store last-used languages in document
        settings.doc.insertLabel('numberedParagraphStyleName', settings.numberedParagraphStyleName);
        settings.doc.insertLabel('markerObjectStyleName', settings.markerObjectStyleName);
        settings.doc.insertLabel('createMissingMarkers', settings.createMissingMarkers ? 'TRUE' : 'FALSE');
        settings.doc.insertLabel('removeOrphanMarkers', settings.removeOrphanMarkers ? 'TRUE' : 'FALSE');
    };

    /** Loads user settings from the document. */
    function loadSettingsFromDocument() {
        // load last-used settings from document
        settings.numberedParagraphStyleName = settings.doc.extractLabel('numberedParagraphStyleName') || settings.numberedParagraphStyleName;
        settings.markerObjectStyleName = settings.doc.extractLabel('markerObjectStyleName') || settings.markerObjectStyleName;
        settings.createMissingMarkers = ('TRUE' === settings.doc.extractLabel('createMissingMarkers')) || settings.createMissingMarkers;
        settings.removeOrphanMarkers = ('TRUE' === settings.doc.extractLabel('removeOrphanMarkers')) || settings.removeOrphanMarkers;
    };

};