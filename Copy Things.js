/**
 * @file Copy Things.js
 *
 * Copy things from one document to other(s).
 *
 * Supports:
 *   - Character Styles
 *   - Paragraph Styles
 *   - Object Styles
 *   - Cell Styles
 *   - TableStyles
 *   - Swatchs
 *
 * @author m1b
 * @version 0.1 / 2025-05-11
 */
function main() {
    // the source document
    var doc = app.activeDocument;

    /* ----------------- *
     *  SETTINGS         *
     * ----------------- */
    var settings = {

        // user settings
        pathDelimiter: ' > ',
        activeThingConstructorName: 'Swatch',
        showDefaultThings: false,
        showResults: true,
        showUI: true,

        // internal settings
        allPaths: [],
        pathsToCopy: [],
        source: doc,
        destinations: [],
        activeThingType: undefined,
    };

    /* ----------------------------------------------------- *
     *  THE TYPES OF THINGS WE CAN COPY                      *
     * ----------------------------------------------------- *
     * These objects contain everything particular to        *
     * a particular thing type, access using the thing's     *
     * constructor's name. (See `Swatch` for special cases   *
     * with documentation.)                                  *
     * ----------------------------------------------------- */
    $.global.ThingTypes = {

        'CharacterStyle': {
            constructorName: 'CharacterStyle',
            documentProperty: 'allCharacterStyles',
            groupConstructorName: 'CharacterStyleGroup',
            labelPlural: 'Character Styles',
            parentName: 'parent',
            paths: [],
            namePlural: 'characterStyles',
            attachPayload: function (textFrame, style) { textFrame.texts[0].appliedCharacterStyle = style },
            getFromDocument: function (doc) { return doc.allCharacterStyles },
            getChildren: function (group) { return group.characterStyles },
            getSubGroups: function (group) { return group.characterStyleGroups },
        },

        'ParagraphStyle': {
            constructorName: 'ParagraphStyle',
            documentProperty: 'allParagraphStyles',
            groupConstructorName: 'ParagraphStyleGroup',
            labelPlural: 'Paragraph Styles',
            parentName: 'parent',
            paths: [],
            namePlural: 'paragraphStyles',
            attachPayload: function (textFrame, style) { textFrame.texts[0].appliedParagraphStyle = style },
            getFromDocument: function (doc) { return doc.allParagraphStyles },
            getChildren: function (group) { return group.paragraphStyles },
            getSubGroups: function (group) { return group.paragraphStyleGroups },
        },

        'ObjectStyle': {
            constructorName: 'ObjectStyle',
            documentProperty: 'allObjectStyles',
            groupConstructorName: 'ObjectStyleGroup',
            labelPlural: 'Object Styles',
            parentName: 'parent',
            paths: [],
            namePlural: 'objectStyles',
            getChildren: function (group) { return group.objectStyles },
            getSubGroups: function (group) { return group.objectStyleGroups },
            getFromDocument: function (doc) { return doc.allObjectStyles },
            attachPayload: function (pageItem, style) { pageItem.appliedObjectStyle = style },
        },

        'CellStyle': {
            constructorName: 'CellStyle',
            documentProperty: 'allCellStyles',
            groupConstructorName: 'CellStyleGroup',
            labelPlural: 'Cell Styles',
            parentName: 'parent',
            paths: [],
            namePlural: 'cellStyles',
            attachPayload: function (pageItem, style) {
                var table = pageItem.texts[0].tables.add();
                table.cells[0].appliedCellStyle = style;
            },
            getFromDocument: function (doc) { return doc.allCellStyles },
            getChildren: function (group) { return group.cellStyles },
            getSubGroups: function (group) { return group.cellStyleGroups },
        },

        'TableStyle': {
            constructorName: 'TableStyle',
            documentProperty: 'allTableStyles',
            groupConstructorName: 'TableStyleGroup',
            labelPlural: 'Table Styles',
            parentName: 'parent',
            paths: [],
            namePlural: 'tableStyles',
            attachPayload: function (pageItem, style) {
                var table = pageItem.texts[0].tables.add();
                table.appliedTableStyle = style;
            },
            getFromDocument: function (doc) { return doc.allTableStyles },
            getChildren: function (group) { return group.tableStyles },
            getSubGroups: function (group) { return group.tableStyleGroups },
        },

        'Swatch': {
            constructorName: 'Swatch',
            documentProperty: 'swatches',
            groupConstructorName: 'ColorGroup',
            labelPlural: 'Swatches',
            parentName: 'parentColorGroup',
            paths: [],
            namePlural: 'swatches',

            // returns the children of group, if any, and requires this special function because
            // colorGroups are weird and don't just contain swatches in the simple way like the others
            getChildren: function (group) {

                if ('Document' === group.constructor.name)
                    return group.swatches;

                var swatches = [],
                    doc = 'Document' === group.constructor.name ? group : group.parent,
                    childrenNames = getProperties(group.colorGroupSwatches, 'swatchItemRef/name');

                for (var i = 0, swatch; i < childrenNames.length; i++)
                    if (swatch = getThing(doc.swatches, 'name', childrenNames[i]))
                        swatches.push(swatch);

                return swatches;

            },

            // get every one of the thing type from the document
            getFromDocument: function (container) { return container.swatches },

            // function that returns subgroups, if any of the container
            getSubGroups: function (container) { if ('undefined' !== typeof container.colorGroups) return container.colorGroups },

            // this function attaches the thing as a payload to the transfer frame
            attachPayload: function (frame, swatch) { frame.fillColor = swatch },

            // if returns false, will discard that path component
            filterPathComponent: function (str) { return -1 === indexOfArray(['[Root Color Group]', 'None', 'Registration', 'Paper', 'Black'], str) },

            // performs processing *after* the thing is copied
            // in this case it re-instates the ColorGroup of the swatch
            postProcess: function (sourceSwatch, destinationDoc, path) {

                var sourceDoc = sourceSwatch.parent,
                    components = path.split(' // ');

                if (1 === components.length)
                    // the swatch is not in a group, so ignore
                    return true;

                var destinationSwatch = getThing(destinationDoc.swatches, 'name', sourceSwatch.name);

                if (!destinationSwatch)
                    return alert('ThingType.Swatch.postProcess: failed to get destination swatch.');

                var sourceGroup = getThing(sourceDoc.colorGroups, 'name', components[0]);

                if (!sourceGroup)
                    return alert('ThingType.Swatch.postProcess: failed to get the source ColorGroup.');

                var destinationGroup = getThing(destinationDoc.colorGroups, 'name', components[0]);

                if (!destinationGroup)
                    destinationDoc.colorGroups.add(sourceGroup.name, [destinationSwatch], sourceGroup.properties);

                return true;

            },

        },

    };

    /* -------------------------- *
     *  COLLECT PATHS TO THINGS   *
     * -------------------------- */
    for (var key in ThingTypes)
        if (ThingTypes.hasOwnProperty(key))
            ThingTypes[key].paths = getPaths(settings.source[ThingTypes[key].documentProperty]);

    // set the active thing type
    settings.activeThingType = ThingTypes[settings.activeThingConstructorName];

    /* ----------- *
     *  SHOW UI    *
     * ----------- */
    if (
        settings.showUI
        && 2 === ui(settings)
    )
        // user cancelled UI
        return;

    if (!settings.activeThingType)
        return alert('Error: no active thing type.');

    /* ------------------ *
     *  COPY THE THINGS   *
     * ------------------ */
    var counter = 0,
        errors = [];

    for (var i = 0; i < settings.pathsToCopy.length; i++) {

        var thing = getThingByPath(doc, settings.activeThingConstructorName, settings.pathsToCopy[i], settings.pathDelimiter);

        if (!thing) {
            errors.push('Could not resolve path "' + settings.pathsToCopy[i] + '".');
            continue;
        }

        for (var j = 0; j < settings.destinations.length; j++)
            copyThing(settings.source, settings.destinations[j], thing) && counter++;

    }

    /* ------------------ *
     *  SHOW RESULTS      *
     * ------------------ */

    if (settings.showResults) {

        var message = 'Copied ' + counter + ' ' + settings.activeThingType.labelPlural + ' to ' + settings.destinations.length + ' documents.';

        if (errors.length)
            message += '\n' + errors.join('\n');

        return alert(message);

    }

    /**
     * Helper function to get paths for things.
     * @param {Array<*>} things - the things we want paths for.
     * @returns {Array<String>}
     */
    function getPaths(things) {

        var paths = [];

        for (var i = 0, path; i < things.length; i++) {

            if (!settings.showDefaultThings && things[i].name[0] === '[')
                // ignore default thing
                continue;

            path = getPathOfThing(settings.source, things[i], settings.pathDelimiter)

            if (path)
                paths.push(path);

        }

        return paths;

    };

};
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Copy Things');

/**
 * Given a document's thing, returns a string path,
 * based on the structure of groups in the document.
 *
 * Valid types of thing are:
 *   - ObjectStyle
 *   - ParagraphStyle
 *   - CharacterStyle
 *   - CellStyle
 *   - TableStyle
 * ---------------------------------------------------------------------------
 * Example output showing path of every object style:
 *
 *   '[None]'
 *   '[Basic Graphics Frame]'
 *   '[Basic Text Frame]'
 *   '[Basic Grid]'
 *   'Style Group 1/Style Group 2/Style Group 4/Object Style 1'
 *   'Style Group 1/Style Group 2/Style Group 4/Object Style 2'
 *   'Style Group 1/Style Group 2/Style Group 4/Object Style 3'
 *   'Style Group 1/Style Group 3/Style Group 6/Style Group 8/Object Style 4'
 *   'Style Group 1/Style Group 3/Style Group 6/Style Group 8/Object Style 5'
 *
 * ---------------------------------------------------------------------------
 * @author m1b
 * @version 2025-05-11
 * @param {Document} doc - an Indesign Document.
 * @param {*} thing - the thing to path.
 * @param {String} [delimiter] - the delimiter between path components (default: ' // ').
 * @returns {String}
 */
function getPathOfThing(doc, thing, delimiter) {

    delimiter = delimiter || ' // ';

    if (
        !doc
        || 'Document' !== doc.constructor.name
    )
        throw new Error('getPathOfThing: bad `doc` supplied. Expected Document.');

    if (
        !thing
        || !thing.constructor
        || !thing.name
    )
        throw new Error('getPathOfThing: bad `thing` supplied.');

    var thingType = ThingTypes[thing.constructor.name];

    if (!thingType)
        throw new Error('getPathOfThing: unknown `thing` supplied. (' + thing.constructor.name + ')');

    if (thingType.filterPathComponent && !thingType.filterPathComponent(thing.name))
        // the thingType rejected this one
        return;

    var components = [thing.name],
        target = thing;

    while (
        // looking for a parent component
        target.hasOwnProperty(thingType.parentName)
        && target[thingType.parentName].constructor.name === thingType.groupConstructorName
        // allow ThingType to exclude this component
        && (!thingType.filterPathComponent || thingType.filterPathComponent(target[thingType.parentName].name))
    ) {

        target = target[thingType.parentName];
        components.unshift(target.name);

    }

    return components.join(delimiter);

};

/**
 * Given a path, returns a thing from `doc`, based on the structure of groups
 * in the document. Supports: ObjectStyle, ParagraphStyle, CharacterStyle,
 * CellStyle, TableStyle, Swatch.
 * ---------------------------------------------------------------------------
 * Example 1 - get a specific thing:
 *
 *   var myStyle = getThingByPath(
 *                   doc,
 *                   'ObjectStyle',
 *                   'Style Group 1/Style Group 2/Style Group 4/Object Style 3'
 *   );
 *
 * ---------------------------------------------------------------------------
 * Example 2 - get all things in a group (important: trailing delimiter!):
 *
 *   var myStyles = getThingByPath(
 *                   doc,
 *                   'ObjectStyle',
 *                   'Style Group 1/Style Group 2/Style Group 4/'
 *   );
 *
 * ---------------------------------------------------------------------------
 * Example 3 - get a specific group (important: no trailing delimiter!):
 *
 *   var myStyleGroup = getThingByPath(
 *                   doc,
 *                   'ObjectStyle',
 *                   'Style Group 1/Style Group 2/Style Group 4'
 *   );
 *
 * ---------------------------------------------------------------------------
 * Example 4 - get all styles from document root (do not specify `path`):
 *
 *   var myRootStyles = getThingByPath(doc, 'ObjectStyle');
 *
 * ---------------------------------------------------------------------------
 * @author m1b
 * @version 2025-05-11
 * @param {Document} doc - an Indesign Document.
 * @param {String} constructorName - the thing's constructor's name, eg. 'ObjectStyle'.
 * @param {String} [path] - the path to navigate, eg. 'Group 1/Group 2/Thing' (default: get things from document root).
 * @param {String} [delimiter] - the delimiter between path components (default: ' // ').
 * @returns {*|Array<*>}
 */
function getThingByPath(doc, constructorName, path, delimiter) {

    var components = (path || '').split(delimiter || ' // ');

    if (
        !doc
        || 'Document' !== doc.constructor.name
    )
        throw new Error('getThingByPath: bad `doc` supplied. Expected Document.');

    if (!components)
        throw new Error('getThingByPath: bad `path` supplied. (' + path + ')');

    var thingType = ThingTypes[constructorName];

    if (!thingType)
        throw new Error('getThingByPath: unknown `constructorName` supplied. (' + constructorName + ')');

    var target = 1 === components.length ? doc : thingType.getSubGroups(doc),
        component,
        children,
        subGroups,
        found = [];

    while (components.length) {

        component = components.shift()

        if (components.length > 0) {

            // change target to the next group
            target = target.itemByName(component);

            if (!target.isValid)
                return;

            subGroups = thingType.getSubGroups(target)

            if (subGroups && subGroups.length)
                // change target to sub groups
                target = subGroups;

            continue;

        }

        // now deal with the last component
        if (!component) {

            // no specific thing specified
            // only the group, so get all things
            found = thingType.getChildren(target).everyItem().getElements();

        }

        else {

            // the specific thing specified in path
            children = thingType.getChildren(target);

            if (children && children.length)
                found = getThing(children, 'name', component);

            if (!found || !found.isValid)
                return;

        }

    }

    return found;

};


/**
 * Copy a thing from source document to destination document(s).
 * If the thing already exists, it will be replaced.
 * @author m1b
 * @version 2025-05-11
 * @param {Document} source - the document containing the source style.
 * @param {Documents|Array<Document>|Document} destinations - the destination document(s).
 * @param {ObjectStyle|ParagraphStyle|CharacterStyle} thing - the thing to copy.
 * @param {TextFrame} [tempFrame] - a temporary text frame as a vehicle for duplicating the thing (default: will create it).
 * @returns {ObjectStyle|ParagraphStyle|CharacterStyle} - the copied thing.
 */
function copyThing(source, destinations, thing, tempFrame) {

    if ('Document' === destinations.constructor.name)
        destinations = [destinations];

    else if (
        'Array' !== destinations.constructor.name
        && 'Documents' !== destinations.constructor.name
    )
        throw Error('copyThing: bad `doc` supplied.');

    // functions to get and set the thing
    var thingType = ThingTypes[thing.constructor.name];

    if (!thingType)
        return alert('Cannot handle "' + thing.constructor.name + '".');

    // make sure we have access to the layer (we'll restore this later)
    var sourceLayerState = prepareLayer(source.layoutWindows[0].activeLayer);

    try {

        // make a temporary text frame
        var tempFrame = source.textFrames.add({ contents: 'a' });

        // attach the thing to the text frame so that
        // when we duplicate it, the payload goes along
        thingType.attachPayload(tempFrame, thing);

        for (var i = 0; i < destinations.length; i++) {

            var doc = destinations[i];

            if (
                doc === source
                || 0 === doc.layoutWindows.length
            )
                continue;

            if ('Document' !== doc.constructor.name)
                throw Error('copyThing: bad `docs` supplied.');

            // does the thing already exist in the destination document?
            var pathToThing = getPathOfThing(source, thing);
            var existingThing = getThingByPath(doc, thing.constructor.name, pathToThing);

            if (existingThing) {

                try {
                    // rename to avoid collision
                    existingThing.name = '__deleteMe__';
                } catch (error) {
                    // this will be caught and displayed by the higher try/catch
                    throw new Error('copyThing: could not replace "' + thing.name + '" in "' + doc.name + '".');
                }
            }

            try {
                // duplicate it to the destination document
                var destinationFrame = tempFrame.duplicate(doc.layoutWindows[0].activePage);
            }
            catch (error) {
                // this will be caught and displayed by the higher try/catch
                throw new Error('copyThing: error duplicating "' + thing.name + '" to "' + doc.name + '".');
            }

            if (
                thingType.postProcess
                // some types need some extra work done here
                && true !== thingType.postProcess(thing, doc, pathToThing)
            )
                // post processing failed, but it should have displayed any errors already
                return;

            var destinationThing = getThingByPath(doc, thing.constructor.name, pathToThing);

            if (!destinationThing)
                return alert('copyThing: "' + thing.name + '" did not copy to "' + doc.name + '". (Path: ' + pathToThing + ')');

            if (existingThing)
                // replace with the duplicated thing
                existingThing.remove(destinationThing)

            // clean up
            var destinationLayer = destinationFrame.itemLayer,
                originalDestinationLayerState = prepareLayer(destinationLayer);
            destinationFrame.remove()
            destinationLayer.properties = originalDestinationLayerState;

        }

    }

    catch (error) {
        return alert(error.message);
    }

    finally {

        if (tempFrame && tempFrame.isValid)
            // clean up
            tempFrame.remove();

        source.layoutWindows[0].activeLayer.properties = sourceLayerState;

    }


    return destinationThing;

};

/**
 * Unlocks and shows a layer in preparation for DOM operation.
 *
 * Example to ensure we have access to remove `item`
 * and then reinstate the layer properties afterwards:
 *    var layer = item.itemLayer;
 *    var originalLayerState = prepareLayer(layer);
 *    item.remove();
 *    layer.properties = originalLayerState;
 *
 * @author m1b
 * @version 2025-05-11
 * @param {Layer} layer - the layer to prepare.
 * @returns {Object} - the previous settings {locked: visible:}.
 */
function prepareLayer(layer) {

    var originalSettings = { locked: layer.locked, visible: layer.visible };

    layer.locked = false;
    layer.visible = true;

    return originalSettings;

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
 * Returns a thing with matching property.
 * @param {Array|collection} things - the things to look through, eg. PageItems.
 * @param {String} key - the property name, eg. 'name'.
 * @param {*} value - the value to match.
 * @returns {*} - the thing.
 */
function getThing(things, key, value) {

    for (var i = 0; i < things.length; i++)
        if (things[i][key] == value)
            return things[i];

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
 * Shows UI for Copy Things script.
 * NOTE: expects to access `ThingTypes` global.
 * @author m1b
 * @version 2025-05-11
 * @param {Object} settings - the settings object.
 * @returns {1|2}
 */
function ui(settings) {

    var openDocuments = getOpenVisibleDocuments(),
        listBoxWidth = 700,
        listBoxHeight = 400;

    var w = new Window("dialog { text:'Copy Things', properties:{ resizeable:false } }"),

        typesPanel = w.add("tabbedpanel {alignment:['fill','top']  }"),

        group = w.add('group {orientation:"row", alignChildren:["fill","fill"], margins:[12,10,12,10], scrolling:true }'),

        thingsGroup = group.add('group {orientation:"column", alignChildren:["fill","fill"], margins:[0,0,15,0] }'),
        abovePathsRow = thingsGroup.add('group {orientation:"row", alignChildren:["fill","fill"] }'),
        pathsLabel = abovePathsRow.add("statictext { text:'', alignment:['left','center'], justify:'left' }"),
        filterGroup = abovePathsRow.add('group {orientation:"row" }'),
        filterLabel = filterGroup.add("statictext { text:'Search:', alignment:['right','bottom'], justify:'right',preferredSize: [-1,24] }"),
        filterField = filterGroup.add('edittext { text: "", alignment:["right","bottom"], preferredSize: [140,24] }'),
        clearButton = filterGroup.add("Button { text:'\u2715', alignment:['right','center'], preferredSize:[21,-1] }"),
        pathsListBox = thingsGroup.add("ListBox {alignment:['fill','fill'], preferredSize:[1000,700], properties:{multiselect:true, showHeaders:false, numberOfColumns:1, scrolling:true } }"),

        destinationsGroup = group.add('group {orientation:"column", alignChildren:["fill","fill"] }'),
        destinationLabel = destinationsGroup.add("statictext { text:'Copy to document(s):', alignment:['fill','bottom'], preferredSize:[-1,33] }"),
        destinationsListBox = destinationsGroup.add("ListBox {alignment:['fill','fill'], preferredSize:[400,-1], properties:{multiselect:true, showHeaders:false, numberOfColumns:1, scrolling:true } }"),

        buttons = w.add("Group { orientation: 'row', alignment: ['right','bottom'], margins:[0,0,10,10] }"),
        cancelButton = buttons.add("Button { text:'Cancel', properties: { name:'cancel' } }"),
        okButton = buttons.add("Button { text:'Copy', properties: { name:'ok' } }");

    pathsLabel.preferredSize.width = Math.floor(listBoxWidth * 0.6);
    pathsListBox.preferredSize = [listBoxWidth, listBoxHeight];
    destinationsListBox.preferredSize.height = listBoxHeight;

    // add tab for each thing type
    for (var key in ThingTypes)
        if (ThingTypes.hasOwnProperty(key))
            typesPanel.add('tab', undefined, ThingTypes[key].labelPlural);

    // populate destinations listbox
    for (var i = 0; i < openDocuments.length; i++)
        if (openDocuments[i] !== settings.source)
            destinationsListBox.add('item', openDocuments[i].name);

    // set the correct tab on the types panel
    for (var i = 0; i < typesPanel.children.length; i++)
        if (typesPanel.children[i].text === settings.activeThingType.labelPlural)
            typesPanel.selection = i;

    /** handle changes to the types panel, ie. user changes the active ThingType */
    typesPanel.onChange = function () {

        if (!typesPanel.selection)
            return;

        settings.activeThingType = getThingType('labelPlural', typesPanel.selection.text);
        settings.activeThingConstructorName = settings.activeThingType.constructorName;

        updatePaths();
        updateUI();

    };

    /** handle changes to the paths listbox, ie. user selects path(s) */
    pathsListBox.onChange = function () {

        var sel = this.selection;
        settings.pathsToCopy = [];

        if (sel)
            for (var i = 0; i < sel.length; i++)
                settings.pathsToCopy.push(sel[i].text);

        updateUI();

    };

    /** handle changes to the destinations listbox, ie. user selects document(s) */
    destinationsListBox.onChange = function () {

        var sel = this.selection;
        settings.destinations = [];

        if (sel)
            for (var i = 0; i < sel.length; i++)
                settings.destinations.push(app.documents.itemByName(sel[i].text));

        updateUI();

    };

    /** clears the filter field */
    clearButton.onClick = function () {
        filterField.text = '';
        filterField.onChanging();
    };

    filterField.onChanging = updatePaths;

    /** updates the contents of the paths listbox, including filtering */
    function updatePaths() {

        var selectedPaths = getProperties(pathsListBox.selection || [], 'text'),
            paths = settings.activeThingType.paths,
            newSelection = [],
            filteredPaths = [];

        if (0 === filterField.text.length)
            filteredPaths = paths.slice();

        else

            for (var i = 0; i < paths.length; i++)
                if (-1 !== paths[i].toLowerCase().search(filterField.text.toLowerCase()))
                    // matched with the filter field text
                    filteredPaths.push(paths[i]);

        // clear the current paths
        pathsListBox.removeAll();

        // populate with the thingType's paths
        for (var i = 0; i < filteredPaths.length; i++) {
            pathsListBox.add('item', filteredPaths[i]);

            // reinstate the selection, if possible
            for (var j = 0; j < selectedPaths.length; j++)
                if (selectedPaths[j] === filteredPaths[i])
                    newSelection.push(i);

        }

        pathsListBox.selection = newSelection;

        updateUI();

    };

    // init layout
    typesPanel.onChange();
    w.layout.layout();

    // cosmetic
    clearButton.size = [21, 21];
    clearButton.location[1] += 7;
    typesPanel.minimumSize[1] = 35;
    typesPanel.size[1] = 35;

    // show the dialog
    w.center();
    return w.show();

    /** update the UI from settings */
    function updateUI() {

        pathsLabel.text = pathsListBox.children.length > 0
            ? settings.activeThingType.labelPlural + ':'
            : 'No ' + settings.activeThingType.labelPlural + ' to show';

        destinationLabel.text = destinationsListBox.children.length > 0
            ? (pathsListBox.selection
                ? 'Copy ' + pathsListBox.selection.length + ' ' + settings.activeThingType.labelPlural
                + (destinationsListBox.selection ? ' to ' + destinationsListBox.selection.length + ' documents' : '')
                : 'Nothing to copy'
            )
            : 'No documents to show';

        okButton.enabled = (
            settings.pathsToCopy.length > 0
            && settings.destinations.length > 0
        );

        clearButton.enabled = filterField.text.length > 0;

    };

    /** Helper: returns the ThingType with property `propName` matching `str` */
    function getThingType(propName, str) {
        for (var key in ThingTypes)
            if (ThingTypes[key][propName] === str)
                return ThingTypes[key];
    };

};