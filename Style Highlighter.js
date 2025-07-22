/**
 * @file Style Highlighter.js
 *
 * Highlights all text with applied paragraph or character style(s).
 *
 * Works by creating "Highlighter" conditions and
 * applying to all texts found in the selected style(s).
 *
 * It is safe to use because it will not modify the styles, but
 * rather applies a "highlighter" condition (in addition to
 * any existing conditions applied).
 *
 * Highlighting is removed by removing the highlighter condition(s),
 * and the script provides a "Remove All Highlighters" option
 * to make this quick and easy.
 *
 * TIP: it is okay to manually remove highlighter conditions.
 *
 * @author m1b
 * @version 2025-06-17
 */
function main() {

    // the source document
    var doc = app.activeDocument;

    var settings = {

        // user settings
        activeThingConstructorName: 'ParagraphStyle',
        highlighterConditionNamePrefix: '_Highlighter ',
        pathDelimiter: ' > ',
        showDefaultThings: false,
        showResults: true,

        // highlight colors
        highlights: [

            { name: 'None', rgb: [99, 99, 99] },
            { name: 'Fiesta', rgb: [255, 85, 85] },
            { name: 'Brick Red', rgb: [203, 90, 90] },
            { name: 'Orange', rgb: [255, 165, 0] },
            { name: 'Gold', rgb: [255, 215, 0] },
            { name: 'Sulphur', rgb: [255, 255, 102] },
            { name: 'Pale Yellow', rgb: [255, 255, 153] },
            { name: 'Light Green', rgb: [153, 255, 153] },
            { name: 'Grass Green', rgb: [124, 252, 0] },
            { name: 'Green', rgb: [50, 168, 50] },
            { name: 'Cute Teal', rgb: [50, 178, 178] },
            { name: 'Cyan', rgb: [102, 255, 255] },
            { name: 'Blue', rgb: [0, 122, 255] },
            { name: 'Lavender', rgb: [204, 153, 255] },
            { name: 'Violet', rgb: [238, 130, 238] },
            { name: 'Magenta', rgb: [255, 0, 255] },
            { name: 'Lipstick', rgb: [222, 20, 84] },

        ],

        // internal settings
        allPaths: [],
        pathsToHighlight: [],
        highlight: undefined,
        doc: doc,
        activeThingType: undefined,
    };

    // default highlight:
    if (!settings.highlight)
        settings.highlight = settings.highlights[2];

    /**
     * Info about the things we can copy.
     * Access using the thing's constructor's name.
     * (See `Swatch` for special cases with documentation.)
     */
    $.global.ThingTypes = {

        'CharacterStyle': {
            constructorName: 'CharacterStyle',
            documentProperty: 'allCharacterStyles',
            grepProperty: 'appliedCharacterStyle',
            groupConstructorName: 'CharacterStyleGroup',
            labelPlural: 'Character Styles',
            parentName: 'parent',
            paths: [],
            getChildren: function (group) { return group.characterStyles },
            getSubGroups: function (group) { return group.characterStyleGroups },
        },

        'ParagraphStyle': {
            constructorName: 'ParagraphStyle',
            documentProperty: 'allParagraphStyles',
            grepProperty: 'appliedParagraphStyle',
            groupConstructorName: 'ParagraphStyleGroup',
            labelPlural: 'Paragraph Styles',
            parentName: 'parent',
            paths: [],
            getChildren: function (group) { return group.paragraphStyles },
            getSubGroups: function (group) { return group.paragraphStyleGroups },
        },

    };

    // populate the ThingTypes paths for the source document
    for (var key in ThingTypes)
        if (ThingTypes.hasOwnProperty(key))
            ThingTypes[key].paths = getPaths(settings.doc[ThingTypes[key].documentProperty]);

    settings.activeThingType = ThingTypes[settings.activeThingConstructorName];

    // add functions used by the UI
    settings.highlightTextsInStyle = highlightTextsInStyle;
    settings.removeAllHighlights = removeAllHighlighters;

    // show UI
    var result = ui(settings);

    if (2 === result)
        // user cancelled UI
        return;

    // change the screen mode so the highlighting is visible
    doc.layoutWindows[0].screenMode = ScreenModeOptions.PREVIEW_OFF;


    function highlightText(settings) {

        /* ----------------------------- *
         *  HIGHLIGHT THE STYLED TEXTS   *
         * ----------------------------- */

        if (!settings.activeThingType)
            return alert('Error: no active thing type.');

        if (!settings.highlight)
            return alert('Error: no highlight chosen.');

        if (0 === settings.pathsToHighlight.length)
            return alert('Error: styles chosen.');

        for (var i = 0; i < settings.pathsToHighlight.length; i++)
            highlightTextsInStyle(settings.pathsToHighlight[i]);

        // change the screen mode so the highlighting is visible
        doc.layoutWindows[0].screenMode = ScreenModeOptions.PREVIEW_OFF;

    };

    function removeAllHighlighters(settings) {

        /* ------------------------------------ *
         *  REMOVE ALL HIGHLIGHTER CONDITIONS   *
         * ------------------------------------ */
        debugger; // 2025-07-02

        for (var i = doc.conditions.length - 1; i >= 0; i--)
            if (0 === doc.conditions[i].name.indexOf(settings.highlighterConditionNamePrefix))
                doc.conditions[i].remove();

    };

    /**
     * Adds a highlighter condition to all text found in the given style.
     * @author m1b
     * @version 2025-06-17
     * @param {String} pathToStyle - the path to the style.
     */
    function highlightTextsInStyle(pathToStyle) {

        var style = getThingByPath(doc, settings.activeThingConstructorName, pathToStyle, settings.pathDelimiter);
        var condition = getHighlighterCondition(doc, settings.highlight);

        // reset grep prefs
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences[settings.activeThingType.grepProperty] = style;

        var found = doc.findGrep();

        for (var i = 0; i < found.length; i++)
            found[i].applyConditions([condition], false);

    };

    /**
     * Retrieves or creates a "highlighter" condition from `doc`.
     * @author m1b
     * @version 2025-06-17
     * @param {Document} doc - an Indesign Document.
     * @param {Object} highlight - the highlight object, eg. { name: 'Fiesta', rgb: [255, 85, 85] }
     * @returns {Condition}
     */
    function getHighlighterCondition(doc, highlight) {

        var conditionName = settings.highlighterConditionNamePrefix + highlight.name;
        var condition = doc.conditions.itemByName(conditionName);

        if (!condition.isValid) {

            // make it
            condition = doc.conditions.add({
                name: conditionName,
                indicatorColor: highlight.rgb,
                indicatorMethod: ConditionIndicatorMethod.USE_HIGHLIGHT,
                visible: true,
            });

        }

        return condition;

    };

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

            path = getPathOfThing(settings.doc, things[i], settings.pathDelimiter)

            if (path)
                paths.push(path);

        }

        return paths;

    };

};
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Style Highlighter');

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
 * Shows UI for Style Highlighter script.
 * NOTE: expects to access `ThingTypes` global.
 * @author m1b
 * @version 2025-06-17
 * @param {Object} settings - the settings object.
 * @returns {1|2}
 */
function ui(settings) {

    const LISTBOX_WIDTH = 700,
        LISTBOX_HEIGHT = 200;

    var w = new Window("dialog { text:'Style Highlighter', properties:{ resizeable:false } }"),

        typesPanel = w.add("tabbedpanel {alignment:['fill','top']  }"),

        group = w.add('group {orientation:"row", alignChildren:["fill","fill"], margins:[12,10,12,0], scrolling:true }'),

        thingsGroup = group.add('group {orientation:"column", alignChildren:["fill","fill"], margins:[0,0,0,0] }'),
        abovePathsRow = thingsGroup.add('group {orientation:"row", alignChildren:["fill","fill"] }'),
        pathsLabel = abovePathsRow.add("statictext { text:'', alignment:['left','center'], justify:'left' }"),
        filterGroup = abovePathsRow.add('group {orientation:"row" }'),
        filterLabel = filterGroup.add("statictext { text:'Search:', alignment:['right','bottom'], justify:'right',preferredSize: [-1,24] }"),
        filterField = filterGroup.add('edittext { text: "", alignment:["right","bottom"], preferredSize: [140,24] }'),
        clearButton = filterGroup.add("Button { text:'\u2715', alignment:['right','center'], preferredSize:[21,-1] }"),
        pathsListBox = thingsGroup.add("ListBox {properties:{multiselect:true, showHeaders:false, numberOfColumns:1, scrolling:true } }"),

        colorSection = w.add("Group { orientation: 'row', alignment: ['fill','fill'], margins:[10,10,10,0] }"),
        highlightButtons = colorSection.add("Group { orientation: 'row', alignment: ['fill','center'], alignChildren:['left','center'], preferredSize:[-1,60] }"),
        highlightSection = colorSection.add("Group { orientation: 'column', alignment: ['fill','fill'], margins:[0,0,0,0] }"),
        activeColorPanel = highlightSection.add("Group { orientation: 'row', alignment: ['center','center'], preferredSize:[60,60] }"),
        highlightLabel = activeColorPanel.add("statictext { text:'Highlight color:', alignment:['fill','fill'], justify:'center', preferredSize:[-1,60] }"),
        highlightButton = highlightSection.add("Button { text:'Add Highlighter', properties: {} }"),

        bottomUI = w.add("Group { orientation: 'row', alignment: ['fill','bottom'], margins:[10,10,10,10] }"),
        altButtons = bottomUI.add("Group { orientation: 'row', alignment: ['left','bottom'] }"),
        removeAllHighlightersButton = altButtons.add("Button { text:'Remove All Highlighters'}"),
        buttons = bottomUI.add("Group { orientation: 'row', alignment: ['right','bottom'] }"),
        // cancelButton = buttons.add("Button { text:'Cancel', properties: { name:'cancel' } }"),
        doneButton = buttons.add("Button { text:'Done', properties: { name:'ok' } }");

    pathsLabel.preferredSize.width = Math.floor(LISTBOX_WIDTH * .6);
    pathsListBox.preferredSize = [LISTBOX_WIDTH, LISTBOX_HEIGHT];

    // add tab for each thing type
    for (var key in ThingTypes)
        if (ThingTypes.hasOwnProperty(key))
            typesPanel.add('tab', undefined, ThingTypes[key].labelPlural);

    // populate highlight color chips
    for (var i = 0, col; i < settings.highlights.length; i++)
        addColorChip(highlightButtons, settings.highlights[i].rgb, setHighlight, i);

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
        settings.pathsToHighlight = [];

        if (sel)
            for (var i = 0; i < sel.length; i++)
                settings.pathsToHighlight.push(sel[i].text);

        updateUI();

    };

    highlightButton.onClick = function () {

        /* ----------------------------- *
         *  HIGHLIGHT THE STYLED TEXTS   *
         * ----------------------------- */

        if (!settings.activeThingType)
            return alert('Error: no active thing type.');

        if (!settings.highlight)
            return alert('Error: no highlight chosen.');

        if (0 === settings.pathsToHighlight.length)
            return alert('Error: styles chosen.');

        for (var i = 0; i < settings.pathsToHighlight.length; i++)
            settings.highlightTextsInStyle(settings.pathsToHighlight[i]);

        // change the screen mode so the highlighting is visible
        settings.doc.layoutWindows[0].screenMode = ScreenModeOptions.PREVIEW_OFF;

    };

    removeAllHighlightersButton.onClick = settings.removeAllHighlighters;

    /** clears the filter field */
    clearButton.onClick = function () {
        filterField.text = '';
        filterField.onChanging();
    };

    filterField.onChanging = updatePaths;

    // init layout
    typesPanel.onChange();
    w.layout.layout();

    // cosmetic
    clearButton.size = [21, 21];
    clearButton.location[1] += 5;
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

        highlightButton.enabled = (
            settings.pathsToHighlight.length > 0
            && undefined != settings.highlight
            && 'None' !== settings.highlight.name
        );

        setActiveHighlight(settings.highlight);

        clearButton.enabled = filterField.text.length > 0;

        removeAllHighlightersButton.enabled = hasHighlighterConditions(settings.doc);

    };

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

    /** sets the active highlight in settings */
    function setHighlight() {
        settings.highlight = settings.highlights[this.index];
        updateUI();
    };

    /** sets the highlight in the UI */
    function setActiveHighlight(highlight) {

        settings.highlight = highlight;

        if (!highlight)
            highlight = settings.highlights[settings.highlights.length - 1];

        var color = highlight.rgb;

        activeColorPanel.graphics.backgroundColor = group.graphics.newBrush(
            group.graphics.BrushType.SOLID_COLOR, [color[0] / 255, color[1] / 255, color[2] / 255, 1]
        );

        highlightLabel.text = highlight.name;

    };

    /** Helper: returns the ThingType with property `propName` matching `str` */
    function getThingType(propName, str) {
        for (var key in ThingTypes)
            if (ThingTypes[key][propName] === str)
                return ThingTypes[key];
    };

    /** returns true when doc contains highlighter conditions */
    function hasHighlighterConditions(doc) {

        for (var i = 0; i < doc.conditions.length; i++)
            if (doc.conditions[i].name.match(settings.highlighterConditionNameMatcher))
                return true;

        return false;

    };

    /**
     * Adds a colored square SUI group to `container`.
     * @author m1b
     * @version 2025-06-17
     * @param {Window|Group|Panel} container - the container of the squares.
     * @param {Array<Number>} color - rgb breakdown, eg. [255,0,0].
     * @param {Function} clickFunction - the onClick event function for each square.
     * @param {Number} index - an identifying index for this chip.
     * @returns {SUI Group}
     */
    function addColorChip(container, color, clickFunction, index) {

        var group = container.add("group {margin: [0,0,0,0] }");
        group.size = [20, 20];
        group.index = index;
        group.graphics.backgroundColor = group.graphics.newBrush(
            group.graphics.BrushType.SOLID_COLOR, [color[0] / 255, color[1] / 255, color[2] / 255, 1]
        );

        group.addEventListener("click", clickFunction);

        if (0 === index) {

            // the NONE color
            group.onDraw = function () {
                var g = this.graphics;

                g.backgroundColor = group.graphics.newBrush(
                    g.BrushType.SOLID_COLOR, [1, 1, 1, 1]
                );

                // Create a black pen
                var blackPen = g.newPen(
                    g.PenType.SOLID_COLOR,
                    [0, 0, 0, 1], // Black (R,G,B,A)
                    1 // 1px line width
                );

                g.pen = blackPen;

                // Draw a line from bottom-left to top-right
                g.moveTo(0, this.size[1]);           // (x=0, y=height)
                g.lineTo(this.size[0], 0);           // (x=width, y=0)
            }

        }

        return group;

    };

};