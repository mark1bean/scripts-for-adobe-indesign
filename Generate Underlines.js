/**
 * @file Generate Underlines.js
 *
 * Usage:
 *
 * 1. Select text and set Conditions based on the `underlinerConfig` array in the script:
 *
 *    (a) In your document, create a Condition called "Underline Black" and apply
 *        that condition to any text you want to underline with the "Underline Black"
 *        configuration in the script.
 *
 *    (b) Also create an "Underline Black" object style.
 *
 *    (c) Do (a) and (b) for other config entries, eg. "Underline Blue" and "Underline Green".
 *
 * 2. Run script. It will draw the underlines (as path items) into a group called "overlay ###".
 *
 * TIP: If the text reflows, re-run the script. It will replace the outdated underlines.
 *
 * @author m1b
 * @version 2025-10-19
 * @discussion https://community.adobe.com/t5/indesign-discussions/multiple-underline/m-p/15416753
 */
const mm = 2.834645;

function main() {

    /**
     * Underline configurations:
     *   - key: the name of both (a) a Condition representing that
     *     type of underline and (b) an Object style to apply
     *     to the drawn path item.
     *   - verticalOffset: the offset, in points, from the baseline
     *     of the text (positive is down).
     *   - objectStyle: will be populated later from the active document.
     */
    var underlinerConfig = [
        { key: 'Underline Black', verticalOffset: 0.5 * mm, objectStyle: undefined },
        { key: 'Underline Blue', verticalOffset: 1.25 * mm, objectStyle: undefined },
        { key: 'Underline Green', verticalOffset: 2 * mm, objectStyle: undefined },
        { key: 'Underline Gold', verticalOffset: 2.75 * mm, objectStyle: undefined },
    ];

    var settings = {
        /** whether to apply underlining to every text frame of the selected texts' stories */
        generateUnderlinesForWholeStories: false,
    };

    if (
        0 === app.documents.length
        || 0 === app.activeDocument.selection.length
        || !app.activeDocument.selection[0].hasOwnProperty('texts')
    )
        return alert('Please select some text and try again.');

    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;

    // map the underlinerConfig for convenience
    var underlinerConfigMap = {};
    for (var i = 0; i < underlinerConfig.length; i++)
        underlinerConfigMap[underlinerConfig[i].key] = underlinerConfig[i];

    var doc = app.activeDocument;
    var frames = settings.generateUnderlinesForWholeStories
        ? getStoryTextContainers(doc.selection)
        : getTextContainers(doc.selection);

    if (0 === frames.length)
        return alert('Please select some text and try again.');

    // check the configuration is good
    for (var i = 0, style; i < underlinerConfig.length; i++) {

        style = getThing(doc.allObjectStyles, 'name', underlinerConfig[i].key);

        if (!style)
            return alert('Configuration error: could not find Object Style "' + underlinerConfig[i].key + '".');

        underlinerConfig[i].objectStyle = style;

        if (isNaN(underlinerConfig[i].verticalOffset))
            return alert('Configuration error: bad `verticalOffset` for "' + underlinerConfig[i].key + '".');

    }

    app.scriptPreferences.enableRedraw = false;

    /* --------------------------------------------- *
     *  Instantiate an Overlay for each text frame   *
     * --------------------------------------------- */
    for (var o = 0, overlay; o < frames.length; o++) {

        overlay = new Overlay(frames[o], "Underlines", true);

        if (!overlay.isValid)
            // frame was probably on pasteboard
            continue;

        // collect the texts needed for this overlay
        overlay.texts = collectConditionalTextLines(overlay.frame);

        /* ---------------------- *
         *  Draw the underlines   *
         * ---------------------- */

        for (var i = 0, text, cond, style, offset; i < overlay.texts.length; i++) {

            text = overlay.texts[i];

            conditionsLoop:
            for (var j = 0; j < text.appliedConditions.length; j++) {

                cond = text.appliedConditions[j];
                style = (underlinerConfigMap[cond.name] || 0).objectStyle;
                offset = (underlinerConfigMap[cond.name] || 0).verticalOffset;

                if (!style)
                    // no style for this condition
                    continue conditionsLoop;

                // draw the underline
                var line = overlay.group.polygons.add({ appliedObjectStyle: style });
                var x0 = text.characters.firstItem().horizontalOffset;
                var x1 = text.characters.lastItem().endHorizontalOffset;
                var y = text.characters.firstItem().baseline + offset;
                line.paths.item(0).entirePath = [[x0, y], [x1, y]];
                line.clearObjectStyleOverrides();

            }

        }

    }

};
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Generate Underlines');

/**
 * Returns lines of text with applied conditions.
 * @author m1b
 * @version 2025-07-18
 * @param {TextFrame} frame - the frame to collect texts from.
 * @returns {Array<Text>}
 */
function collectConditionalTextLines(frame) {

    /* -------------------------------------- *
     *  Collect all the underlineable texts   *
     * -------------------------------------- */
    var texts = [];

    for (var i = 0, tsr, line, len = frame.textStyleRanges.length; i < len; i++) {

        tsr = frame.textStyleRanges[i];

        if (0 === tsr.appliedConditions.length)
            continue;

        if (1 === tsr.lines.length) {
            texts.push(tsr);
            continue;
        }

        for (var j = 0; j < tsr.lines.length; j++) {

            line = tsr.lines[j];

            if (0 === j)
                texts.push(tsr.characters.itemByRange(tsr.characters.firstItem(), line.characters.lastItem()).getElements()[0]);

            else if ((tsr.lines.length - 1) === j)
                texts.push(tsr.characters.itemByRange(line.characters.firstItem(), tsr.characters.lastItem()).getElements()[0]);

            else
                texts.push(tsr.characters.itemByRange(line.characters.firstItem(), line.characters.lastItem()).getElements()[0]);

        }

    }

    return texts;

};

/**
 * Returns a layer `name`, making it if necessary.
 * @param {Document} doc - an Indesign Document.
 * @param {String} name - the name of the layer.
 * @returns {Layer}
 */
function getLayer(doc, name) {

    var layer = getThing(doc.layers, 'name', name) || doc.layers.add({ name: name });

    if (layer.isValid)
        return layer;

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
 * Returns page item on page by matching label.
 * @author m1b
 * @version 2023-04-25
 * @param {Page} page - the page to search.
 * @param {String} label - the label to search for.
 * @param {Array<PageItem>}
 */
function getPageItemsByLabel(page, label) {

    var found = [];
    var items = page.allPageItems;

    for (var i = 0; i < items.length; i++)
        if (items[i].label === label)
            found.push(items[i]);

    return found;

};

/**
 * Returns all textFrame containers of item's parent stories.
 * Note: even if `items` only contains one text frame of a multi-frame
 * story, this function will return all of the containers for that story.
 * @author m1b
 * @version 2025-07-19
 * @param {Array<PageItem>} items - the items to search.
 * @returns {Array<TextFrame>}
 */
function getStoryTextContainers(items) {

    var frames = [];
    var done = {};

    for (var i = 0, story; i < items.length; i++) {

        if (!items[i].hasOwnProperty('texts'))
            continue;

        story = items[i].texts[0].parentStory;

        if (!story || done[story.id])
            continue;

        frames = frames.concat(story.textContainers);
        done[story.id] = true;

    }

    return frames;

};

/**
 * Returns textFrame containers of parent stories from items.
 * @author m1b
 * @version 2025-07-22
 * @param {Array<PageItem>} items - the items to search.
 * @returns {Array<TextFrame>}
 */
function getTextContainers(items) {

    var frames = [];
    var done = {};

    for (var i = 0, textFrames; i < items.length; i++) {

        if (
            !items[i].hasOwnProperty('texts')
            || !items[i].texts[0].hasOwnProperty('parentTextFrames')
        )
            continue;

        textFrames = items[i].texts[0].parentTextFrames;

        for (var j = 0; j < textFrames.length; j++) {

            if (!textFrames[j] || done[textFrames[j].id])
                continue;

            frames.push(textFrames[j]);
            done[textFrames[j].id] = true;

        }

    }

    return frames;

};

/**
 * A text frame overlay. A simple conceptual helper object.
 * Handles removal of previous overlays, and making an overlay
 * group for storing overlay-related DOM items.
 * @author m1b
 * @version 2025-07-19
 * @constructor
 * @param {TextFrame} frame - the text frame to overlay.
 * @param {Boolean} behind - whether to place the overlay behind the text frame (default: false - place in front).
 */
function Overlay(frame, name, behind) {

    this.frame = frame;

    if (!this.frame.parentPage) {
        // no overlays on pasteboard
        this.isValid = false;
        return;
    }

    this.page = frame.parentPage;

    // identifier for this overlay
    this.name = name;
    this.id = this.frame.id;
    this.frame.name = 'text ' + this.id;

    // remove previous overlay groups
    var oldGroups = getPageItemsByLabel(this.page, String(this.id));

    for (var i = oldGroups.length - 1; i >= 0; i--)
        oldGroups[i].remove();

    // make a group for this overlay
    this.group = this.page.groups.add([
        this.frame.parent.rectangles.add({ geometricBounds: this.frame.geometricBounds, filled: false, stroked: false }),
        this.frame.parent.rectangles.add({ geometricBounds: this.frame.geometricBounds, filled: false, stroked: false }),
    ], undefined, true === behind ? LocationOptions.AFTER : LocationOptions.BEFORE, this.frame,);

    // delete unnecessary rectangle
    this.group.rectangles[0].remove();

    // set the group rectangle to no fill no stroke and no styling
    var noneSwatch = app.activeDocument.swatches.itemByName(app.translateKeyString("$ID/None"));
    this.group.rectangles[0].fillColor = noneSwatch;
    this.group.rectangles[0].strokeColor = noneSwatch;
    this.group.rectangles[0].strokeWeight = 0;
    this.group.rectangles[0].appliedObjectStyle = app.activeDocument.objectStyles.item(0);

    // identifier for the group
    this.group.label = '' + this.id;
    this.group.name = this.name + ' ' + this.id;

    this.isValid = true;

};