/**
 * @file Synchronize All Documents Text Selection.js
 *
 * Attempts to move the active view of each open document
 * so that it shows the same text as the selected text.
 *
 * The idea is that when working on similar documents
 * that share a piece of text, we want to quickly
 * navigate to those places in each of the open documents.
 *
 * @author m1b
 * @version 2025-03-15
 */
function main() {

    var settings = {

        /* when these properties are true, will use the
         * target text's properties in the find preferences.
         * - you can add any other valid properties here, too.
         */
        findTextPreferences: {
            appliedCharacterStyle: true,
            appliedParagraphStyle: true,
        },

        /* whether to select text in each document */
        selectText: true,

        /* show results in an alert */
        showResults: true,

    };

    if (0 === app.documents.length)
        return alert('Please open some documents and try again.');

    var doc = app.activeDocument,
        text = doc.selection[0];

    if (!text)
        return alert('Please select some target text and try again.');

    if (
        'undefined' === typeof text.showText
        && text.hasOwnProperty('texts')
        && text.texts.length > 0
    )
        text = text.texts[0];

    if ('function' !== typeof text.showText)
        return alert('Please select some target text and try again.');

    var zoom = doc.layoutWindows[0].zoomPercentage;

    // configure the find
    app.findChangeTextOptions.properties = {
        caseSensitive: false,
        includeFootnotes: true,
        includeHiddenLayers: false,
        includeLockedLayersForFind: false,
        includeLockedStoriesForFind: false,
        includeMasterPages: true,
        wholeWord: false,
    };

    // reset
    app.findTextPreferences = NothingEnum.NOTHING;

    // the text to search
    app.findTextPreferences.findWhat = text.contents;

    // configure the find with settings
    for (var key in settings.findTextPreferences) {

        if (
            true === settings.findTextPreferences[key]
            && settings.findTextPreferences.hasOwnProperty(key)
            && app.findTextPreferences.hasOwnProperty(key)
            && text.hasOwnProperty(key)
        )
            app.findTextPreferences[key] = text[key];

    }

    // perform the find
    var found = app.findText(),
        foundDoc,
        already = {};

    for (var i = 0; i < found.length; i++) {

        // get the found text's document
        foundDoc = found[i];

        while ('Document' !== foundDoc.constructor.name)
            foundDoc = foundDoc.parent;

        if (
            0 === foundDoc.layoutWindows.length
            || already[foundDoc.name]
        )
            continue;

        already[foundDoc.name] = true;

        // show the text
        found[i].showText();
        app.activeDocument.layoutWindows[0].zoomPercentage = zoom;

        if (settings.selectText)
            found[i].select();

    }

    if (settings.showResults)
        alert('Found ' + found.length + ' texts.');

};
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Show Selected Text In All Documents');