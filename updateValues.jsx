/* --------------------------------------
Update Indents and Spaces
by Aaron Troia (@atroia)
Modified Date: 7/22/25

Description: 
Update InDesign paragraph indents and spaces 
based on font size and leading
-------------------------------------- */


var d = app.activeDocument;
var paraStyles = d.allParagraphStyles;

main();

function main() {
    if (app.documents.length === 0) {
        alert("No active document open.");
        return;
    } else {
        indents();
        spaceBefore();
    }
};


function indents() {
    var numberPattern = /\d+/;

    for (var i = 0; i < paraStyles.length; i++) {
        var style = paraStyles[i];
        var name = style.name;
        if (name === "[No Paragraph Style]") continue;

        try {
            var fontSize = style.pointSize;
            if (fontSize === NothingEnum.NOTHING) continue;

            var leftIndent = style.leftIndent;
            var rightIndent = style.rightIndent;
            var firstLineIndent = style.firstLineIndent;

            var numberMatch = name.match(numberPattern);
            var multiplier = numberMatch ? parseInt(numberMatch[0], 10) : null;

            // Left Indent logic
            if (multiplier) {
                style.leftIndent = fontSize * multiplier;
            } else if (leftIndent !== 0 && leftIndent !== fontSize) {
                style.leftIndent = fontSize;
            }

            // Right Indent logic
            if (rightIndent !== 0 && rightIndent !== fontSize) {
                style.rightIndent = fontSize;
            }

            // First Line Indent logic (negative -> negative fontSize only)
            if (firstLineIndent !== 0) {
                if (firstLineIndent < 0) {
                    style.firstLineIndent = -fontSize;
                } else if (firstLineIndent !== fontSize) {
                    style.firstLineIndent = fontSize;
                }
            }

            $.writeln("Updated: " + name);

        } catch (e) {
            $.writeln("Error with style: " + name + " - " + e);
        }
    }

    // alert("Indent updates complete based on font size and style number.");
};


function spaceBefore() {
    // List of exclusion words (case-insensitive)
    var excludeWords = ['footnote', 'endnote', 'caption', 'table'];

    for (var i = 0; i < paraStyles.length; i++) {
        var style = paraStyles[i];
        var styleName = style.name;

        // Skip [No Paragraph Style] and 'Basic Paragraph Small'
        if (styleName == "[No Paragraph Style]" || styleName == "Basic Paragraph Small") continue;

        // Skip if style name contains any excluded words (case-insensitive)
        var lowerName = styleName.toLowerCase();
        var skip = false;
        for (var j = 0; j < excludeWords.length; j++) {
            if (lowerName.indexOf(excludeWords[j]) !== -1) {
                skip = true;
                break;
            }
        }
        if (skip) continue;

        try {
            var leadingValue = style.leading;
            if (leadingValue === NothingEnum.NOTHING) {
                $.writeln("Skipping (no leading set): " + styleName);
                continue;
            }

            // Special case for 'break' style
            if (styleName.toLowerCase() === "break") {
                var leading = Number(leadingValue);
                style.spaceBefore = leading;
                style.spaceAfter = leading;
                $.writeln("Updated 'break' style: spaceBefore and spaceAfter = " + leading);
                continue;
            }

            if (style.spaceBefore !== 0) {
                style.spaceBefore = leadingValue;
                $.writeln("Updated: " + styleName + " | spaceBefore = " + leadingValue);
            }
        } catch (e) {
            $.writeln("Error processing " + styleName + ": " + e);
        }
    }

    // alert("Completed updating spaceBefore and break styles to leading where applicable.");
};
