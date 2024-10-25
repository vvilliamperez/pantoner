function generatePantoneSwatchArray() {
    if (app.documents.length == 0) {
        alert("Please open a document and select an object.");
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    if (selection.length == 0) {
        alert("Please select at least one object.");
        return;
    }

    // Get unique colors from selected objects
    var colors = [];
    var colorSet = {};
    extractColorsFromSelection(selection, colors, colorSet);

    if (colors.length == 0) {
        alert("No colors found in the selected object.");
        return;
    }

    // Create a new layer for the swatches
    var swatchLayer = doc.layers.add();
    swatchLayer.name = "Pantone Swatches";

    // Define swatch size, stroke width, and spacing
    var swatchSize = 50;
    var strokeWidth = 1;
    var padding = 30;
    var x = 0, y = 0;

    for (var i = 0; i < colors.length; i++) {
        // Draw the swatch square
        var swatchRect = swatchLayer.pathItems.rectangle(y, x, swatchSize, swatchSize);
        swatchRect.fillColor = colors[i];
        swatchRect.stroked = true;
        swatchRect.strokeColor = doc.swatches.getByName("Black").color;
        swatchRect.strokeWidth = strokeWidth;

        // Get Pantone name or color details
        var colorName = getPantoneName(colors[i]);
        if (!colorName) {
            colorName = getColorFromSwatches(colors[i]);
            if (!colorName) {
                colorName = "Custom Color";
            }
        } 
        
        if (colorName.toLowerCase().indexOf("pantone") === 0) {
            colorName = colorName.slice(7); // Remove "Pantone" (7 characters)
        }
        // Remove any extra leading or trailing spaces
        colorName = colorName.replace(/^\s+|\s+$/g, "");

        // Add editable text below the swatch
        var swatchText = swatchLayer.textFrames.add();
        swatchText.contents = colorName;
        swatchText.position = [x, y - swatchSize - padding / 2];
        swatchText.textRange.characterAttributes.size = 12;
        swatchText.textRange.characterAttributes.fillColor = doc.swatches.getByName("Black").color;
        swatchText.textRange.characterAttributes.textFont = app.textFonts.getByName("Helvetica");

        // Calculate the text width and adjust its position to be centered
        var textWidth = swatchText.width;
        var textX = x + (swatchSize / 2) - (textWidth / 2);
        swatchText.position = [textX, y - swatchSize - 5];


        // Move to next swatch position
        x += swatchSize + padding;
        if (x > doc.width - swatchSize) {
            x = 0;
            y -= swatchSize + padding;
        }
    }
}

// Recursively extract colors from selected objects, including groups and compound paths
function extractColorsFromSelection(items, colors, colorSet) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        
        if (item.typename === "GroupItem") {
            extractColorsFromSelection(item.pageItems, colors, colorSet);
        } else if (item.typename === "CompoundPathItem") {
            extractColorsFromSelection(item.pathItems, colors, colorSet);
        } else if (item.filled) {
            var color = item.fillColor;
            var colorKey = getColorKey(color);

            if (!colorSet[colorKey]) {
                colors.push(color);
                colorSet[colorKey] = true;
            }
        }
    }
}

// Helper function to create a unique key for each color
function getColorKey(color) {
    if (color.typename === "SpotColor") {
        return color.spot.name;
    } else if (color.typename === "RGBColor") {
        return color.red + "-" + color.green + "-" + color.blue;
    } else if (color.typename === "CMYKColor") {
        return color.cyan + "-" + color.magenta + "-" + color.yellow + "-" + color.black;
    }
    return "";
}

// Helper function to get Pantone name from SpotColor
function getPantoneName(color) {
    if (color.typename === "SpotColor" && color.spot.colorType === ColorModel.SPOT) {
        return color.spot.name;
    }
    return null;
}

// Additional helper to match colors against existing swatches
function getColorFromSwatches(color) {
    var doc = app.activeDocument;
    var swatches = doc.swatches;

    for (var i = 0; i < swatches.length; i++) {
        var swatch = swatches[i];
        if (swatch.color.typename === color.typename) {
            if (color.typename === "SpotColor" && swatch.color.spot === color.spot) {
                return swatch.name;
            }
            if (color.typename === "CMYKColor" && compareCMYK(swatch.color, color)) {
                return swatch.name;
            }
            if (color.typename === "RGBColor" && compareRGB(swatch.color, color)) {
                return swatch.name;
            }
        }
    }
    return null;
}

// Compare two RGB colors
function compareRGB(color1, color2) {
    return (color1.red === color2.red && color1.green === color2.green && color1.blue === color2.blue);
}

// Compare two CMYK colors
function compareCMYK(color1, color2) {
    return (color1.cyan === color2.cyan && color1.magenta === color2.magenta &&
            color1.yellow === color2.yellow && color1.black === color2.black);
}

// Run the script
generatePantoneSwatchArray();
