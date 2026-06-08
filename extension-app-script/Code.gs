// Adds a custom menu to the Google Slides UI when opened
function onOpen() {
  SlidesApp.getUi()
      .createMenu('SVG Importer')
      .addItem('Open Importer', 'showSidebar')
      .addToUi();
}

// Opens the sidebar HTML panel
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Import SVG Data')
      .setWidth(300);
  SlidesApp.getUi().showSidebar(html);
}


// NEW: Resolves global variables, expands short hex, and converts rgb() to hex
function resolveColor(color, templateVars) {
  if (!color) return color;
  
  const trimmedColor = color.trim();
  let finalColor = trimmedColor;
  
  // 1. Check for 'global-color-X' override or exact variable match
  if (templateVars['global-color-' + trimmedColor] !== undefined) {
    finalColor = templateVars['global-color-' + trimmedColor];
  } else if (templateVars[trimmedColor] !== undefined) {
    finalColor = templateVars[trimmedColor];
  }
  
  // 2. Convert rgb(r,g,b) format to Hex format
  const rgbMatch = finalColor.match(/^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1]).toString(16).padStart(2, '0');
    const g = parseInt(rgbMatch[2]).toString(16).padStart(2, '0');
    const b = parseInt(rgbMatch[3]).toString(16).padStart(2, '0');
    finalColor = `#${r}${g}${b}`;
  }
  
  // 3. Expand short hex (e.g. #000 to #000000)
  if (/^#[0-9a-fA-F]{3}$/.test(finalColor)) {
    return finalColor.replace(/^#([0-9a-fA-F])([0-9a-fA-F])([0-9a-fA-F])$/, '#$1$1$2$2$3$3');
  }
  
  return finalColor;
}

//Parses <style> tags (both direct children and inside <defs>)
function parseCssStyles(root, ns) {
  const styles = {};
  
  // Find direct <style> tags
  let styleNodes = root.getChildren('style', ns);
  
  // Find <style> tags nested inside <defs>
  const defs = root.getChild('defs', ns);
  if (defs) {
    styleNodes = styleNodes.concat(defs.getChildren('style', ns));
  }
  
  styleNodes.forEach(styleNode => {
    const cssText = styleNode.getValue();
    const ruleRegex = /([^{]+)\{\s*([^}]+)\s*\}/g; 
    let match;
    
    while ((match = ruleRegex.exec(cssText)) !== null) {
      const selectors = match[1].split(',').map(s => s.trim());
      const properties = match[2].split(';').reduce((acc, prop) => {
        const parts = prop.split(':');
        if (parts.length === 2) acc[parts[0].trim()] = parts[1].trim();
        return acc;
      }, {});
      
      selectors.forEach(selector => {
        styles[selector] = Object.assign(styles[selector] || {}, properties);
      });
    }
  });
  return styles;
}

// Resolves cascading styles according to your fallback rules
function resolveStyles(elementType, classAttr, cssStyles) {
  // 1. Hardcoded Fallbacks
  const resolved = elementType === 'text' 
    ? { 'font-family': 'Red Hat Text', 'font-size': '18', 'fill': '#000000', 'width': '200', 'height': '50', 'text-align': 'left' }
    : elementType === 'circle'
    ? { 'fill': '#cccccc', 'stroke': '#000000', 'r': '50' }
    : { 'fill': '#cccccc', 'stroke': '#000000', 'width': '100', 'height': '100', 'rx': '0', 'ry': '0' };

  const merge = (selector) => {
    if (cssStyles[selector]) Object.assign(resolved, cssStyles[selector]);
  };

  // 2. Fallback to 'svg'
  merge('svg');
  // 3. Fallback to element type ('text', '.text', 'rect')
  merge(elementType); 
  merge('.' + elementType); 
  // 4. Specific class mentioned
  if (classAttr) {
    classAttr.split(' ').forEach(cls => merge('.' + cls));
  }
  
  return resolved;
}


//Recursively finds elements, calculating offsets and enforcing conditional visibility
function getElementsWithOffsets(elem, ns, templateVars, offsetX = 0, offsetY = 0, results = []) {
  const name = elem.getName();
  
  // NEW: Check conditional display. If the variable is empty or doesn't exist, halt and skip this node & its children
  const displayIfVar = elem.getAttribute('data-var-displayif')?.getValue();
  if (displayIfVar !== undefined) {
    const conditionValue = templateVars[displayIfVar.trim()];
    if (!conditionValue || conditionValue.trim() === "") {
      return results; // Abort processing this element branch immediately
    }
  }

  let dx = offsetX;
  let dy = offsetY;
  
  const getDynamicShift = (attrName) => {
    const varName = elem.getAttribute(attrName)?.getValue();
    return (varName && templateVars[varName]) ? (parseFloat(templateVars[varName]) || 0) : 0;
  };

  dx += getDynamicShift('data-var-x-add');
  dy += getDynamicShift('data-var-y-add');
  
  if (name === 'g') {
    dx += parseFloat(elem.getAttribute('x')?.getValue() || 0);
    dy += parseFloat(elem.getAttribute('y')?.getValue() || 0);
    
    const transform = elem.getAttribute('transform')?.getValue() || "";
    const translateMatch = transform.match(/translate\(([^,\s]+)[\s,]+([^)]+)\)/);
    if (translateMatch) {
      dx += parseFloat(translateMatch[1]);
      dy += parseFloat(translateMatch[2]);
    }
  }
  
  if (['rect', 'circle', 'text', 'path', 'polygon', 'polyline'].includes(name)) {
    results.push({ element: elem, type: name, offsetX: dx, offsetY: dy });
  }
  
  elem.getChildren().forEach(child => {
    getElementsWithOffsets(child, ns, templateVars, dx, dy, results);
  });
  
  return results;
}


// Main function to parse SVG and inject shapes into the active slide
function importSvgToSlide(svgString, varString = "") {
  const slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
  
  try {
    const templateVars = {};
    if (varString) {
      varString.split('\n').forEach(line => {
        const colonIndex = line.indexOf(':');
        if (colonIndex > -1) {
          const key = line.substring(0, colonIndex).trim();
          const val = line.substring(colonIndex + 1).trim();
          templateVars[key] = val;
        }
      });
    }

    const document = XmlService.parse(svgString);
    const root = document.getRootElement();
    const ns = XmlService.getNamespace('http://www.w3.org/2000/svg');
    const cssStyles = parseCssStyles(root, ns);
    

    const pageElements = []; //Collector for grouping
    // Process ALL elements in the exact order they appear in the SVG (preserves layering)
    const elements = getElementsWithOffsets(root, ns, templateVars);
    
    elements.forEach(item => {
      const { element, type, offsetX, offsetY } = item;
      const classAttr = element.getAttribute('class')?.getValue();
      
      // Load standard CSS / fallbacks first
      const styles = resolveStyles(type, classAttr, cssStyles);
      
      // NEW: Parse inline style="..." attribute and overwrite CSS defaults
      const inlineStyleAttr = element.getAttribute('style')?.getValue();
      if (inlineStyleAttr) {
        inlineStyleAttr.split(';').forEach(rule => {
          const parts = rule.split(':');
          if (parts.length === 2) {
            styles[parts[0].trim().toLowerCase()] = parts[1].trim();
          }
        });
      }
      
      //Calculate colors using the global resolver
      const varFill = element.getAttribute('data-var-fill')?.getValue();
      const baseFill = (varFill && templateVars[varFill]) ? templateVars[varFill] : (element.getAttribute('fill')?.getValue() || styles['fill']);
      const finalFill = resolveColor(baseFill, templateVars);
      
      const varStroke = element.getAttribute('data-var-stroke')?.getValue();
      const baseStroke = (varStroke && templateVars[varStroke]) ? templateVars[varStroke] : (element.getAttribute('stroke')?.getValue() || styles['stroke']);
      const finalStroke = resolveColor(baseStroke, templateVars);

      const rawStrokeWidth = element.getAttribute('stroke-width')?.getValue() || styles['stroke-width'];
      const finalStrokeWidth = rawStrokeWidth ? parseFloat(rawStrokeWidth) : null;

      const varWidth = element.getAttribute('data-var-width')?.getValue();
      const finalWidth = (varWidth && templateVars[varWidth]) ? templateVars[varWidth] : (element.getAttribute('width')?.getValue() || styles['width']);
      
      const varHeight = element.getAttribute('data-var-height')?.getValue();
      const finalHeight = (varHeight && templateVars[varHeight]) ? templateVars[varHeight] : (element.getAttribute('height')?.getValue() || styles['height']);

    if (type === 'rect') {
        const x = parseFloat(element.getAttribute('x')?.getValue() || 0) + offsetX;
        const y = parseFloat(element.getAttribute('y')?.getValue() || 0) + offsetY;
        const width = parseFloat(finalWidth);
        const height = parseFloat(finalHeight);
        
        //Grab rounding values from attributes or CSS
        const rx = parseFloat(element.getAttribute('rx')?.getValue() || styles['rx']);
        const ry = parseFloat(element.getAttribute('ry')?.getValue() || styles['ry']);
        
        if (width <= 0 || height <= 0 || isNaN(width) || isNaN(height)) return; 
        
        // NEW: If radius is greater than 0, use Google's native Rounded Rectangle enum
        const shapeEnum = (rx > 0 || ry > 0) ? SlidesApp.ShapeType.ROUND_RECTANGLE : SlidesApp.ShapeType.RECTANGLE;
        const shape = slide.insertShape(shapeEnum, x, y, width, height);
        
        const fill = finalFill;
        if (fill && fill.startsWith('#')) shape.getFill().setSolidFill(fill); 
        
        if (finalStroke && finalStroke.startsWith('#')) {
          const border = shape.getBorder();
          border.getLineFill().setSolidFill(finalStroke);
          if (finalStrokeWidth !== null) border.setWeight(finalStrokeWidth);
        }
        else if (stroke === 'none') shape.getBorder().setTransparent();
        
        pageElements.push(shape);
      }
      else if (type === 'circle') {
        const r = parseFloat(element.getAttribute('r')?.getValue() || styles['r']);
        
        // Prevent GAS crashes if geometry is invalid
        if (r <= 0 || isNaN(r)) return; 
        
        const cx = parseFloat(element.getAttribute('cx')?.getValue() || 0) + offsetX;
        const cy = parseFloat(element.getAttribute('cy')?.getValue() || 0) + offsetY;
        
        const x = cx - r;
        const y = cy - r;
        const diameter = r * 2;
        
        // FIX: The correct Apps Script Enum is ELLIPSE, not OVAL
        const shape = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, diameter, diameter);
        
        const fill = finalFill;
        if (fill && fill.startsWith('#')) shape.getFill().setSolidFill(fill); 
        
        if (finalStroke && finalStroke.startsWith('#')) {
          const border = shape.getBorder();
          border.getLineFill().setSolidFill(finalStroke);
          if (finalStrokeWidth !== null) border.setWeight(finalStrokeWidth);
        }
        else if (stroke === 'none') shape.getBorder().setTransparent();
        
        pageElements.push(shape);
      } 
      else if (type === 'text') {
        const x = parseFloat(element.getAttribute('x')?.getValue() || 0) + offsetX;
        const y = parseFloat(element.getAttribute('y')?.getValue() || 0) + offsetY;
        const boxWidth = parseFloat(finalWidth); 
        const boxHeight = parseFloat(finalHeight);
        
        if (boxWidth <= 0 || boxHeight <= 0 || isNaN(boxWidth) || isNaN(boxHeight)) return;

        let content = element.getValue();
        content = content.replace(/\{([^}]+)\}/g, (match, varName) => {
          const trimmedName = varName.trim();
          return templateVars[trimmedName] !== undefined ? templateVars[trimmedName] : "";
        });
        
        if (content.trim() !== "") {
          const explicitAnchor = element.getAttribute('text-anchor')?.getValue();
          const explicitBaseline = element.getAttribute('dominant-baseline')?.getValue();
          
          const align = (explicitAnchor || styles['text-align'] || "").toLowerCase();
          const baseline = (explicitBaseline || "").toLowerCase();

          let startX = x;
          let startY = y;

          // ==========================================
          // DUAL-MODE GEOMETRY
          // ==========================================
          // 1. Strict SVG Mode: Shift the text backward if explicit anchor attributes exist (Pikchr)
          if (explicitAnchor) {
            const a = explicitAnchor.toLowerCase();
            if (a === 'middle' || a === 'center') startX -= (boxWidth / 2);
            else if (a === 'end' || a === 'right') startX -= boxWidth;
          }
          if (explicitBaseline) {
            const b = explicitBaseline.toLowerCase();
            if (b === 'central' || b === 'middle') startY -= (boxHeight / 2);
            else startY -= (boxHeight * 0.8); // Native SVG Y baseline sits near the bottom
          }
          // 2. CSS/HTML Mode: If no explicit SVG anchors, X/Y act as standard top-left bounding box origins

          const PADDING_COMPENSATION = 7;
          const adjustedX = startX - PADDING_COMPENSATION;
          const adjustedY = startY - PADDING_COMPENSATION;
          const adjustedWidth = boxWidth + (PADDING_COMPENSATION * 2);
          const adjustedHeight = boxHeight + (PADDING_COMPENSATION * 2);

          const textBox = slide.insertTextBox(content, adjustedX, adjustedY, adjustedWidth, adjustedHeight);
          
          // Contextual Vertical Alignment
          if (baseline === 'central' || baseline === 'middle') {
            textBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } else if (explicitBaseline) {
            textBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
          } else {
            textBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); // Default for Bounding Box mode
          }

          const textRange = textBox.getText();
          const textStyle = textRange.getTextStyle();
          
          let gasAlign = SlidesApp.ParagraphAlignment.START;
          if (align === 'middle' || align === 'center') gasAlign = SlidesApp.ParagraphAlignment.CENTER;
          else if (align === 'end' || align === 'right') gasAlign = SlidesApp.ParagraphAlignment.END;
          textRange.getParagraphStyle().setParagraphAlignment(gasAlign);

          if (styles['font-family']) textStyle.setFontFamily(styles['font-family'].replace(/['"]/g, '')); 
          
          const rawFontSize = element.getAttribute('font-size')?.getValue() || styles['font-size'];
          if (rawFontSize) {
            let fSize = parseFloat(rawFontSize);
            if (rawFontSize.includes('%')) fSize = (fSize / 100) * 18; 
            if (fSize > 0) textStyle.setFontSize(fSize);
          }

          if (finalFill && finalFill.startsWith('#')) textStyle.setForegroundColor(finalFill);
          
          pageElements.push(textBox); 
        }
      }
      else if (type === 'path') {
        const dAttr = element.getAttribute('d')?.getValue();
        if (!dAttr) return;

        // Tokenize the SVG path string (separates letters from numbers safely)
        const tokens = dAttr.match(/[a-df-zA-DF-Z]|[+-]?\d*\.?\d+(?:[eE][+-]?\d+)?/g) || [];
        
        let startX = 0, startY = 0;
        let currentX = 0, currentY = 0;
        let i = 0;
        
        const pathSegments = []; // Temporary collector for raw line vectors

        while (i < tokens.length) {
          let cmd = tokens[i];
          
          if (isNaN(cmd)) {
            i++;
          } else {
            // If no explicit command token is provided, reuse the previous implicit action
            cmd = (cmd === tokens[i-1]) ? cmd : "L"; 
          }

          if (cmd === 'M' || cmd === 'm') {
            let nx = parseFloat(tokens[i++]);
            let ny = parseFloat(tokens[i++]);
            if (cmd === 'm') { nx += currentX; ny += currentY; }
            currentX = nx; currentY = ny;
            startX = currentX; startY = currentY; // Set anchor to close path later
          } 
          else if (cmd === 'L' || cmd === 'l') {
            let nx = parseFloat(tokens[i++]);
            let ny = parseFloat(tokens[i++]);
            if (cmd === 'l') { nx += currentX; ny += currentY; }
            pathSegments.push({ x1: currentX, y1: currentY, x2: nx, y2: ny });
            currentX = nx; currentY = ny;
          } 
          else if (cmd === 'H' || cmd === 'h') {
            let nx = parseFloat(tokens[i++]);
            if (cmd === 'h') nx += currentX;
            pathSegments.push({ x1: currentX, y1: currentY, x2: nx, y2: currentY });
            currentX = nx;
          } 
          else if (cmd === 'V' || cmd === 'v') {
            let ny = parseFloat(tokens[i++]);
            if (cmd === 'v') ny += currentY;
            pathSegments.push({ x1: currentX, y1: currentY, x2: currentX, y2: ny });
            currentY = ny;
          }
          else if (cmd === 'Z' || cmd === 'z') {
            // Close path command: Draw a line back to the original starting anchor
            pathSegments.push({ x1: currentX, y1: currentY, x2: startX, y2: startY });
            currentX = startX; currentY = startY;
          }
          else {
            // FIX: Skip all parameter numbers to find the next command letter
            let nextCmdIndex = i;
            while (nextCmdIndex < tokens.length && !isNaN(tokens[nextCmdIndex])) {
              nextCmdIndex++;
            }
            
            // Draw a straight line to the final X/Y coordinates of the curve command
            if (nextCmdIndex - 2 >= i) {
              let nx = parseFloat(tokens[nextCmdIndex - 2]);
              let ny = parseFloat(tokens[nextCmdIndex - 1]);
              
              // If the curve command is lowercase (relative), add the current coordinates
              if (cmd === cmd.toLowerCase()) {
                nx += currentX;
                ny += currentY;
              }
              
              pathSegments.push({ x1: currentX, y1: currentY, x2: nx, y2: ny });
              currentX = nx; currentY = ny;
            }
            
            i = nextCmdIndex; // Jump safely to the next command letter
          }
        }


        // Filter out zero-length artifacts before analyzing
        const cleanSegments = pathSegments.filter(seg => !(seg.x1 === seg.x2 && seg.y1 === seg.y2));

        // ==========================================
        // BOX DETECTION ENGINE
        // ==========================================
        let isBox = false;
        let boxX = 0, boxY = 0, boxW = 0, boxH = 0;

        if (cleanSegments.length === 4) {
          // Gather all distinct coordinate bounds
          const xCoords = cleanSegments.map(s => s.x1).concat(cleanSegments.map(s => s.x2));
          const yCoords = cleanSegments.map(s => s.y1).concat(cleanSegments.map(s => s.y2));
          
          const minX = Math.min(...xCoords);
          const maxX = Math.max(...xCoords);
          const minY = Math.min(...yCoords);
          const maxY = Math.max(...yCoords);
          
          // Verify that every single segment falls precisely along the bounding box borders
          const validSegments = cleanSegments.every(seg => {
            const isHoriz = (seg.y1 === seg.y2 && (seg.y1 === minY || seg.y1 === maxY));
            const isVert = (seg.x1 === seg.x2 && (seg.x1 === minX || seg.x1 === maxX));
            return isHoriz || isVert;
          });

          if (validSegments && (maxX > minX) && (maxY > minY)) {
            isBox = true;
            boxX = minX + offsetX;
            boxY = minY + offsetY;
            boxW = maxX - minX;
            boxH = maxY - minY;
          }
        }

        // ==========================================
        // GENERATION DESPATCH
        // ==========================================
        if (isBox) {
          // Check if any rx/ry rounded context applies (using our native rounding strategy)
          const rx = parseFloat(element.getAttribute('rx')?.getValue() || styles['rx'] || 0);
          const ry = parseFloat(element.getAttribute('ry')?.getValue() || styles['ry'] || 0);
          const shapeEnum = (rx > 0 || ry > 0 || dAttr.includes('A')) ? SlidesApp.ShapeType.ROUND_RECTANGLE : SlidesApp.ShapeType.RECTANGLE;
          
          const shape = slide.insertShape(shapeEnum, boxX, boxY, boxW, boxH);
          pageElements.push(shape);

          if (finalFill && finalFill.startsWith('#')) shape.getFill().setSolidFill(finalFill);
          else if (finalFill === 'none') shape.getFill().setTransparent();
          
          if (finalStroke && finalStroke.startsWith('#')) {
            const border = shape.getBorder();
            border.getLineFill().setSolidFill(finalStroke);
            if (finalStrokeWidth !== null) border.setWeight(finalStrokeWidth);
          }
          else if (finalStroke === 'none') shape.getBorder().setTransparent();
        } 
        else {
          // Generate the lines onto the Google Slide layout canvas
          cleanSegments.forEach(seg => {
            // NEW: Skip zero-length lines to prevent Google Slides API crashes
            if (seg.x1 === seg.x2 && seg.y1 === seg.y2) return; 

            const x1 = seg.x1 + offsetX;
            const y1 = seg.y1 + offsetY;
            const x2 = seg.x2 + offsetX;
            const y2 = seg.y2 + offsetY;
            
            const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x1, y1, x2, y2);
            pageElements.push(line); 

            if (finalStroke && finalStroke.startsWith('#')) {
              line.getLineFill().setSolidFill(finalStroke);
              if (finalStrokeWidth !== null) line.setWeight(finalStrokeWidth);
            } else {
              line.setTransparent();
            }
          
          });
        }
      }
      else if (type === 'polygon' || type === 'polyline') {
        const points = element.getAttribute('points')?.getValue() || "";
        const coords = points.match(/-?\d*\.?\d+/g) || [];
        
        if (coords.length >= 4) {
          const startX = parseFloat(coords[0]);
          const startY = parseFloat(coords[1]);
          let currentX = startX;
          let currentY = startY;
          
          const pathSegments = [];
          for (let j = 2; j < coords.length; j += 2) {
            let nx = parseFloat(coords[j]);
            let ny = parseFloat(coords[j+1]);
            pathSegments.push({ x1: currentX, y1: currentY, x2: nx, y2: ny });
            currentX = nx;
            currentY = ny;
          }
          
          if (type === 'polygon') {
            pathSegments.push({ x1: currentX, y1: currentY, x2: startX, y2: startY });
          }
          
        // Generate the lines onto the Google Slide layout canvas
        pathSegments.forEach(seg => {
          // NEW: Skip zero-length lines to prevent Google Slides API crashes
          if (seg.x1 === seg.x2 && seg.y1 === seg.y2) return; 

          const x1 = seg.x1 + offsetX;
          const y1 = seg.y1 + offsetY;
          const x2 = seg.x2 + offsetX;
          const y2 = seg.y2 + offsetY;
          
          const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x1, y1, x2, y2);
          pageElements.push(line); 

          if (finalStroke && finalStroke.startsWith('#')) {
            line.getLineFill().setSolidFill(finalStroke);
            if (finalStrokeWidth !== null) line.setWeight(finalStrokeWidth);
          } else {
            line.setTransparent();
          }
        });
        }
      }




    });
    // APPEND BASE64 ENCODED VARIABLES TO SPEAKER NOTES
    if (varString && varString.trim() !== "") {
      // 2. Compute the Base64 representation using Google's native Utilities engine
      const base64EncodedVars = Utilities.base64Encode(varString, Utilities.Charset.UTF_8);

      // 3. Fetch the specific speaker notes shape, prepare the string, and append
      const notesShape = slide.getNotesPage().getSpeakerNotesShape();
      
      if (notesShape) {
        const textRange = notesShape.getText();
        const currentNotes = textRange.asString();
        
        // If the notes are already populated, append on a fresh line; otherwise, insert directly
        const combinedNotes = currentNotes.trim() === "" 
          ? base64EncodedVars 
          : currentNotes.replace(/\n$/, "") + "\n" + base64EncodedVars;

        textRange.setText(combinedNotes);
      }
    }
    if (pageElements.length > 1) {
      slide.group(pageElements);
    }
    return "Success! Imported shapes and text.";
  } catch (error) {
    return "Error parsing SVG: " + error.toString();
  }
}
