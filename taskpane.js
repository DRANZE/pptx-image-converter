/* ================================================================
   Image → Editable Slide  |  PowerPoint Add-in
   Requires: Office.js (loaded in taskpane.html)
   Model:    Claude Vision via Anthropic API
   ================================================================ */

// ── Initialise Office.js ─────────────────────────────────────────
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById('convertBtn').disabled = false;
    loadSavedApiKey();
  }
});

// ── API-key helpers ──────────────────────────────────────────────
function loadSavedApiKey() {
  const k = localStorage.getItem('img2slide_api_key');
  if (k) {
    document.getElementById('apiKeyInput').value = k;
    document.getElementById('keySaved').style.display = 'block';
  }
}

function saveApiKey() {
  const k = document.getElementById('apiKeyInput').value.trim();
  if (!k) { showStatus('Please enter a valid API key.', 'error'); return; }
  localStorage.setItem('img2slide_api_key', k);
  document.getElementById('keySaved').style.display = 'block';
  showStatus('API key saved!', 'success');
}

function toggleSettings() {
  const p = document.getElementById('settingsPanel');
  p.style.display = p.style.display === 'none' ? 'block' : 'none';
}

// ── UI helpers ───────────────────────────────────────────────────
function showStatus(msg, type = 'info') {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className   = 'status ' + type;
  el.style.display = 'block';
}

function setProgress(visible, text = '') {
  document.getElementById('progress').style.display = visible ? 'block' : 'none';
  if (text) document.getElementById('progressText').textContent = text;
}

// ── Detect image MIME type from base64 prefix ────────────────────
function detectMime(b64) {
  if (b64.startsWith('/9j/'))          return 'image/jpeg';
  if (b64.startsWith('iVBORw0KGgo'))  return 'image/png';
  if (b64.startsWith('R0lGOD'))       return 'image/gif';
  if (b64.startsWith('UklGR'))        return 'image/webp';
  return 'image/png'; // safe default
}

// ── Step 1: Read selected image from the slide ───────────────────
async function getSelectedImageData() {
  let result = null;

  await PowerPoint.run(async ctx => {
    const sel = ctx.presentation.getSelectedShapes();
    sel.load('items');
    await ctx.sync();

    if (!sel.items.length) {
      throw new Error('No image selected.\n\nPlease click on a picture on your slide first, then press Convert.');
    }

    const shape = sel.items[0];
    shape.load('id,left,top,width,height,shapeType');
    await ctx.sync();

    let base64;
    try {
      const imgRef = shape.image.getBase64EncodedSrc();
      await ctx.sync();
      base64 = imgRef.value;
    } catch {
      throw new Error('The selected object is not a picture.\n\nSelect a photo or image (not a chart, table, or text box) and try again.');
    }

    result = {
      id:     shape.id,
      base64,
      left:   shape.left,
      top:    shape.top,
      width:  shape.width,
      height: shape.height,
    };
  });

  return result;
}

// ── Step 2: Send image to Claude and get JSON of elements ─────────
async function analyzeWithClaude(base64, mime) {
  const apiKey = localStorage.getItem('img2slide_api_key');
  if (!apiKey) {
    throw new Error('API key not set.\n\nClick "⚙️ API Key Settings" at the bottom and save your Anthropic API key first.');
  }

  const systemPrompt = `You are an expert at analyzing presentation slides and extracting every visual element.
Return ONLY a valid JSON object — no markdown fences, no explanation, nothing else.`;

  const userPrompt = `Analyze every visual element in this slide image.

Return ONLY this exact JSON structure (no markdown, no text before or after):

{
  "backgroundColor": "#FFFFFF",
  "elements": [
    {
      "type": "shape",
      "shape": "rectangle",
      "x": 0.0, "y": 0.0, "width": 1.0, "height": 0.14,
      "fillColor": "#1F3864",
      "borderColor": "none",
      "borderWidth": 0,
      "text": "",
      "textFontSize": 12,
      "textColor": "#FFFFFF",
      "textBold": false,
      "textAlignment": "center"
    },
    {
      "type": "text",
      "text": "Slide Title Here",
      "x": 0.05, "y": 0.03,
      "width": 0.9, "height": 0.08,
      "fontSize": 28,
      "fontFamily": "Calibri",
      "color": "#FFFFFF",
      "bold": true,
      "italic": false,
      "underline": false,
      "alignment": "center"
    },
    {
      "type": "line",
      "x1": 0.05, "y1": 0.16,
      "x2": 0.95, "y2": 0.16,
      "color": "#AAAAAA",
      "weight": 1.5
    }
  ]
}

RULES — read carefully:
1. ALL position/size values (x, y, width, height, x1, y1, x2, y2) are DECIMAL FRACTIONS of the image's total dimensions (0.0 = left/top edge, 1.0 = right/bottom edge). Never use pixel values.
2. Process elements back-to-front: background fills first, then mid-layer shapes, then text on top.
3. Extract EVERY visible text element with exact content, estimated font name, size in points, hex color, bold/italic.
4. Detect shapes: rectangle, roundedRectangle, ellipse, circle, triangle, diamond, pentagon, hexagon, arrow, rightArrow, leftArrow.
5. Detect ALL horizontal/vertical/diagonal lines, dividers, borders as type "line".
6. Estimate font sizes relative to the image height (a title that is ~6% of the image height ≈ 28–32pt on a standard slide).
7. Use precise hex colors (#RRGGBB). Match gradients to their dominant color.
8. Include text inside shapes in the shape's "text" field; also add a separate "text" element on top if the text needs independent formatting.
9. If the entire slide has a background color or image-fill, set "backgroundColor".
10. Return ONLY the raw JSON — no prose, no markdown code fences.`;

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-calls': 'true',
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 4096,
      system: systemPrompt,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mime, data: base64 } },
          { type: 'text', text: userPrompt },
        ],
      }],
    }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    const msg = err.error?.message || `API error ${res.status}`;
    if (res.status === 401) throw new Error('Invalid API key. Please check your key in Settings.');
    if (res.status === 429) throw new Error('Rate limit reached. Please wait a moment and try again.');
    throw new Error('API error: ' + msg);
  }

  const data = await res.json();
  const raw  = data.content.map(c => c.text || '').join('').trim();

  // Strip accidental markdown fences
  const cleaned   = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
  const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('AI returned an unexpected format. Please try again.');

  try {
    return JSON.parse(jsonMatch[0]);
  } catch {
    throw new Error('Could not parse AI response. Please try again.');
  }
}

// ── Step 3: Insert all extracted elements onto the slide ──────────
async function insertElements(analysis, img) {
  const { elements = [], backgroundColor } = analysis;
  const { left: ox, top: oy, width: ow, height: oh, id: origId } = img;

  // Map shape names → Office.js GeometricShapeType enum strings
  const GEO = {
    rectangle:        'Rectangle',
    roundedRectangle: 'RoundRectangle',
    ellipse:          'Ellipse',
    circle:           'Ellipse',
    triangle:         'IsoscelesTriangle',
    diamond:          'Diamond',
    pentagon:         'Pentagon',
    hexagon:          'Hexagon',
    arrow:            'RightArrow',
    rightArrow:       'RightArrow',
    leftArrow:        'LeftArrow',
  };

  const ALIGN = {
    center:  'Center',
    right:   'Right',
    left:    'Left',
    justify: 'Justify',
  };

  await PowerPoint.run(async ctx => {
    // Get active slide
    const selectedSlides = ctx.presentation.getSelectedSlides();
    selectedSlides.load('items');
    await ctx.sync();
    if (!selectedSlides.items.length) throw new Error('Could not access the current slide.');
    const slide = selectedSlides.items[0];

    // Optional: set slide background
    if (backgroundColor && backgroundColor !== 'none' && backgroundColor !== 'transparent') {
      try { slide.background.fill.setSolidColor(backgroundColor); } catch { /* non-critical */ }
    }

    // Insert each element
    for (const el of elements) {
      try {
        if (el.type === 'text') {
          _addTextBox(slide, el, ox, oy, ow, oh, ALIGN);

        } else if (el.type === 'shape') {
          _addShape(slide, el, ox, oy, ow, oh, GEO, ALIGN);

        } else if (el.type === 'line') {
          _addLine(slide, el, ox, oy, ow, oh);
        }
      } catch (e) {
        console.warn('Skipped element:', JSON.stringify(el), e);
      }
    }

    // Remove the original image
    try {
      const allShapes = slide.shapes;
      allShapes.load('items');
      await ctx.sync();
      for (const s of allShapes.items) s.load('id');
      await ctx.sync();
      for (const s of allShapes.items) {
        if (s.id === origId) { s.delete(); break; }
      }
    } catch (e) {
      console.warn('Could not remove original image:', e);
    }

    await ctx.sync();
  });
}

// ── Element insertion helpers ────────────────────────────────────

function _addTextBox(slide, el, ox, oy, ow, oh, ALIGN) {
  const left   = ox + el.x      * ow;
  const top    = oy + el.y      * oh;
  const width  = Math.max(el.width  * ow, 20);
  const height = Math.max(el.height * oh, 14);

  const tb = slide.shapes.addTextBox(el.text || '', { left, top, width, height });
  tb.fill.clear();
  tb.lineFormat.visible = false;

  const tf = tb.textFrame;
  tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;

  const rng = tf.textRange;
  rng.font.size      = el.fontSize   || 12;
  rng.font.bold      = el.bold       || false;
  rng.font.italic    = el.italic     || false;
  rng.font.underline = el.underline  || false;
  rng.font.color     = el.color      || '#000000';
  if (el.fontFamily)  rng.font.name  = el.fontFamily;

  rng.paragraphFormat.horizontalAlignment =
    PowerPoint.ParagraphHorizontalAlignment[ALIGN[el.alignment] || 'Left'];
}

function _addShape(slide, el, ox, oy, ow, oh, GEO, ALIGN) {
  const left   = ox + el.x      * ow;
  const top    = oy + el.y      * oh;
  const width  = Math.max(el.width  * ow, 4);
  const height = Math.max(el.height * oh, 4);

  const geoType = PowerPoint.GeometricShapeType[GEO[el.shape] || 'Rectangle'];
  const shape   = slide.shapes.addGeometricShape(geoType, { left, top, width, height });

  // Fill
  if (el.fillColor && el.fillColor !== 'none' && el.fillColor !== 'transparent') {
    shape.fill.setSolidColor(el.fillColor);
  } else {
    shape.fill.clear();
  }

  // Border
  if (el.borderColor && el.borderColor !== 'none' && el.borderColor !== 'transparent' && el.borderWidth > 0) {
    shape.lineFormat.visible = true;
    shape.lineFormat.color   = el.borderColor;
    shape.lineFormat.weight  = el.borderWidth || 1;
  } else {
    shape.lineFormat.visible = false;
  }

  // Text inside shape
  if (el.text) {
    shape.textFrame.textRange.text          = el.text;
    shape.textFrame.textRange.font.size     = el.textFontSize  || 12;
    shape.textFrame.textRange.font.bold     = el.textBold      || false;
    shape.textFrame.textRange.font.color    = el.textColor     || '#000000';
    shape.textFrame.textRange.paragraphFormat.horizontalAlignment =
      PowerPoint.ParagraphHorizontalAlignment[ALIGN[el.textAlignment] || 'Center'];
  }
}

function _addLine(slide, el, ox, oy, ow, oh) {
  const x1 = ox + el.x1 * ow;
  const y1 = oy + el.y1 * oh;
  const x2 = ox + el.x2 * ow;
  const y2 = oy + el.y2 * oh;

  if (Math.abs(x2 - x1) < 1 && Math.abs(y2 - y1) < 1) return; // skip zero-length

  const line = slide.shapes.addLine(x1, y1, x2, y2);
  line.lineFormat.color  = el.color  || '#000000';
  line.lineFormat.weight = el.weight || 1;
}

// ── Main orchestrator ─────────────────────────────────────────────
async function convertToEditable() {
  const btn = document.getElementById('convertBtn');
  btn.disabled = true;
  document.getElementById('status').style.display = 'none';

  try {
    // 1. Read image
    setProgress(true, 'Reading selected image from slide…');
    const imgData = await getSelectedImageData();

    // 2. Analyse
    setProgress(true, 'Analysing with AI — this takes 10–30 sec…');
    const mime     = detectMime(imgData.base64);
    const analysis = await analyzeWithClaude(imgData.base64, mime);

    if (!analysis.elements || analysis.elements.length === 0) {
      throw new Error('No elements were detected in this image. Try a clearer, higher-resolution image.');
    }

    // 3. Insert
    setProgress(true, `Inserting ${analysis.elements.length} elements onto slide…`);
    await insertElements(analysis, imgData);

    setProgress(false);
    showStatus(
      `✅ Done! ${analysis.elements.length} elements placed on your slide. The original image has been removed. All elements are now fully editable.`,
      'success'
    );

  } catch (err) {
    setProgress(false);
    showStatus('❌ ' + (err.message || 'Unexpected error. Please try again.'), 'error');
  } finally {
    btn.disabled = false;
  }
}
