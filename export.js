/**
 * HawkSpot · export.js
 * Professional PowerPoint export engine using PptxGenJS.
 *
 * Slide dimensions: 13.33" × 7.5" (standard 16:9)
 * Supports 1–6 images per slide with intelligent grid layout.
 * Auto overflow: creates continuation slides for >6 photos per shop.
 */

// ─── Layout Constants ─────────────────────────────────────────────────────────
const SLIDE_W   = 13.33;
const SLIDE_H   = 7.5;
const HEADER_H  = 1.1;
const GAP       = 0.1;
const PAD       = 0.12;
const CONTENT_Y = HEADER_H + PAD;
const CONTENT_H = SLIDE_H - CONTENT_Y - PAD;
const CONTENT_W = SLIDE_W - PAD * 2;
const MAX_SLIDE = 6;

// Brand palette (hex strings, no #)
const C = {
    navy:  '0B1120', navy2: '111827', navy3: '1E2A3A', navy4: '243447',
    blue:  '3B82F6', gray:  '94A3B8', gray2: '64748B', gray3: '4B6078',
    white: 'FFFFFF', bg:    'F8FAFC', border:'DDDDDD',
};

// Grid map: image count → { cols, rows }
const GRIDS = { 1:{c:1,r:1}, 2:{c:2,r:1}, 3:{c:3,r:1}, 4:{c:2,r:2}, 5:{c:3,r:2}, 6:{c:3,r:2} };

// ─── Main Export Function ─────────────────────────────────────────────────────
async function exportToPPT() {
    const selected = STATE.submissions.filter(s => STATE.selectedIds.has(s.id));

    if (!selected.length) { showToast('Select at least one shop to export.', 'error'); return; }
    if (typeof PptxGenJS === 'undefined') { showToast('PptxGenJS not loaded.', 'error'); return; }

    const btn = document.getElementById('export-btn');
    if (btn) { btn.textContent = '⟳ Building…'; btn.disabled = true; }

    try {
        const pptx = new PptxGenJS();
        pptx.author  = 'HawkSpot Field Intelligence';
        pptx.company = getPref('ppt_division', 'Bunnys Snacks Division');
        pptx.subject = 'Field Execution Evidence Report';
        pptx.title   = getPref('ppt_title', 'STANDS & POSM DEPLOYMENT');
        pptx.defineLayout({ name: 'W169', width: SLIDE_W, height: SLIDE_H });
        pptx.layout  = 'W169';

        // Build slides
        buildCoverSlide(pptx, selected.length);
        buildTocSlide(pptx, selected);

        for (let i = 0; i < selected.length; i++) {
            const sub    = selected[i];
            const photos = STATE.photosMap[sub.id] || [];
            if (!photos.length) continue;
            const chunks = chunkArr(photos, MAX_SLIDE);
            chunks.forEach((chunk, ci) => buildEvidenceSlide(pptx, sub, chunk, i + 1, selected.length, ci + 1, chunks.length));
        }

        buildSummarySlide(pptx, selected);

        // Write file
        const div  = getPref('ppt_division', 'BunnysSnacks').replace(/\s+/g, '_');
        const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        await pptx.writeFile({ fileName: `HawkSpot_${div}_${date}.pptx` });

        showToast(`✓ Exported — ${selected.length} shops · ${countPhotos(selected)} images`, 'success');

    } catch (err) {
        console.error('[HawkSpot] Export error:', err);
        showToast(`Export failed: ${err.message}`, 'error');
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Export PPTX <span class="export-count" id="export-count">${STATE.selectedIds.size}</span>`;
        }
    }
}

// ─── Cover Slide ──────────────────────────────────────────────────────────────
function buildCoverSlide(pptx, shopCount) {
    const sl  = pptx.addSlide();
    const SHT = pptx.ShapeType;
    sl.background = { fill: C.navy };

    const division = getPref('ppt_division', 'Bunnys Snacks Division');
    const title    = getPref('ppt_title', 'STANDS & POSM DEPLOYMENT');
    const dateStr  = new Date().toLocaleDateString('en-GB', { day:'2-digit', month:'long', year:'numeric' });

    // Blue accent bar
    sl.addShape(SHT.rect, { x:0, y:0, w:0.06, h:SLIDE_H, fill:{color:C.blue}, line:{type:'none'} });

    // Dot grid (decoration)
    for (let r=0;r<7;r++) for (let c=0;c<9;c++) {
        sl.addShape(SHT.ellipse, { x:8.8+c*0.52, y:0.5+r*0.52, w:0.07, h:0.07, fill:{color:C.navy3}, line:{type:'none'} });
    }

    // Main title
    sl.addText(title, { x:0.8, y:1.9, w:8.5, h:1.5, fontSize:50, bold:true, color:C.white, fontFace:'Calibri', charSpacing:1.5 });

    // Separator line
    sl.addShape(SHT.rect, { x:0.8, y:3.55, w:4.8, h:0.024, fill:{color:C.blue}, line:{type:'none'} });

    // Division
    sl.addText(division.toUpperCase(), { x:0.8, y:3.75, w:9, h:0.45, fontSize:14, italic:true, color:C.gray, fontFace:'Calibri', charSpacing:3 });

    // Info pills
    [
        { label:'SHOPS', value:String(shopCount) },
        { label:'DATE', value:dateStr },
        { label:'STATUS', value:'CONFIDENTIAL' },
    ].forEach((p, i) => {
        const px = 0.8 + i * 2.7;
        sl.addShape(SHT.rect, { x:px, y:5.3, w:2.5, h:0.78, fill:{color:C.navy2}, line:{color:C.navy3, width:0.75} });
        sl.addText(p.value, { x:px+0.14, y:5.38, w:2.2, h:0.36, fontSize:14, bold:true, color:C.white, fontFace:'Calibri' });
        sl.addText(p.label, { x:px+0.14, y:5.74, w:2.2, h:0.24, fontSize:7.5, color:C.gray3, fontFace:'Calibri', charSpacing:2 });
    });

    // Watermark footer
    sl.addText('HAWKSPOT FIELD INTELLIGENCE · CONFIDENTIAL', {
        x:0, y:SLIDE_H-0.32, w:SLIDE_W, h:0.26, fontSize:7, color:C.navy3, fontFace:'Calibri', align:'center', charSpacing:5,
    });
}

// ─── TOC Slide ────────────────────────────────────────────────────────────────
function buildTocSlide(pptx, selected) {
    const sl  = pptx.addSlide();
    const SHT = pptx.ShapeType;
    sl.background = { fill: C.bg };

    addHeader(pptx, sl, 'SHOP DIRECTORY', 'Field Submission Overview', null, null, null, null);

    const cols  = 2;
    const itemH = 0.7;
    const itemW = (CONTENT_W - GAP) / cols;
    const startY = CONTENT_Y + 0.1;

    selected.forEach((sub, idx) => {
        const col = idx % cols;
        const row = Math.floor(idx / cols);
        const x   = PAD + col * (itemW + GAP);
        const y   = startY + row * (itemH + 0.1);
        if (y + itemH > SLIDE_H - 0.15) return;

        const imgCount = (STATE.photosMap[sub.id] || []).length;

        sl.addShape(SHT.rect, { x, y, w:itemW, h:itemH, fill:{color:C.navy2}, line:{color:C.navy3, width:0.5} });

        sl.addText(String(idx+1).padStart(2,'0'), { x:x+0.14, y:y+0.11, w:0.52, h:0.42, fontSize:20, bold:true, color:C.blue, fontFace:'Calibri' });
        sl.addText(sub.shop_name, { x:x+0.7, y:y+0.1, w:itemW-1.1, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
        sl.addText(`${sub.area}  ·  ${sub.distributor}`, { x:x+0.7, y:y+0.39, w:itemW-1.1, h:0.22, fontSize:8.5, color:C.gray2, fontFace:'Calibri' });

        sl.addShape(SHT.rect, { x:x+itemW-0.78, y:y+0.2, w:0.65, h:0.28, fill:{color:C.navy3}, line:{type:'none'} });
        sl.addText(`${imgCount} IMG`, { x:x+itemW-0.77, y:y+0.21, w:0.63, h:0.24, fontSize:8, bold:true, color:C.blue, fontFace:'Calibri', align:'center' });
    });
}

// ─── Evidence Slide ───────────────────────────────────────────────────────────
function buildEvidenceSlide(pptx, sub, photos, shopIdx, totalShops, pageNum, totalPages) {
    const sl  = pptx.addSlide();
    const SHT = pptx.ShapeType;
    sl.background = { fill: C.bg };

    const subtitle = pageNum > 1
        ? `Continued (${pageNum} of ${totalPages})  ·  ${sub.distributor}`
        : `${sub.area}  ·  ${sub.distributor}`;

    addHeader(pptx, sl, sub.shop_name, subtitle, shopIdx, totalShops, pageNum, totalPages);

    const g     = GRIDS[photos.length] || GRIDS[6];
    const cellW = (CONTENT_W - GAP * (g.c - 1)) / g.c;
    const cellH = (CONTENT_H - GAP * (g.r - 1)) / g.r;

    photos.forEach((photo, idx) => {
        const col   = idx % g.c;
        const row   = Math.floor(idx / g.c);
        const cellX = PAD + col * (cellW + GAP);
        const cellY = CONTENT_Y + row * (cellH + GAP);

        // White frame
        sl.addShape(SHT.rect, { x:cellX, y:cellY, w:cellW, h:cellH, fill:{color:C.white}, line:{color:C.border, width:0.75} });

        // Image
        sl.addImage({
            path: photo.public_url,
            x: cellX + 0.02, y: cellY + 0.02,
            w: cellW - 0.04, h: cellH - 0.04,
            sizing: { type:'contain', w: cellW - 0.04, h: cellH - 0.04 },
        });

        // Image index
        const globalIdx = (pageNum - 1) * MAX_SLIDE + idx + 1;
        sl.addText(String(globalIdx), {
            x: cellX + cellW - 0.3, y: cellY + cellH - 0.22, w:0.26, h:0.18,
            fontSize:6.5, color:'CCCCCC', fontFace:'Calibri', align:'right',
        });
    });

    addFooter(sl, sub.submitted_at);
}

// ─── Summary Slide ────────────────────────────────────────────────────────────
function buildSummarySlide(pptx, selected) {
    const sl  = pptx.addSlide();
    const SHT = pptx.ShapeType;
    sl.background = { fill: C.navy };

    sl.addShape(SHT.rect, { x:0, y:0, w:0.06, h:SLIDE_H, fill:{color:C.blue}, line:{type:'none'} });

    sl.addText('EXECUTION SUMMARY', {
        x:0.8, y:1.1, w:11, h:1.0, fontSize:44, bold:true, color:C.white, fontFace:'Calibri', charSpacing:2,
    });

    const totalPhotos  = countPhotos(selected);
    const areas        = [...new Set(selected.map(s => s.area))];
    const distributors = [...new Set(selected.map(s => s.distributor))];

    [
        { label:'SHOPS COVERED', value:selected.length },
        { label:'TOTAL IMAGES',  value:totalPhotos },
        { label:'AREAS',         value:areas.length },
        { label:'DISTRIBUTORS',  value:distributors.length },
    ].forEach((st, i) => {
        const x = 0.8 + i * 3.1;
        sl.addShape(SHT.rect, { x, y:2.7, w:2.8, h:1.5, fill:{color:C.navy2}, line:{color:C.navy3, width:0.75} });
        sl.addText(String(st.value), { x:x+0.15, y:2.82, w:2.5, h:0.7, fontSize:38, bold:true, color:C.blue, fontFace:'Calibri' });
        sl.addText(st.label, { x:x+0.15, y:3.58, w:2.5, h:0.42, fontSize:8.5, color:C.gray3, fontFace:'Calibri', charSpacing:1.5 });
    });

    sl.addText('AREAS COVERED', { x:0.8, y:4.65, w:11, h:0.28, fontSize:8, color:C.gray3, fontFace:'Calibri', charSpacing:2.5 });
    sl.addText(areas.join('  ·  '), { x:0.8, y:4.96, w:11.5, h:0.36, fontSize:12, color:C.gray, fontFace:'Calibri' });

    sl.addText('DISTRIBUTORS', { x:0.8, y:5.5, w:11, h:0.28, fontSize:8, color:C.gray3, fontFace:'Calibri', charSpacing:2.5 });
    sl.addText(distributors.join('  ·  '), { x:0.8, y:5.8, w:11.5, h:0.36, fontSize:12, color:C.gray, fontFace:'Calibri' });

    const now = new Date().toLocaleDateString('en-GB', { day:'2-digit', month:'long', year:'numeric' });
    sl.addText(`Generated by HawkSpot Field Intelligence  ·  ${now}  ·  CONFIDENTIAL`, {
        x:0, y:SLIDE_H-0.34, w:SLIDE_W, h:0.28, fontSize:7, color:C.navy3, fontFace:'Calibri', align:'center', charSpacing:2,
    });
}

// ─── Shared Header ────────────────────────────────────────────────────────────
function addHeader(pptx, sl, title, subtitle, shopIdx, totalShops, pageNum, totalPages) {
    const SHT = pptx.ShapeType;

    sl.addShape(SHT.rect, { x:0, y:0, w:SLIDE_W, h:HEADER_H, fill:{color:C.navy}, line:{type:'none'} });
    sl.addShape(SHT.rect, { x:0, y:0, w:0.05, h:HEADER_H, fill:{color:C.blue}, line:{type:'none'} });

    sl.addText(title.toUpperCase(), { x:0.18, y:0.18, w:10.5, h:0.5, fontSize:21, bold:true, color:C.white, fontFace:'Calibri', charSpacing:0.3 });
    sl.addText(subtitle, { x:0.18, y:0.68, w:10.5, h:0.26, fontSize:10, color:C.gray2, fontFace:'Calibri' });

    if (shopIdx !== null && totalShops !== null) {
        sl.addText(`${shopIdx} / ${totalShops}`, { x:SLIDE_W-1.7, y:0.12, w:1.56, h:0.38, fontSize:12, bold:true, color:C.blue, fontFace:'Calibri', align:'right' });
        sl.addText('SHOPS', { x:SLIDE_W-1.7, y:0.48, w:1.56, h:0.2, fontSize:7, color:C.gray3, fontFace:'Calibri', align:'right', charSpacing:2 });
    }

    if (pageNum > 1) {
        sl.addText(`PG ${pageNum}/${totalPages}`, { x:SLIDE_W-1.7, y:0.72, w:1.56, h:0.2, fontSize:7, color:C.gray3, fontFace:'Calibri', align:'right' });
    }
}

// ─── Shared Footer ────────────────────────────────────────────────────────────
function addFooter(sl, submittedAt) {
    let ds = '';
    try { ds = submittedAt ? new Date(submittedAt).toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' }) : ''; } catch(_){}
    sl.addText(`Submitted: ${ds}  ·  HawkSpot Field Intelligence  ·  CONFIDENTIAL`, {
        x:PAD, y:SLIDE_H-0.18, w:SLIDE_W-PAD*2, h:0.15, fontSize:6.5, color:'94A3B8', fontFace:'Calibri', align:'right',
    });
}

// ─── Utils ────────────────────────────────────────────────────────────────────
function chunkArr(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
}

function countPhotos(selected) {
    return selected.reduce((a, s) => a + (STATE.photosMap[s.id] || []).length, 0);
}

function getPref(key, fallback) {
    return localStorage.getItem(key) || fallback;
}
