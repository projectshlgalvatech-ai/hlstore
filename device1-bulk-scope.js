/* ---------------- Bulk import materials from an uploaded Excel file (Materials
   section, Device 1). Reuses the ExcelJS library already loaded for the "Save
   entries" export, so no extra dependency. Column headers are matched loosely
   (case/spacing-insensitive, several aliases per field) so the user's own
   sheet doesn't need to match a fixed layout exactly — only "Material Name"
   is required. Blank optional cells fall back to sensible defaults instead of
   blocking the row, since a bulk stock-take sheet often won't have every
   field filled in (price, PO/SO refs, etc). A row is skipped as a duplicate
   under the exact same rule the manual "New material" form uses: same name +
   type + size + grade as one that already exists — so same-name-different-
   size/grade rows (e.g. "MS HR Sheet" in several thicknesses) are each kept
   as their own material, never collapsed into one. ---------------- */
const MATERIAL_IMPORT_FIELD_ALIASES = {
  name: ['material name','name','material'],
  type: ['type','material type','ut'],
  category: ['category','material category'],
  rack: ['rack','rack / location','rack/location','location'],
  size: ['size','size/dimension','size / dimension','dimension','size (dimension)'],
  grade: ['grade','grade / quality','grade/quality','quality'],
  price: ['price','price / unit','price/unit','unit price'],
  unit: ['unit','sales unit'],
  po: ['po number','po number (reference)','po'],
  so: ['so number','so number (reference)','so'],
  opening: ['opening stock','opening stock (nos.)','opening'],
  productCode: ['product code','product no','product no.','product number','code']
};
function normalizeHeaderText(v){ return String(v==null?'':v).trim().toLowerCase(); }
// Pulls the first number out of a cell's text even if it's mixed with letters
// or units — "100abc" -> 100, "10 Nos" -> 10, "abc" -> NaN. Used for any
// numeric import column (Qty Needed, Price, Opening Stock) so a stray letter
// or unit typed into a quantity cell doesn't silently drop the whole row —
// only a cell with no number in it at all is treated as invalid.
function extractLeadingNumber(text){
  const m = String(text==null?'':text).match(/-?\d+(\.\d+)?/);
  return m ? Number(m[0]) : NaN;
}
function inferMaterialTypeFromText(text){
  const t = (text||'').toLowerCase();
  if(/\bss\b|stainless/.test(t)) return 'SS';
  if(/\bms\b|mild steel/.test(t)) return 'MS';
  if(/aluminium|aluminum|plastic|\bpp\b|rubber/.test(t)) return 'Plastic';
  if(/electric|electronic|cable|wire/.test(t)) return 'Electronic';
  return 'Other';
}
async function downloadMaterialImportTemplate(){
  if(typeof ExcelJS==='undefined'){ toast('Could not build template — check your internet connection', true); return; }
  const wb = new ExcelJS.Workbook();
  wb.creator = 'HL Galvatech'; wb.created = new Date();
  const instr = wb.addWorksheet('Instructions');
  instr.getColumn(1).width = 90;
  [
    ['HL Galvatech — Material Import Template'],
    [],
    ['How to use this file:'],
    ['1. Open the "Materials" tab.'],
    ['2. Row 2 is an example — replace it with your own first material, or delete it.'],
    ['3. Add one row per material. "Material Name" is the only required column.'],
    ["4. Leave any other column blank if you don't have that detail yet — the system fills in a sensible default."],
    ['5. "Product Code" is optional but must be unique per material if you use it — it\'s what lets the system tell apart several identical-name items (e.g. 3 grinders) when they\'re sent to a site or for repair. Leave it blank to skip this for a row.'],
    ['6. Save the file, then in the app go to Materials → Upload Excel file → select this file.']
  ].forEach(r=>instr.addRow(r));
  instr.getCell('A1').font = {bold:true, size:14};
  const mat = wb.addWorksheet('Materials');
  const headers = ['Material Name','Type','Category','Rack / Location','Size/Dimension','Grade','Price','Unit','PO Number','SO Number','Opening Stock','Product Code'];
  mat.addRow(headers).font = {bold:true};
  mat.addRow(['MS HR Sheet','MS','MS Sheet','','4 x 1500 x 3000 mm','IS 2062 E250','','','','','','']);
  mat.columns.forEach(c=>c.width=18);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = 'HL_Galvatech_Material_Import_Template.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 4000);
}
window.downloadMaterialImportTemplate = downloadMaterialImportTemplate;
async function bulkImportMaterialsFromExcel(fileInputEl){
  const file = fileInputEl.files && fileInputEl.files[0];
  if(!file) return;
  if(typeof ExcelJS==='undefined'){ toast('Could not read file — ExcelJS did not load. Check your internet connection.', true); fileInputEl.value=''; return; }
  try{
    const buf = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    let sheet = wb.worksheets.find(s=>normalizeHeaderText(s.name)==='materials')
      || wb.worksheets.find(s=>normalizeHeaderText(s.name)!=='instructions')
      || wb.worksheets[0];
    if(!sheet){ toast('No sheet found in that file', true); fileInputEl.value=''; return; }

    // Map each known field to a column index by matching the header row
    // (row 1) against the alias list above.
    const headerRow = sheet.getRow(1);
    const colOf = {};
    headerRow.eachCell({includeEmpty:false}, (cell, colNumber)=>{
      const h = normalizeHeaderText(cell.value);
      for(const field in MATERIAL_IMPORT_FIELD_ALIASES){
        if(colOf[field]) continue;
        if(MATERIAL_IMPORT_FIELD_ALIASES[field].includes(h)) colOf[field] = colNumber;
      }
    });
    if(!colOf.name){ toast('Could not find a "Material Name" column in that file', true); fileInputEl.value=''; return; }
    const cellText = (row, col)=> col ? String(row.getCell(col).value==null?'':row.getCell(col).value).trim() : '';

    const existingKeys = new Set(DB.materials.map(m=>
      `${m.name.toLowerCase()}|${m.type}|${(m.size||'').toLowerCase()}|${(m.grade||'').toLowerCase()}`));
    const existingCodes = new Set(DB.materials.filter(m=>m.productCode).map(m=>m.productCode.toLowerCase()));
    let added=0, skippedDup=0, skippedNoName=0, zeroPrice=0, skippedDupCode=0;
    const newCategories=[], newTypes=[];
    const toAdd=[];

    sheet.eachRow({includeEmpty:false}, (row, rowNumber)=>{
      if(rowNumber===1) return; // header
      const name = cellText(row, colOf.name);
      if(!name){ skippedNoName++; return; }
      const size = cellText(row, colOf.size);
      const grade = cellText(row, colOf.grade);
      let type = cellText(row, colOf.type);
      const category = cellText(row, colOf.category) || 'General';
      if(!type){ type = inferMaterialTypeFromText(`${category} ${name}`); }
      const key = `${name.toLowerCase()}|${type}|${size.toLowerCase()}|${grade.toLowerCase()}`;
      if(existingKeys.has(key)){ skippedDup++; return; }
      existingKeys.add(key);

      if(type.toLowerCase()!=='other' && !allMaterialTypeOptions().some(t=>t.toLowerCase()===type.toLowerCase()) && !newTypes.some(t=>t.toLowerCase()===type.toLowerCase())){
        newTypes.push(type);
      }
      if(!DB.categories.some(c=>c.toLowerCase()===category.toLowerCase()) && !newCategories.some(c=>c.toLowerCase()===category.toLowerCase())){
        newCategories.push(category);
      }
      const priceRaw = cellText(row, colOf.price);
      const price = extractLeadingNumber(priceRaw) || 0;
      if(!priceRaw) zeroPrice++;
      const unitRaw = cellText(row, colOf.unit).toLowerCase();
      const unit = UNITS.includes(unitRaw) ? unitRaw : 'pcs';
      const opening = extractLeadingNumber(cellText(row, colOf.opening)) || 0;
      // Product Code must stay unique across the whole materials list to be
      // useful for tracking — a row whose code collides with one already in
      // the system, or with an earlier row in this same file, still gets
      // imported, just without that code (edit it in afterward once it's
      // fixed in the sheet).
      let productCode = cellText(row, colOf.productCode);
      if(productCode){
        if(existingCodes.has(productCode.toLowerCase())){ skippedDupCode++; productCode = ''; }
        else existingCodes.add(productCode.toLowerCase());
      }

      toAdd.push({
        id: uid(), name, type, category,
        rack: cellText(row, colOf.rack),
        size, grade, price, unit,
        poNumber: cellText(row, colOf.po), soNumber: cellText(row, colOf.so),
        dateAdded: todayStr(), trackNos: true, productCode, _opening: opening
      });
      added++;
    });

    if(!added && !skippedDup && !skippedNoName){ toast('No material rows found in that file', true); fileInputEl.value=''; return; }

    if(newTypes.length) DB.customMaterialTypes = [...(DB.customMaterialTypes||[]), ...newTypes];
    if(newCategories.length) DB.categories = [...DB.categories, ...newCategories];
    toAdd.forEach(mat=>{
      const opening = mat._opening; delete mat._opening;
      DB.materials.push(mat);
      if(opening>0) addStock(mat.id, DB.locations[0], opening);
    });

    await saveKey('materials'); await saveKey('stock');
    if(newTypes.length) await saveKey('customMaterialTypes');
    if(newCategories.length) await saveKey('categories');

    fileInputEl.value = '';
    showModal(`
      <h3>Import complete</h3>
      <div class="row" style="flex-direction:column;gap:6px;margin:10px 0">
        <div><b>${added}</b> material${added===1?'':'s'} added</div>
        <div>${skippedDup} skipped as duplicate${skippedDup===1?'':'s'} of a material already in the system</div>
        ${skippedNoName?`<div>${skippedNoName} row${skippedNoName===1?'':'s'} skipped — no material name</div>`:''}
        ${zeroPrice?`<div class="status-low">${zeroPrice} added material${zeroPrice===1?'':'s'} had no price — set to 0, edit later if needed</div>`:''}
        ${skippedDupCode?`<div class="status-low">${skippedDupCode} row${skippedDupCode===1?'':'s'} had a Product Code already used elsewhere — material${skippedDupCode===1?' was':'s were'} still added, just without that code</div>`:''}
        ${newCategories.length?`<div>New categories added: ${newCategories.join(', ')}</div>`:''}
        ${newTypes.length?`<div>New material types added: ${newTypes.join(', ')}</div>`:''}
      </div>
      <div class="modal-actions"><button class="btn" type="button" onclick="closeModal()">Done</button></div>
    `);
    render();
  }catch(e){
    console.error('material import failed', e);
    toast('Could not read that file — make sure it is a valid .xlsx file', true);
    fileInputEl.value = '';
  }
}
window.bulkImportMaterialsFromExcel = bulkImportMaterialsFromExcel;

/* ---------------- Bulk attach Size / Grade to EXISTING materials from an uploaded
   Excel file (Materials section, Device 1). Covers the case where a material was
   created without knowing its size/grade yet, and that detail only turns up later
   (e.g. from a supplier sheet) — instead of opening "Edit" on each material one by
   one, fill one sheet and upload it once. Matching prefers Product Code (exact,
   unambiguous even across same-name variants); falls back to Material Name only
   when the name is unique across the material list. A name that matches more than
   one existing material is left alone and reported as skipped, rather than guessing
   which variant to update. Only cells that actually have a value in the sheet
   overwrite the material's Size/Grade — a blank Size or Grade cell leaves that
   field untouched, so a sheet that only fills in Grade doesn't erase Size. */
const MATERIAL_SIZE_GRADE_FIELD_ALIASES = {
  productCode: ['product code','code'],
  name: ['material name','name','material'],
  size: ['size','size/dimension','size / dimension','dimension','size (dimension)'],
  grade: ['grade','grade / quality','grade/quality','quality']
};
async function downloadMaterialSizeGradeTemplate(){
  if(typeof ExcelJS==='undefined'){ toast('Could not build template — check your internet connection', true); return; }
  const wb = new ExcelJS.Workbook();
  wb.creator = 'HL Galvatech'; wb.created = new Date();
  const instr = wb.addWorksheet('Instructions');
  instr.getColumn(1).width = 92;
  [
    ['HL Galvatech — Attach Size / Grade Template'],
    [],
    ['How to use this file:'],
    ['1. Open the "Materials" tab — it is pre-filled with every material currently in the system, including ones already missing Size and/or Grade.'],
    ['2. Fill in the Size and/or Grade columns wherever you now have that information. Leave a cell blank to leave that field unchanged.'],
    ['3. Do not edit the Product Code or Material Name columns — they are what the system uses to find the right material. If Product Code is filled in, it is used; otherwise the Material Name is used (only works if that name is unique).'],
    ['4. Save the file, then in the app go to Materials → Upload Excel file (Attach Size / Grade) → select this file.']
  ].forEach(r=>instr.addRow(r));
  instr.getCell('A1').font = {bold:true, size:14};
  const mat = wb.addWorksheet('Materials');
  const headers = ['Product Code','Material Name','Size','Grade'];
  mat.addRow(headers).font = {bold:true};
  DB.materials.forEach(m=>mat.addRow([m.productCode||'', m.name, m.size||'', m.grade||'']));
  mat.columns.forEach(c=>c.width=22);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = 'HL_Galvatech_Material_Size_Grade_Template.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 4000);
}
window.downloadMaterialSizeGradeTemplate = downloadMaterialSizeGradeTemplate;
async function bulkUpdateMaterialSizeGradeFromExcel(fileInputEl){
  const file = fileInputEl.files && fileInputEl.files[0];
  if(!file) return;
  if(typeof ExcelJS==='undefined'){ toast('Could not read file — ExcelJS did not load. Check your internet connection.', true); fileInputEl.value=''; return; }
  try{
    const buf = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    let sheet = wb.worksheets.find(s=>normalizeHeaderText(s.name)==='materials')
      || wb.worksheets.find(s=>normalizeHeaderText(s.name)!=='instructions')
      || wb.worksheets[0];
    if(!sheet){ toast('No sheet found in that file', true); fileInputEl.value=''; return; }

    const headerRow = sheet.getRow(1);
    const colOf = {};
    headerRow.eachCell({includeEmpty:false}, (cell, colNumber)=>{
      const h = normalizeHeaderText(cell.value);
      for(const field in MATERIAL_SIZE_GRADE_FIELD_ALIASES){
        if(colOf[field]) continue;
        if(MATERIAL_SIZE_GRADE_FIELD_ALIASES[field].includes(h)) colOf[field] = colNumber;
      }
    });
    if(!colOf.name && !colOf.productCode){ toast('Could not find a "Material Name" or "Product Code" column in that file', true); fileInputEl.value=''; return; }
    const cellText = (row, col)=> col ? String(row.getCell(col).value==null?'':row.getCell(col).value).trim() : '';

    const byCode = new Map(DB.materials.filter(m=>m.productCode).map(m=>[m.productCode.toLowerCase(), m]));
    const nameCounts = new Map();
    DB.materials.forEach(m=>{ const k=m.name.toLowerCase(); nameCounts.set(k, (nameCounts.get(k)||0)+1); });
    const byUniqueName = new Map(DB.materials.filter(m=>nameCounts.get(m.name.toLowerCase())===1).map(m=>[m.name.toLowerCase(), m]));

    let updated=0, skippedNoMatch=0, skippedAmbiguous=0, skippedNoChange=0, skippedNoKey=0;
    const touched = new Set();

    sheet.eachRow({includeEmpty:false}, (row, rowNumber)=>{
      if(rowNumber===1) return; // header
      const code = cellText(row, colOf.productCode);
      const name = cellText(row, colOf.name);
      const size = cellText(row, colOf.size);
      const grade = cellText(row, colOf.grade);
      if(!code && !name){ skippedNoKey++; return; }
      if(!size && !grade){ return; } // nothing to update on this row, don't count as anything

      let mat = null;
      if(code){
        mat = byCode.get(code.toLowerCase()) || null;
        if(!mat){ skippedNoMatch++; return; }
      } else {
        const nameKey = name.toLowerCase();
        if(nameCounts.get(nameKey)>1){ skippedAmbiguous++; return; }
        mat = byUniqueName.get(nameKey) || null;
        if(!mat){ skippedNoMatch++; return; }
      }

      let changed = false;
      if(size && mat.size!==size){ mat.size = size; changed = true; }
      if(grade && mat.grade!==grade){ mat.grade = grade; changed = true; }
      if(changed){ updated++; touched.add(mat.id); } else { skippedNoChange++; }
    });

    if(!updated){
      toast(skippedAmbiguous||skippedNoMatch ? 'No materials updated — see details' : 'No Size/Grade values found to update', true);
    }
    if(updated) await saveKey('materials');

    fileInputEl.value = '';
    showModal(`
      <h3>Size / Grade update complete</h3>
      <div class="row" style="flex-direction:column;gap:6px;margin:10px 0">
        <div><b>${updated}</b> material${updated===1?'':'s'} updated</div>
        ${skippedNoMatch?`<div class="status-low">${skippedNoMatch} row${skippedNoMatch===1?'':'s'} skipped — no matching material found</div>`:''}
        ${skippedAmbiguous?`<div class="status-low">${skippedAmbiguous} row${skippedAmbiguous===1?'':'s'} skipped — name matches more than one material; add a Product Code to disambiguate</div>`:''}
        ${skippedNoKey?`<div>${skippedNoKey} row${skippedNoKey===1?'':'s'} skipped — no Product Code or Material Name</div>`:''}
      </div>
      <div class="modal-actions"><button class="btn" type="button" onclick="closeModal()">Done</button></div>
    `);
    render();
  }catch(e){
    console.error('material size/grade update failed', e);
    toast('Could not read that file — make sure it is a valid .xlsx file', true);
    fileInputEl.value = '';
  }
}
window.bulkUpdateMaterialSizeGradeFromExcel = bulkUpdateMaterialSizeGradeFromExcel;

/* ---------------- Bulk import MULTIPLE materials into one SO/Projects product from an
   uploaded Excel file (Device 1). Same idea as the Materials-section import above, just
   aimed at a single product's material list: instead of submitting the "+ Add material"
   form once per line (which gets old fast with 10-20 materials under one product), fill
   a sheet and upload it. Existing materials are matched by name (using size/grade from
   the sheet to pick the right variant when a name has several); anything genuinely new
   is created via the same path the manual form uses, then attached — no separate trip
   to Materials needed. */
const SO_MATERIAL_IMPORT_FIELD_ALIASES = {
  name: ['material name','name','material'],
  qty: ['qty needed','qty','quantity','quantity needed','needed'],
  size: ['size','size/dimension','size / dimension','dimension'],
  grade: ['grade','grade / quality','grade/quality','quality'],
  category: ['category','material category']
};
async function downloadSOMaterialImportTemplate(){
  if(typeof ExcelJS==='undefined'){ toast('Could not build template — check your internet connection', true); return; }
  const wb = new ExcelJS.Workbook();
  wb.creator = 'HL Galvatech'; wb.created = new Date();
  const instr = wb.addWorksheet('Instructions');
  instr.getColumn(1).width = 92;
  [
    ['HL Galvatech — SO / Projects Material Import Template'],
    [],
    ['How to use this file:'],
    ['1. Open the "Materials" tab.'],
    ['2. Row 2 is an example — replace it with your own first material, or delete it.'],
    ['3. Add one row per material needed under this product. "Material Name" and "Qty Needed" are the only required columns.'],
    ["4. If the material already exists in the system, its name alone is enough — leave Size/Grade/Category blank."],
    ['5. If the material is new, fill in Size / Grade / Category if you have them — it will be created automatically and attached to the product in the same step. Leave them blank and it will still be created with sensible defaults.'],
    ['6. Category should be one of MS, SS, Plastic, Rubber, or Other — leave blank to default to Other.'],
    ['7. Save the file, then in the app open the product under its SO and use "Upload Excel file" next to "+ Add material".']
  ].forEach(r=>instr.addRow(r));
  instr.getCell('A1').font = {bold:true, size:14};
  const mat = wb.addWorksheet('Materials');
  const headers = ['Material Name','Qty Needed','Size','Grade','Category'];
  mat.addRow(headers).font = {bold:true};
  mat.addRow(['MS HR Sheet', 10, '4 x 1500 x 3000 mm','IS 2062 E250','MS']);
  mat.columns.forEach(c=>c.width=22);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = 'HL_Galvatech_SO_Material_Import_Template.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 4000);
}
window.downloadSOMaterialImportTemplate = downloadSOMaterialImportTemplate;

async function bulkImportSOMaterialsFromExcel(fileInputEl, soId, productId){
  const file = fileInputEl.files && fileInputEl.files[0];
  if(!file) return;
  if(typeof ExcelJS==='undefined'){ toast('Could not read file — ExcelJS did not load. Check your internet connection.', true); fileInputEl.value=''; return; }
  const so = findSO(soId);
  const product = so && so.products.find(p=>p.id===productId);
  if(!so || !product){ toast('Could not find that product to import into', true); fileInputEl.value=''; return; }
  if(so.status==='completed'){ toast(`SO ${so.soNumber} is marked Complete by Device 3 — no new material can be added.`, true); fileInputEl.value=''; return; }
  try{
    const buf = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    let sheet = wb.worksheets.find(s=>normalizeHeaderText(s.name)==='materials')
      || wb.worksheets.find(s=>normalizeHeaderText(s.name)!=='instructions')
      || wb.worksheets[0];
    if(!sheet){ toast('No sheet found in that file', true); fileInputEl.value=''; return; }

    const headerRow = sheet.getRow(1);
    const colOf = {};
    headerRow.eachCell({includeEmpty:false}, (cell, colNumber)=>{
      const h = normalizeHeaderText(cell.value);
      for(const field in SO_MATERIAL_IMPORT_FIELD_ALIASES){
        if(colOf[field]) continue;
        if(SO_MATERIAL_IMPORT_FIELD_ALIASES[field].includes(h)) colOf[field] = colNumber;
      }
    });
    if(!colOf.name){ toast('Could not find a "Material Name" column in that file', true); fileInputEl.value=''; return; }
    if(!colOf.qty){ toast('Could not find a "Qty Needed" column in that file', true); fileInputEl.value=''; return; }
    const cellText = (row, col)=> col ? String(row.getCell(col).value==null?'':row.getCell(col).value).trim() : '';

    // ExcelJS row iteration is synchronous, but resolving/creating materials is async
    // (createMaterialFromDraft awaits a save) — collect the raw rows first, then process
    // them one at a time below so each row's material exists before the next is checked.
    const rows = [];
    sheet.eachRow({includeEmpty:false}, (row, rowNumber)=>{
      if(rowNumber===1) return; // header
      rows.push({
        name: cellText(row, colOf.name),
        qty: extractLeadingNumber(cellText(row, colOf.qty)),
        size: cellText(row, colOf.size),
        grade: cellText(row, colOf.grade),
        category: cellText(row, colOf.category)
      });
    });

    let added=0, skippedNoName=0, skippedNoQty=0, skippedDupOnProduct=0, createdNew=0;
    const newlyCreated = [];
    // Loose text match for Size/Grade — trims, collapses internal whitespace and
    // lowercases, so "10 Liter" and "10  liter" line up; a second, tighter pass
    // also strips all whitespace so "10ltr" lines up with "10 ltr". This is only
    // used to recognise an existing variant that's really the same thing typed
    // slightly differently — it never invents a match between genuinely
    // different sizes.
    const normSG = s=>String(s||'').toLowerCase().replace(/\s+/g,' ').trim();
    const tightSG = s=>normSG(s).replace(/\s+/g,'');

    for(const r of rows){
      if(!r.name){ skippedNoName++; continue; }
      if(!r.qty || r.qty<=0){ skippedNoQty++; continue; }
      const nameMatches = DB.materials.filter(m=>m.name.toLowerCase()===r.name.toLowerCase());
      let mat = null;
      if(nameMatches.length===1){
        mat = nameMatches[0];
      } else if(nameMatches.length>1){
        // Several variants share this name — pick the one whose Size/Grade line
        // up with the sheet (allowing for spacing/case differences like "10ltr"
        // vs "10 Liter"). If none line up, fall through below and create a new
        // variant from what the sheet gives, instead of skipping the row.
        mat = nameMatches.find(m=>normSG(m.size)===normSG(r.size) && normSG(m.grade)===normSG(r.grade))
           || nameMatches.find(m=>tightSG(m.size)===tightSG(r.size) && tightSG(m.grade)===tightSG(r.grade))
           || null;
      }
      if(!mat){
        // No existing material this row can confidently attach to — either the
        // name is brand new, or it matches several variants but none share this
        // Size/Grade. Either way, create it (same path the manual "new material"
        // form uses) using the Size/Grade exactly as typed in the sheet, and use
        // it immediately — a row is never dropped just because its Size/Grade
        // text doesn't byte-for-byte match what's already on file.
        const category = r.category || 'Other';
        const categoryIsOther = !SO_MATERIAL_CATEGORIES.includes(category);
        const draft = { name: r.name, type: category, typeIsOther: categoryIsOther, category: DB.categories[0]||'General', size: r.size, grade: r.grade, price:0, unit:'pcs', rack:'', trackNos:true };
        const created = await createMaterialFromDraft(draft);
        if(created==='duplicate'){
          // An exact duplicate (name+type+size+grade) already exists — use it
          // rather than skipping the row.
          mat = DB.materials.find(m=>m.name.toLowerCase()===r.name.toLowerCase() && normSG(m.size)===normSG(r.size) && normSG(m.grade)===normSG(r.grade)) || nameMatches[0] || null;
        } else {
          mat = created; createdNew++; newlyCreated.push(mat.name);
        }
      }
      if(!mat) continue; // should not happen, but guards against a stray null
      if(product.materials.some(x=>x.materialId===mat.id)){ skippedDupOnProduct++; continue; }
      product.materials.push({id:uid(), materialId:mat.id, materialName:mat.name, qtyNeeded:r.qty, qtyFulfilled:0});
      added++;
    }

    if(!added && !skippedDupOnProduct && !skippedNoName && !skippedNoQty){
      toast('No material rows found in that file', true); fileInputEl.value=''; return;
    }

    checkSOCompletion(so);
    await saveKey('soList'); await saveKey('materials');

    fileInputEl.value = '';
    showModal(`
      <h3>Import complete — ${product.name}</h3>
      <div class="row" style="flex-direction:column;gap:6px;margin:10px 0">
        <div><b>${added}</b> material line${added===1?'':'s'} added to "${product.name}"</div>
        ${createdNew?`<div>${createdNew} new material${createdNew===1?'':'s'} created in the master list: ${newlyCreated.join(', ')}</div>`:''}
        ${skippedDupOnProduct?`<div>${skippedDupOnProduct} row${skippedDupOnProduct===1?'':'s'} skipped — already listed on this product</div>`:''}
        ${skippedNoName?`<div>${skippedNoName} row${skippedNoName===1?'':'s'} skipped — no material name</div>`:''}
        ${skippedNoQty?`<div>${skippedNoQty} row${skippedNoQty===1?'':'s'} skipped — no valid Qty Needed</div>`:''}
      </div>
      <div class="modal-actions"><button class="btn" type="button" onclick="closeModal()">Done</button></div>
    `);
    render();
  }catch(e){
    console.error('SO material import failed', e);
    toast('Could not read that file — make sure it is a valid .xlsx file', true);
    fileInputEl.value = '';
  }
}
window.bulkImportSOMaterialsFromExcel = bulkImportSOMaterialsFromExcel;

/* ---------------- shared collapsible-panel helper — wraps a panel's list
   content behind a click-to-open header, collapsed by default. Used by
   Factory Use, Material Request, Customers & Suppliers and Categories &
   Locations so their record lists don't take up screen space until opened,
   same idea as the GRN / SO / All Entries collapsible lists. Kept outside DB,
   like the other UI-state objects, so it resets to fully collapsed on reload. ---------------- */
let panelCollapseUIState = {};
let locInvKeyword = '';
function toggleCollapsePanel(id){ panelCollapseUIState[id] = !panelCollapseUIState[id]; render(); }
window.toggleCollapsePanel = toggleCollapsePanel;
function collapsePanel(id, title, bodyHtml, meta){
  const open = !!panelCollapseUIState[id];
  return `<div class="panel">
    <button type="button" class="panel-toggle" onclick="toggleCollapsePanel('${id}')">
      <span class="chev">${open?'▾':'▸'}</span> ${title}${meta?` <span class="panel-toggle-meta">${meta}</span>`:''}
    </button>
    <div style="display:${open?'block':'none'}">${bodyHtml}</div>
  </div>`;
}
// Same as collapsePanel, but the wrapping <div> also carries id="${id}" so a KPI
// card can jump straight to it (see kpi()/scrollToAnchor below), auto-opening it.
function collapsePanelAnchored(id, title, bodyHtml, meta){
  const open = !!panelCollapseUIState[id];
  return `<div class="panel" id="${id}">
    <button type="button" class="panel-toggle" onclick="toggleCollapsePanel('${id}')">
      <span class="chev">${open?'▾':'▸'}</span> ${title}${meta?` <span class="panel-toggle-meta">${meta}</span>`:''}
    </button>
    <div style="display:${open?'block':'none'}">${bodyHtml}</div>
  </div>`;
}

/* ---------------- generic Day / Last 6 Days / Month / Year / All scope engine —
   same picker UI as the GRN and SO/Projects lists, generalised so any dated list
   (recent gate entries, recent issues, damaged stock, SO completion, etc.) can
   reuse it instead of dumping every record on screen at once. Keyed by an id
   per list so several lists can each remember their own scope. Kept outside DB,
   like the other UI-state objects, so it resets on reload. ---------------- */
let scopeUIState = {};
function getScopeState(id, defaultScope){
  if(!scopeUIState[id]) scopeUIState[id] = {scope:defaultScope||'week', customDate:null, customMonth:null, customYear:null, calendarOpen:false};
  return scopeUIState[id];
}
function scopeFilterRows(id, rows, dateKey){
  dateKey = dateKey || 'date';
  const s = getScopeState(id);
  const now = new Date();
  if(s.scope==='today') return rows.filter(r=>r[dateKey]===todayStr());
  if(s.scope==='week') return rows.filter(r=>{ const d=daysSinceToday(r[dateKey]); return d>=0 && d<=6; });
  if(s.scope==='month') return rows.filter(r=>{ if(!r[dateKey]) return false; const d=new Date(r[dateKey]+'T00:00:00'); return d.getFullYear()===now.getFullYear() && d.getMonth()===now.getMonth(); });
  if(s.scope==='year') return rows.filter(r=>{ if(!r[dateKey]) return false; return new Date(r[dateKey]+'T00:00:00').getFullYear()===now.getFullYear(); });
  if(s.scope==='custom-day') return s.customDate ? rows.filter(r=>r[dateKey]===s.customDate) : [];
  if(s.scope==='custom-month'){
    if(!s.customMonth) return [];
    const [y,m] = s.customMonth.split('-').map(Number);
    return rows.filter(r=>{ if(!r[dateKey]) return false; const d=new Date(r[dateKey]+'T00:00:00'); return d.getFullYear()===y && (d.getMonth()+1)===m; });
  }
  if(s.scope==='custom-year'){
    if(!s.customYear) return [];
    return rows.filter(r=>{ if(!r[dateKey]) return false; return new Date(r[dateKey]+'T00:00:00').getFullYear()===Number(s.customYear); });
  }
  return rows; // 'all'
}
function scopeHeading(id){
  const s = getScopeState(id);
  if(s.scope==='today') return 'Today — '+formatDateNice(todayStr());
  if(s.scope==='week') return 'Last 6 Days';
  if(s.scope==='month') return 'This Month — '+new Date().toLocaleDateString('en-IN',{month:'long',year:'numeric'});
  if(s.scope==='year') return 'This Year — '+new Date().getFullYear();
  if(s.scope==='custom-day') return s.customDate ? formatDateNice(s.customDate) : 'Pick a day from the calendar below';
  if(s.scope==='custom-month') return s.customMonth ? formatMonthNice(s.customMonth) : 'Pick a month from the calendar below';
  if(s.scope==='custom-year') return s.customYear ? ('Year '+s.customYear) : 'Pick a year from the calendar below';
  return 'All';
}
function renderScopeControls(id){
  const s = getScopeState(id);
  const inCalendar = ['custom-day','custom-month','custom-year'].includes(s.scope);
  return `<div class="entries-scope-row">
      <button type="button" class="scope-btn ${s.scope==='today'?'active':''}" onclick="setScope('${id}','today')">Today</button>
      <button type="button" class="scope-btn ${s.scope==='week'?'active':''}" onclick="setScope('${id}','week')">Last 6 Days</button>
      <button type="button" class="scope-btn ${s.scope==='month'?'active':''}" onclick="setScope('${id}','month')">This Month</button>
      <button type="button" class="scope-btn ${s.scope==='year'?'active':''}" onclick="setScope('${id}','year')">This Year</button>
      <button type="button" class="scope-btn ${s.scope==='all'?'active':''}" onclick="setScope('${id}','all')">All</button>
      <button type="button" class="scope-btn ${inCalendar?'active':''}" onclick="toggleScopeCalendar('${id}')">📅 Calendar (past day / month / year)</button>
    </div>
    <div class="entries-calendar-row" style="display:${s.calendarOpen?'flex':'none'}">
      <label>Pick a day
        <input type="date" value="${s.customDate||''}" max="${todayStr()}" onchange="pickScopeDay('${id}', this.value)">
      </label>
      <label>Pick a month/year
        <input type="month" value="${s.customMonth||''}" max="${todayStr().slice(0,7)}" onchange="pickScopeMonth('${id}', this.value)">
      </label>
      <label>Pick a year
        <input type="number" value="${s.customYear||''}" min="2000" max="${new Date().getFullYear()}" placeholder="${new Date().getFullYear()}" style="width:90px" onchange="pickScopeYear('${id}', this.value)">
      </label>
    </div>
    <div class="entries-heading">${scopeHeading(id)}</div>`;
}
function setScope(id, scope){ getScopeState(id).scope = scope; getScopeState(id).calendarOpen = false; render(); }
function toggleScopeCalendar(id){ getScopeState(id).calendarOpen = !getScopeState(id).calendarOpen; render(); }
function pickScopeDay(id, v){ if(!v) return; const s=getScopeState(id); s.customDate=v; s.scope='custom-day'; s.calendarOpen=true; render(); }
function pickScopeMonth(id, v){ if(!v) return; const s=getScopeState(id); s.customMonth=v; s.scope='custom-month'; s.calendarOpen=true; render(); }
function pickScopeYear(id, v){ if(!v) return; const s=getScopeState(id); s.customYear=v; s.scope='custom-year'; s.calendarOpen=true; render(); }
window.setScope = setScope;
window.toggleScopeCalendar = toggleScopeCalendar;
window.pickScopeDay = pickScopeDay;
window.pickScopeMonth = pickScopeMonth;
window.pickScopeYear = pickScopeYear;
// Collapsible panel (collapsed by default) whose body is also scoped by
// day/week/month/year — used for dated lists like recent gate entries, recent
// issues, and damaged stock, so a growing log doesn't force endless scrolling.
function scopedCollapsePanel(id, title, rows, dateKey, tableFn, anchored, defaultScope){
  getScopeState(id, defaultScope); // seed the default scope the first time this list is opened
  const open = !!panelCollapseUIState[id];
  const scoped = scopeFilterRows(id, rows, dateKey);
  const idAttr = anchored ? ` id="${id}"` : '';
  return `<div class="panel"${idAttr}>
    <button type="button" class="panel-toggle" onclick="toggleCollapsePanel('${id}')">
      <span class="chev">${open?'▾':'▸'}</span> ${title} <span class="panel-toggle-meta">${scoped.length}</span>
    </button>
    <div style="display:${open?'block':'none'}">
      ${renderScopeControls(id)}
      ${tableFn(scoped.slice().reverse())}
    </div>
  </div>`;
}

function renderContacts(el){
  const editable = actingAsDevice1();
  const catOptions = allContactCategoryOptions();
  el.innerHTML = `
    <h2 class="section-title">Customers &amp; Suppliers</h2>
    <div class="section-sub">${editable? 'Maintain the database of who materials are bought from and issued to.' : 'Read-only view synced from Device 1.'}</div>
    ${editable? `<div class="panel"><h3>Add contact</h3>
      <form id="c-form">
        <div class="row">
          <div class="field"><label>Name</label><input required id="c-name"></div>
          <div class="field"><label>Type</label><select id="c-type"><option>Supplier</option><option>Customer</option></select></div>
          <div class="field"><label>Phone</label><input id="c-phone"></div>
          <div class="field"><label>Address</label><input id="c-addr"></div>
        </div>
        <div class="row">
          <div class="field">
            <label>Category</label>
            <select id="c-cat">${catOptions.map(c=>`<option>${c}</option>`).join('')}<option value="__other__">Other…</option></select>
          </div>
          <div class="field" id="c-cat-other-wrap" style="display:none;position:relative">
            <label>Other — type category</label>
            <input id="c-cat-other" placeholder="e.g. Rubber, Adhesives…" autocomplete="off">
            <div class="autolist" id="c-cat-other-list"></div>
          </div>
        </div>
        <button class="btn" type="submit">Save contact</button>
      </form></div>` : ''}
    ${collapsePanel('contacts-directory', 'Directory', DB.contacts.length? `<table><thead><tr><th>Name</th><th>Type</th><th>Category</th><th>Phone</th><th>Address</th></tr></thead><tbody>
      ${DB.contacts.map(c=>`<tr><td>${c.name}</td><td>${c.type}</td><td>${c.category||'—'}</td><td>${c.phone||'—'}</td><td>${c.address||'—'}</td></tr>`).join('')}
      </tbody></table>` : `<div class="empty">No contacts yet.</div>`, `${DB.contacts.length} contact${DB.contacts.length===1?'':'s'}`)}`;
  if(editable){
    const catSelect = document.getElementById('c-cat');
    const otherWrap = document.getElementById('c-cat-other-wrap');
    const otherInput = document.getElementById('c-cat-other');
    catSelect.addEventListener('change', ()=>{
      otherWrap.style.display = catSelect.value==='__other__' ? 'flex' : 'none';
    });
    attachAutocomplete(otherInput, document.getElementById('c-cat-other-list'), ()=>(DB.contactCategories||[]).map(c=>({name:c})), (o)=>{ otherInput.value = o.name; });

    document.getElementById('c-form').addEventListener('submit', async (e)=>{
      e.preventDefault();
      let category = catSelect.value;
      if(category==='__other__'){
        category = otherInput.value.trim();
        if(!category){ toast('Type a category, or pick one from the list', true); return; }
        const known = allContactCategoryOptions().some(c=>c.toLowerCase()===category.toLowerCase());
        if(!known){
          DB.contactCategories = DB.contactCategories||[];
          DB.contactCategories.push(category);
          await saveKey('contactCategories');
        }
      }
      DB.contacts.push({id:uid(), name:document.getElementById('c-name').value, type:document.getElementById('c-type').value,
        phone:document.getElementById('c-phone').value, address:document.getElementById('c-addr').value, category});
      await saveKey('contacts'); toast('Contact saved'); e.target.reset(); otherWrap.style.display='none'; render();
    });
  }
}

function renderCategories(el){
  el.innerHTML = `
    <h2 class="section-title">Categories &amp; Locations</h2>
    <div class="section-sub"> Define Categories and locations of Materials.</div>
    <div class="grid2">
      <div class="panel">
        <form id="cat-form" class="row"><div class="field"><input required id="cat-name" placeholder="e.g. Fasteners"></div><button class="btn" type="submit">Add</button></form>
      </div>
      <div class="panel">
        <form id="loc-form" class="row"><div class="field"><input required id="loc-name" placeholder="e.g. Warehouse B"></div><button class="btn" type="submit">Add</button></form>
      </div>
    </div>
    ${collapsePanel('categories-list', 'Categories', `<table><thead><tr><th>Category</th><th>Materials</th><th></th></tr></thead><tbody>${DB.categories.map(c=>{
      const inUse = DB.materials.filter(m=>m.category===c).length;
      return `<tr><td>${c}</td><td>${inUse}</td><td>
        <button class="btn small secondary" onclick="renameCategory('${c.replace(/'/g,"\\'")}')">Rename</button>
        <button class="btn small danger" onclick="deleteCategory('${c.replace(/'/g,"\\'")}')">Delete</button>
      </td></tr>`;
    }).join('')}</tbody></table>`, `${DB.categories.length} categor${DB.categories.length===1?'y':'ies'}`)}
    ${collapsePanel('locations-list', 'Locations', `<table><thead><tr><th>Location</th><th>Stock here</th><th></th></tr></thead><tbody>${DB.locations.map(l=>{
      const totalHere = DB.materials.reduce((s,m)=>s+getStock(m.id,l),0);
      return `<tr><td>${l}</td><td>${totalHere}</td><td>
        <button class="btn small secondary" onclick="renameLocation('${l.replace(/'/g,"\\'")}')">Rename</button>
        <button class="btn small danger" onclick="deleteLocation('${l.replace(/'/g,"\\'")}')">Delete</button>
      </td></tr>`;
    }).join('')}</tbody></table>`, `${DB.locations.length} location${DB.locations.length===1?'':'s'}`)}`;
  document.getElementById('cat-form').addEventListener('submit', async e=>{
    e.preventDefault(); const v=document.getElementById('cat-name').value.trim();
    if(v && !DB.categories.includes(v)){ DB.categories.push(v); await saveKey('categories'); render(); }
  });
  document.getElementById('loc-form').addEventListener('submit', async e=>{
    e.preventDefault(); const v=document.getElementById('loc-name').value.trim();
    if(v && !DB.locations.includes(v)){ DB.locations.push(v); await saveKey('locations'); render(); }
  });
}
async function renameCategory(oldName){
  const next = prompt(`Rename category "${oldName}" to:`, oldName);
  if(next===null) return;
  const v = next.trim();
  if(!v || v===oldName) return;
  if(DB.categories.some(c=>c.toLowerCase()===v.toLowerCase())){ toast('A category with that name already exists', true); return; }
  DB.categories = DB.categories.map(c=> c===oldName? v : c);
  DB.materials.forEach(m=>{ if(m.category===oldName) m.category=v; });
  await saveKey('categories'); await saveKey('materials');
  toast(`Renamed "${oldName}" → "${v}"`); render();
}
async function deleteCategory(name){
  const inUse = DB.materials.filter(m=>m.category===name).length;
  if(inUse){ toast(`Can't delete — ${inUse} material(s) still use "${name}". Rename it or move those materials to another category first.`, true); return; }
  if(DB.categories.length<=1){ toast("Can't delete the last remaining category", true); return; }
  if(!confirm(`Delete category "${name}"?`)) return;
  DB.categories = DB.categories.filter(c=>c!==name);
  await saveKey('categories'); toast(`Deleted "${name}"`); render();
}
async function renameLocation(oldName){
  const next = prompt(`Rename location "${oldName}" to:`, oldName);
  if(next===null) return;
  const v = next.trim();
  if(!v || v===oldName) return;
  if(DB.locations.some(l=>l.toLowerCase()===v.toLowerCase())){ toast('A location with that name already exists', true); return; }
  DB.locations = DB.locations.map(l=> l===oldName? v : l);
  // Stock is keyed "<materialId>|<location>" — rewrite keys so quantities follow the rename.
  const rewriteKeys = (obj)=>{
    const out = {};
    for(const k in obj){
      const idx = k.lastIndexOf('|');
      const loc = k.slice(idx+1);
      const newKey = loc===oldName ? k.slice(0,idx+1)+v : k;
      out[newKey] = obj[k];
    }
    return out;
  };
  DB.stock = rewriteKeys(DB.stock);
  DB.stockNos = rewriteKeys(DB.stockNos||{});
  DB.gateEntries.forEach(g=>{ if(g.location===oldName) g.location=v; });
  DB.issues.forEach(i=>{ if(i.location===oldName) i.location=v; });
  DB.damaged.forEach(d=>{ if(d.location===oldName) d.location=v; });
  (DB.materialRepair||[]).forEach(r=>{ if(r.location===oldName) r.location=v; });
  await saveKey('locations'); await saveKey('stock'); await saveKey('stockNos');
  await saveKey('gateEntries'); await saveKey('issues'); await saveKey('damaged'); await saveKey('materialRepair');
  toast(`Renamed "${oldName}" → "${v}"`); render();
}
async function deleteLocation(name){
  const totalHere = DB.materials.reduce((s,m)=>s+getStock(m.id,name),0);
  if(totalHere>0){ toast(`Can't delete — ${totalHere} unit(s) of stock still sit at "${name}". Transfer them out first.`, true); return; }
  if(DB.locations.length<=1){ toast("Can't delete the last remaining location", true); return; }
  if(!confirm(`Delete location "${name}"?`)) return;
  DB.locations = DB.locations.filter(l=>l!==name);
  await saveKey('locations'); toast(`Deleted "${name}"`); render();
}
window.renameCategory = renameCategory;
window.deleteCategory = deleteCategory;
window.renameLocation = renameLocation;
window.deleteLocation = deleteLocation;

function renderReports1(el){ renderReportsCommon(el); }
function renderReports3(el){ renderReportsCommon(el); }
function renderReportsCommon(el){
  const totalValue = DB.materials.reduce((s,m)=> s + getStock(m.id)*(m.price||0), 0);
  const avgPrice = DB.materials.length? (DB.materials.reduce((s,m)=>s+(m.price||0),0)/DB.materials.length) : 0;
  const byCat = {};
  DB.materials.forEach(m=>{ byCat[m.category]=byCat[m.category]||{count:0,value:0}; byCat[m.category].count++; byCat[m.category].value += getStock(m.id)*(m.price||0); });
  el.innerHTML = `
    <h2 class="section-title">Reports</h2>
    <div class="section-sub">Stock levels, costs, and averages.</div>
    <div class="cards">
      ${kpi('Total stock value', money(totalValue))}
      ${kpi('Average item price', money(avgPrice))}
      ${kpi('Materials', DB.materials.length)}
      ${kpi('Open tickets', DB.tickets.filter(t=>t.status==='open').length, 'bad')}
    </div>
    ${collapsePanel('reports-bycat', 'By category', `<table><thead><tr><th>Category</th><th>Items</th><th>Stock value</th></tr></thead><tbody>
    ${Object.keys(byCat).map(c=>`<tr><td>${c}</td><td>${byCat[c].count}</td><td>${money(byCat[c].value)}</td></tr>`).join('') || '<tr><td colspan="3" class="empty">No data</td></tr>'}
    </tbody></table>`, `${Object.keys(byCat).length} categor${Object.keys(byCat).length===1?'y':'ies'}`)}
    ${scopedCollapsePanel('reports-damaged', 'Damaged stock — by branch/location', DB.damaged, 'date', renderDamagedByBranch)}
    ${scopedCollapsePanel('reports-so', 'SO / Project completion', DB.soList, 'date', renderSOSummary)}
    ${collapsePanel('reports-fullstock', 'Full stock & cost sheet', renderMaterialsTable(), `${DB.materials.length} material${DB.materials.length===1?'':'s'}`)}`;
}
function renderDamagedByBranch(rows){
  rows = rows || DB.damaged;
  if(!rows.length) return `<div class="empty">No damaged stock logged.</div>`;
  const byLoc = {};
  rows.forEach(d=>{
    const loc = d.location||'Unspecified';
    if(!byLoc[loc]) byLoc[loc] = {count:0, qty:0, value:0};
    byLoc[loc].count++; byLoc[loc].qty += d.qty;
    const mat = materialById(d.materialId);
    byLoc[loc].value += d.qty * (mat? mat.price||0 : 0);
  });
  const summary = `<table><thead><tr><th>Branch / Location</th><th>Incidents</th><th>Qty damaged</th><th>Est. value lost</th></tr></thead><tbody>
    ${Object.keys(byLoc).map(l=>`<tr><td>${l}</td><td>${byLoc[l].count}</td><td>${byLoc[l].qty}</td><td>${money(byLoc[l].value)}</td></tr>`).join('')}
  </tbody></table>`;
  const detail = `<table style="margin-top:12px"><thead><tr><th>Date</th><th>Material</th><th>Qty</th><th>Branch/Location</th><th>Note</th></tr></thead><tbody>
    ${rows.map(d=>`<tr><td>${d.date}</td><td>${d.materialName}</td><td>${d.qty}</td><td>${d.location||'—'}</td><td>${d.note||'—'}</td></tr>`).join('')}
  </tbody></table>`;
  return summary + detail;
}
function renderSOSummary(rows){
  rows = rows || DB.soList;
  if(!rows.length) return `<div class="empty">No SOs in this range.</div>`;
  return `<table><thead><tr><th>SO</th><th>Status</th><th>Products</th><th>Completed products</th><th>Created</th><th>Completed</th></tr></thead><tbody>
    ${rows.map(so=>`<tr><td>${so.soNumber}</td><td>${soStatusBadge(so)}</td><td>${so.products.length}</td>
    <td>${so.products.filter(p=>p.status==='completed').length}</td><td>${so.date}</td><td>${so.completedDate||'—'}</td></tr>`).join('')}
  </tbody></table>`;
}

