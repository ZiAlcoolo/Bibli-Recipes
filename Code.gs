// Déploiement : Déployer > Nouveau déploiement > Application Web
// Exécuter en tant que : Moi | Accès : Tout le monde
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36';
const TAB = 'Recettes';
const HEAD = ['Titre', 'Source', 'Image', 'Portions', 'Préparation (min)', 'Cuisson (min)', 'Ingrédients', 'Étapes', 'Notes', 'Ajouté par', 'Date'];

function doGet(e) { return out_(route_(e.parameter)); }
function doPost(e) { return out_(route_(JSON.parse(e.postData.contents))); }
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function route_(p) {
  try {
    if (p.action === 'extract') return { ok: true, recipe: extract_(p.url) };
    if (p.action === 'list') return { ok: true, recipes: list_(p.sheet) };
    if (p.action === 'save') return { ok: true, row: save_(p.sheet, p.recipe, p.pseudo) };
    if (p.action === 'delete') return { ok: true, deleted: delete_(p.sheet, +p.row, p.title) };
    if (p.action === 'ping') { sheet_(p.sheet); return { ok: true }; }
    return { ok: false, error: 'Action inconnue' };
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

/* ---------- Extraction ---------- */
function extract_(url) {
  const res = UrlFetchApp.fetch(url, {
    headers: { 'User-Agent': UA, 'Accept-Language': 'fr-FR,fr;q=0.9,en;q=0.8' },
    muteHttpExceptions: true, followRedirects: true
  });
  if (res.getResponseCode() >= 400) throw new Error('Site inaccessible (code ' + res.getResponseCode() + '). Il bloque peut-être les robots.');
  const html = res.getContentText();
  const re = /<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  let m, r = null;
  while (!r && (m = re.exec(html))) {
    try { r = findRecipe_(JSON.parse(m[1].trim())); } catch (e) { /* bloc invalide, on passe au suivant */ }
  }
  if (!r) throw new Error("Aucune recette structurée trouvée sur cette page.");
  const steps = steps_(r.recipeInstructions);
  return {
    title: clean_(r.name),
    source: url,
    image: img_(r.image),
    servings: parseInt((String([].concat(r.recipeYield || '')[0]).match(/\d+/) || [0])[0], 10),
    prep: dur_(r.prepTime),
    cook: dur_(r.cookTime) || (r.prepTime ? 0 : dur_(r.totalTime)),
    ingredients: [].concat(r.recipeIngredient || []).map(clean_).filter(Boolean),
    steps: steps,
    notes: ''
  };
}

function findRecipe_(n) {
  if (Array.isArray(n)) { for (const x of n) { const f = findRecipe_(x); if (f) return f; } return null; }
  if (n && typeof n === 'object') {
    if ([].concat(n['@type'] || []).indexOf('Recipe') > -1) return n;
    if (n['@graph']) return findRecipe_(n['@graph']);
  }
  return null;
}

function steps_(x) {
  if (!x) return [];
  if (typeof x === 'string') return clean_(x.replace(/<\/?(br|p|li)[^>]*>/gi, '\n')).split(/\n+/).map(s => s.trim()).filter(Boolean);
  let out = [];
  [].concat(x).forEach(s => {
    if (typeof s === 'string') out.push(clean_(s));
    else if (s.itemListElement) out = out.concat(steps_(s.itemListElement));
    else if (s.text || s.name) out.push(clean_(s.text || s.name));
  });
  return out.filter(Boolean);
}

function img_(i) {
  if (!i) return '';
  if (Array.isArray(i)) i = i[0];
  return typeof i === 'string' ? i : (i.url || '');
}

function dur_(iso) {
  const m = String(iso || '').match(/P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?/);
  return m ? (+m[1] || 0) * 1440 + (+m[2] || 0) * 60 + (+m[3] || 0) : 0;
}

function clean_(s) {
  return String(s || '').replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;|&#0?39;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&#(\d+);/g, (_, c) => String.fromCharCode(c)).replace(/\s+/g, ' ').trim();
}

/* ---------- Google Sheet ---------- */
function sheet_(url) {
  const id = (String(url).match(/\/d\/([a-zA-Z0-9_-]+)/) || [])[1];
  if (!id) throw new Error('Lien Google Sheet invalide.');
  let ss;
  try { ss = SpreadsheetApp.openById(id); }
  catch (e) { throw new Error("Accès refusé : partage le Sheet en édition avec le compte propriétaire de l'app."); }
  let sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function save_(url, r, pseudo) {
  const sh = sheet_(url);
  sh.appendRow([r.title, r.source, r.image, r.servings || '', r.prep || '', r.cook || '',
    (r.ingredients || []).join('\n'), (r.steps || []).join('\n'), r.notes || '', pseudo || '', new Date()]);
  return sh.getLastRow();
}

function list_(url) {
  const v = sheet_(url).getDataRange().getValues().slice(1);
  const lines = s => String(s || '').split('\n').map(x => x.trim()).filter(Boolean);
  return v.map((r, i) => ({
    row: i + 2, title: r[0], source: r[1], image: r[2], servings: +r[3] || 0, prep: +r[4] || 0, cook: +r[5] || 0,
    ingredients: lines(r[6]), steps: lines(r[7]), notes: r[8], by: r[9]
  })).filter(x => x.title).reverse();
}

function delete_(url, row, title) {
  const sh = sheet_(url);
  if (row < 2 || String(sh.getRange(row, 1).getValue()) !== String(title)) throw new Error('La ligne a changé dans le Sheet, recharge ta liste.');
  sh.deleteRow(row);
  return true;
}
