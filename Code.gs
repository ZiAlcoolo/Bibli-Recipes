// Déploiement : Déployer > Nouveau déploiement > Application Web
// Exécuter en tant que : Moi | Accès : Tout le monde
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36';
const TAB = 'Recettes';
const HEAD = ['Titre', 'Source', 'Image', 'Portions', 'Préparation (min)', 'Cuisson (min)', 'Ingrédients', 'Étapes', 'Notes', 'Ajouté par', 'Date', 'ID'];

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
    if (p.action === 'groups') return { ok: true, groups: groups_(p.sheet) };
    if (p.action === 'group_save') return { ok: true, id: group_save_(p.sheet, p.group) };
    if (p.action === 'group_delete') return { ok: true, deleted: group_delete_(p.sheet, p.id) };
    if (p.action === 'share') return { ok: true, code: share_(p.sheet, p) };
    if (p.action === 'shared') return { ok: true, shared: shared_(p.code) };
    if (p.action === 'import') return { ok: true, result: import_(p.sheet, p.group, p.pseudo) };
    if (p.action === 'delete') return { ok: true, deleted: delete_(p.sheet, +p.row, p.title) };
    if (p.action === 'ping') { sheet_(p.sheet); return { ok: true }; }
    return { ok: false, error: 'Action inconnue' };
  } catch (err) {
    return { ok: false, error: String(err.message || err), manual: !!err.manual };
  }
}

/* ---------- Extraction ---------- */
function extract_(url) {
  if (/^https?:\/\/([^\/]*\.)?(instagram\.com|tiktok\.com)\//i.test(url)) return social_(url);
  const res = UrlFetchApp.fetch(url, {
    headers: { 'User-Agent': UA, 'Accept-Language': 'fr-FR,fr;q=0.9,en;q=0.8' },
    muteHttpExceptions: true, followRedirects: true
  });
  if (res.getResponseCode() >= 400) fail_('Site inaccessible (code ' + res.getResponseCode() + '). Il bloque peut-être les robots.');
  const html = res.getContentText();
  const re = /<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  let m, r = null;
  while (!r && (m = re.exec(html))) {
    try { r = findRecipe_(JSON.parse(m[1].trim())); } catch (e) { /* bloc invalide, on passe au suivant */ }
  }
  if (!r) fail_("Aucune recette structurée trouvée sur cette page.");
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
  } else if (sh.getRange(1, 12).getValue() !== 'ID') {
    sh.getRange(1, 12).setValue('ID').setFontWeight('bold');
  }
  return sh;
}

function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

function row_(r, pseudo, id) {
  return [r.title, r.source || '', cacheImage_(r.image || ''), r.servings || '', r.prep || '', r.cook || '',
    (r.ingredients || []).join('\n'), (r.steps || []).join('\n'), r.notes || '', pseudo || '', new Date(), id || newId_()];
}

function save_(url, r, pseudo) {
  const sh = sheet_(url);
  sh.appendRow(row_(r, pseudo));
  return sh.getLastRow();
}

function list_(url) {
  const sh = sheet_(url);
  const v = sh.getDataRange().getValues().slice(1);
  let missing = false;
  v.forEach(r => { if (r[0] && !r[11]) { r[11] = newId_(); missing = true; } });
  if (missing) sh.getRange(2, 12, v.length, 1).setValues(v.map(r => [r[11] || '']));
  const lines = s => String(s || '').split('\n').map(x => x.trim()).filter(Boolean);
  return v.map((r, i) => ({
    row: i + 2, id: r[11], title: r[0], source: r[1], image: r[2], servings: +r[3] || 0, prep: +r[4] || 0, cook: +r[5] || 0,
    ingredients: lines(r[6]), steps: lines(r[7]), notes: r[8], by: r[9]
  })).filter(x => x.title).reverse();
}

function delete_(url, row, title) {
  const sh = sheet_(url);
  if (row < 2 || String(sh.getRange(row, 1).getValue()) !== String(title)) throw new Error('La ligne a changé dans le Sheet, recharge ta liste.');
  sh.deleteRow(row);
  return true;
}

/* ---------- Instagram / TikTok (légende via oEmbed public, sans clé) ---------- */
function fail_(msg) { const e = new Error(msg); e.manual = true; throw e; }

function social_(url) {
  return /tiktok\.com/i.test(url) ? tiktok_(resolve_(url)) : instagram_(url);
}

// Instagram : page d'intégration publique (l'oEmbed ne renvoie plus la légende).
// Source principale : le JSON « contextJSON » de la page (légende complète + image) ; repli : le bloc HTML Caption.
function instagram_(url) {
  const m = url.match(/\/(reel|reels|p|tv)\/([A-Za-z0-9_-]+)/);
  if (!m) fail_('Lien Instagram non reconnu.');
  const res = UrlFetchApp.fetch('https://www.instagram.com/' + (m[1] === 'reels' ? 'reel' : m[1]) + '/' + m[2] + '/embed/captioned/', {
    headers: { 'User-Agent': UA, 'Accept-Language': 'fr-FR,fr;q=0.9,en;q=0.8' }, muteHttpExceptions: true
  });
  if (res.getResponseCode() === 429) fail_('Instagram limite les requêtes : réessaie dans quelques minutes.');
  if (res.getResponseCode() >= 400) fail_('Post inaccessible (supprimé ou service indisponible).');
  const h = res.getContentText();
  let caption = '', image = '', author = '', media = null;
  const x = h.match(/"contextJSON":"((?:[^"\\]|\\.)*)"/);
  if (x) { try { media = deep_(JSON.parse(JSON.parse('"' + x[1] + '"')), 'shortcode_media'); } catch (e) { /* repli HTML */ } }
  if (media) {
    const e = media.edge_media_to_caption && media.edge_media_to_caption.edges && media.edge_media_to_caption.edges[0];
    caption = e ? e.node.text : '';
    image = media.display_url || media.thumbnail_src || ((media.display_resources || []).slice(-1)[0] || {}).src || '';
    author = (media.owner && media.owner.username) || '';
  }
  if (!caption) {
    const c = h.match(/<div class="Caption">([\s\S]*?)<div class="CaptionComments"/);
    if (c) caption = text_(c[1].replace(/<a class="CaptionUsername"[\s\S]*?<\/a>/, ''));
  }
  if (!image) { const i = h.match(/class="EmbeddedMediaImage"[^>]*src="([^"]+)"/); if (i) image = i[1].replace(/&amp;/g, '&'); }
  if (!author) { const u = h.match(/<a class="CaptionUsername"[^>]*>([\s\S]*?)<\/a>/); if (u) author = text_(u[1]); }
  if (!caption.trim()) {
    if (!media && /"contextJSON":null/.test(h)) fail_('Ce post est privé ou indisponible : seuls les posts publics peuvent être lus. Colle la légende à la main.');
    fail_('Aucune légende trouvée sur ce post.');
  }
  return { caption: caption, author: author, image: image, source: url.split('?')[0] };
}

function deep_(o, key) {
  if (o && typeof o === 'object') {
    if (key in o && o[key] != null) return o[key];
    for (const k in o) { const r = deep_(o[k], key); if (r != null) return r; }
  }
  return null;
}

// TikTok : l'oEmbed ne renvoie plus la légende (en attente du résultat des tests)
function tiktok_(url) {
  const res = UrlFetchApp.fetch('https://www.tiktok.com/oembed?url=' + encodeURIComponent(url), { headers: { 'User-Agent': UA }, muteHttpExceptions: true });
  if (res.getResponseCode() >= 400) fail_('Vidéo inaccessible (privée, supprimée ou service indisponible).');
  const d = JSON.parse(res.getContentText());
  if (!(d.title || '').trim()) fail_('Aucune légende récupérable pour cette vidéo.');
  return { caption: d.title, author: d.author_name || '', image: d.thumbnail_url || '', source: url };
}

function resolve_(u) {
  for (let i = 0; i < 4 && /\/\/(vm|vt)\.tiktok\.com|tiktok\.com\/t\//i.test(u); i++) {
    const h = UrlFetchApp.fetch(u, { headers: { 'User-Agent': UA }, followRedirects: false, muteHttpExceptions: true }).getHeaders();
    const l = h['Location'] || h['location'];
    if (!l) break;
    u = l.split('?')[0];
  }
  return u;
}

/* ---------- Groupes de recettes (onglet « Groupes » du Sheet) ---------- */
const GHEAD = ['ID', 'Nom', 'Description', 'Recettes (IDs)', 'Date'];

function ss_(url) {
  const id = (String(url).match(/\/d\/([a-zA-Z0-9_-]+)/) || [])[1];
  if (!id) throw new Error('Lien Google Sheet invalide.');
  try { return SpreadsheetApp.openById(id); }
  catch (e) { throw new Error("Accès refusé : partage le Sheet en édition avec le compte propriétaire de l'app."); }
}

function gtab_(url) {
  const ss = ss_(url);
  let sh = ss.getSheetByName('Groupes');
  if (!sh) {
    sh = ss.insertSheet('Groupes');
    sh.appendRow(GHEAD);
    sh.getRange(1, 1, 1, GHEAD.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function groups_(url) {
  return gtab_(url).getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({
    id: r[0], name: r[1], description: r[2], ids: String(r[3] || '').split('\n').map(x => x.trim()).filter(Boolean)
  }));
}

function grow_(sh, id) {
  const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < col.length; i++) if (col[i][0] === id) return i + 1;
  return 0;
}

function group_save_(url, g) {
  if (!g || !String(g.name || '').trim()) throw new Error('Le groupe doit avoir un nom.');
  const sh = gtab_(url), ids = (g.ids || []).join('\n');
  const row = g.id ? grow_(sh, g.id) : 0;
  if (row) { sh.getRange(row, 2, 1, 3).setValues([[g.name, g.description || '', ids]]); return g.id; }
  const id = newId_();
  sh.appendRow([id, g.name, g.description || '', ids, new Date()]);
  return id;
}

function group_delete_(url, id) {
  const sh = gtab_(url), row = grow_(sh, id);
  if (row) sh.deleteRow(row);
  return !!row;
}

/* ---------- Partage par lien (instantané stocké dans un Sheet de l'administrateur) ---------- */
const CHUNK = 45000;

function store_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHARE_SS');
  if (id) { try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) { /* supprimé : on recrée */ } }
  const ss = SpreadsheetApp.create('Recettes — partages (ne pas supprimer)');
  props.setProperty('SHARE_SS', ss.getId());
  const sh = ss.getSheets()[0];
  sh.appendRow(['Code', 'Type', 'Par', 'Date', 'Données (JSON, par blocs)']);
  return sh;
}

function strip_(r) {
  return { title: r.title, source: r.source || '', image: r.image || '', servings: +r.servings || 0, prep: +r.prep || 0, cook: +r.cook || 0,
    ingredients: r.ingredients || [], steps: r.steps || [], notes: r.notes || '' };
}

function share_(url, p) {
  let data;
  if (p.type === 'group') {
    const g = groups_(url).filter(x => x.id === p.id)[0];
    if (!g) throw new Error('Groupe introuvable.');
    data = { name: g.name, description: g.description, recipes: list_(url).filter(r => g.ids.indexOf(r.id) > -1).map(strip_) };
  } else {
    if (!p.recipe || !p.recipe.title) throw new Error('Recette invalide.');
    data = strip_(p.recipe);
    data.image = cacheImage_(data.image);
  }
  const json = JSON.stringify(data);
  if (json.length > CHUNK * 8) throw new Error('Trop volumineux pour être partagé.');
  const chunks = json.match(new RegExp('[\\s\\S]{1,' + CHUNK + '}', 'g'));
  const code = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { store_().appendRow([code, p.type === 'group' ? 'group' : 'recipe', p.pseudo || '', new Date()].concat(chunks)); }
  finally { lock.releaseLock(); }
  return code;
}

function shared_(code) {
  if (!/^[a-f0-9]{12}$/.test(String(code))) throw new Error('Lien de partage invalide.');
  const sh = store_();
  const f = sh.createTextFinder(code).matchEntireCell(true).findNext();
  if (!f || f.getColumn() !== 1) throw new Error('Partage introuvable (lien expiré ou supprimé).');
  const row = sh.getRange(f.getRow(), 1, 1, sh.getLastColumn()).getValues()[0];
  return { type: row[1], by: row[2], data: JSON.parse(row.slice(4).join('')) };
}

// Copie un groupe partagé (recettes + groupe) dans le Sheet de l'utilisateur, en une seule fois
function import_(url, grp, pseudo) {
  if (!grp || !grp.name) throw new Error('Groupe invalide.');
  const sh = sheet_(url);
  const rows = (grp.recipes || []).map(r => row_(strip_(r), pseudo));
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
  const id = group_save_(url, { name: grp.name, description: grp.description, ids: rows.map(r => r[11]) });
  return { id: id, count: rows.length };
}

/* ---------- Images : copie permanente dans Drive ---------- */
// Les URLs d'images Instagram / TikTok sont signées et expirent au bout de quelques jours.
// On les copie dans un dossier Drive (lisible par lien) ; si Drive n'est pas autorisé, on garde l'URL d'origine.
const SOCIAL_CDN = /(cdninstagram\.com|fbcdn\.net|tiktokcdn[\w-]*\.com)/i;

function cacheImage_(url) {
  if (!url || !SOCIAL_CDN.test(url)) return url;
  try {
    const res = UrlFetchApp.fetch(url, { headers: { 'User-Agent': UA }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return url;
    const blob = res.getBlob();
    if (!/^image\//.test(blob.getContentType())) return url;
    const f = imgFolder_().createFile(blob.setName('recette-' + newId_() + '.jpg'));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1000';
  } catch (e) { return url; }
}

function imgFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('IMG_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* supprimé : on recrée */ } }
  const f = DriveApp.createFolder('Recettes – images');
  props.setProperty('IMG_FOLDER', f.getId());
  return f;
}

// À exécuter UNE FOIS depuis l'éditeur pour autoriser Drive, puis redéployer une nouvelle version
function autoriserDrive() { Logger.log('Drive autorisé : ' + DriveApp.getRootFolder().getName()); }
