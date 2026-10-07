// ============================================================
// PF1e Character Database — Inventory form v2
// ============================================================
// The inventory form has two branches (question "Who with?"):
//
// The world — question "Items and gold", one entry per line or comma:
//   +2 Potion of Cure Light Wounds                 bought, catalog price
//   +1 Potion of Dangerous Lion (250, 1 lb) [UE]   bought, stated price/weight
//   +1 Rope (free)  /  (900, free)                 got for free, value kept
//   -1 Potion of Cure Light Wounds                 used up, lost: no gold
//   -2 Potion of Cure Light Wounds (sell)          sold for half the catalog price
//   -2 Potion of Dangerous Lion (125)              sold for 125 each
//   -2 gold / +15 gold                             gold itself
//
// Other characters — "Participants" (checkboxes) and "Transfers",
// one transfer per line:  From -> To: 2 Potion of X, 11 gold
//
// A request is all or nothing: if any entry can't be applied (bad line,
// item the character doesn't have, unknown character), nothing is
// written and the request stays unprocessed — fix the answer in
// Inventory(raw) and run again. Warnings (e.g. bought for 0) don't stop it.
// ============================================================

// Transactions.source values written by this form
var REQUEST_SOURCES = ['inventory', 'inventory-free', 'transfer'];

var GOLD_NAMES = ['gold', 'gp', 'зм', 'золото', 'золота'];


// Returns { txns, warnings, errors }. txns is empty when errors isn't.
function parseInventoryRequest(row, I, inventory, charName, catalog, characters) {
  var timestamp = String(getField(row, I, 'timestamp'));
  var note = shortNote(getField(row, I, 'whatHappened'));
  var out = { txns: [], warnings: [], errors: [] };

  // Work on a copy: a failed request must leave the inventory untouched
  var inv = copyInventoryMap(inventory);
  function add(txn) {
    txn.timestamp = timestamp;
    out.txns.push(txn);
    updateInventoryMap(inv, txn);
  }

  var world = String(getField(row, I, 'worldItems') || '').trim();
  var transfers = String(getField(row, I, 'transfers') || '').trim();
  if (world) applyWorld(world, charName, inv, catalog, note, add, out);
  if (transfers) {
    applyTransfers(transfers, String(getField(row, I, 'participants') || ''),
                   inv, catalog, characters, note, add, out);
  }

  if (out.errors.length) { out.txns = []; return out; }
  for (var key in inv) inventory[key] = inv[key];
  return out;
}


// ============================================================
// THE WORLD
// ============================================================

function applyWorld(text, charName, inv, catalog, note, add, out) {
  var entries = splitEntries(text);
  for (var i = 0; i < entries.length; i++) {
    var e = parseWorldEntry(entries[i], catalog);
    if (e.error) { out.errors.push('"' + entries[i] + '": ' + e.error); continue; }

    var notes = note + (e.source ? (note ? ' ' : '') + '[' + e.source + ']' : '');

    if (e.isGold) {
      add({ source: 'inventory', character: charName, itemName: 'gold',
            quantity: e.sign * e.qty, cost: 1, weight: '', notes: notes });
      continue;
    }

    if (e.sign > 0) {
      var ref = catalog[e.name.toLowerCase()];
      var name = ref ? ref.name : e.name;
      var value = e.price !== null ? e.price : (ref && ref.cost !== '' ? Number(ref.cost) : null);
      var weight = e.weight !== null ? e.weight : (ref ? ref.weight : '');

      if (e.free) {
        add({ source: 'inventory-free', character: charName, itemName: name,
              quantity: e.qty, cost: value || 0, weight: weight, notes: notes });
        continue;
      }
      if (value === null) {
        out.warnings.push('"' + name + '" bought for 0: no price stated and not in catalog');
        value = 0;
      }
      add({ source: 'inventory', character: charName, itemName: name,
            quantity: e.qty, cost: value, weight: weight, notes: notes });
      var paid = roundGold(value * e.qty);
      if (paid > 0) {
        add({ source: 'inventory', character: charName, itemName: 'gold',
              quantity: -paid, cost: 1, weight: '',
              notes: 'Bought ' + (e.qty > 1 ? e.qty + ' ' : '') + name });
      }
      continue;
    }

    // Minus: used up, lost — or sold, if a price or "sell" is given
    var have = findOwned(inv, charName, e.name, catalog);
    if (!have.found || have.quantity < e.qty) {
      out.errors.push(notOwned(inv, charName, e.name, e.qty, have));
      continue;
    }
    add({ source: 'inventory', character: charName, itemName: have.name,
          quantity: -e.qty, cost: have.cost, weight: have.weight, notes: notes });

    var price = null, how = '';
    if (e.price !== null) {
      price = e.price; how = 'stated price';
    } else if (e.sell) {
      var ref2 = catalog[have.name.toLowerCase()];
      if (ref2 && Number(ref2.cost) > 0) { price = Number(ref2.cost) / 2; how = 'half catalog price'; }
      else if (Number(have.cost) > 0) { price = Number(have.cost) / 2; how = 'half its value'; }
      else out.warnings.push('"' + have.name + '" sold for 0: no price stated, none in catalog or inventory');
    }
    var got = price === null ? 0 : roundGold(price * e.qty);
    if (got > 0) {
      add({ source: 'inventory', character: charName, itemName: 'gold',
            quantity: got, cost: 1, weight: '',
            notes: 'Sold ' + (e.qty > 1 ? e.qty + ' ' : '') + have.name + ' (' + how + ')' });
    }
  }
}

// "+2 Name (price, 1 lb, sell|free) [source]" -> entry or { error }
function parseWorldEntry(raw, catalog) {
  var s = String(raw).trim();
  var signChar = s.charAt(0);
  var sign = signChar === '+' ? 1 : ('-−–—'.indexOf(signChar) !== -1 ? -1 : 0);
  if (!sign) return { error: 'line must start with + or -' };

  var rest = s.substring(1).trim();
  var e = { sign: sign, qty: 1, name: '', price: null, weight: null,
            sell: false, free: false, source: '', isGold: false };

  var q = rest.match(/^(\d+(?:[.,]\d+)?)\s*(?:[xх×]|шт\.?)?\s+(.+)$/i);
  if (q) { e.qty = toNumber(q[1]); rest = q[2]; }
  if (!(e.qty > 0)) return { error: 'quantity must be positive' };

  var src = rest.match(/^(.*?)\s*\[([^\]]*)\]\s*$/);
  if (src) { e.source = src[2].trim(); rest = src[1]; }

  // One "(…)" group of price / weight / sell / free at the end. A name
  // that itself ends in parentheses wins if it is in the catalog:
  // "Arrows (20)" is the item, "Arrows (20) (1)" costs 1.
  if (!catalog[rest.trim().toLowerCase()]) {
    var p = rest.match(/^(.*?)\s*\(([^()]*)\)\s*$/);
    var parsed = p ? parseDetails(p[2]) : null;
    if (parsed) {
      rest = p[1];
      e.price = parsed.price; e.weight = parsed.weight;
      e.sell = parsed.sell; e.free = parsed.free;
    }
  }

  e.name = rest.trim();
  if (!e.name) return { error: 'no item name' };
  e.isGold = GOLD_NAMES.indexOf(e.name.toLowerCase()) !== -1;
  return e;
}

// "250, 1 lb, free" -> { price, weight, sell, free }, or null if any
// part is not one of these (then the brackets belong to the name)
function parseDetails(inner) {
  var d = { price: null, weight: null, sell: false, free: false };
  var parts = inner.split(/[,;](?!\d)/);
  for (var i = 0; i < parts.length; i++) {
    var t = parts[i].trim();
    var w = t.match(/^(\d+(?:[.,]\d+)?|\d+\/\d+)\s*(?:lbs?|фнт|фунт\S*)\.?$/i);
    if (/^(?:sell|продаж[аи]|продал[аи]?)$/i.test(t)) d.sell = true;
    else if (/^(?:free|бесплатно)$/i.test(t)) d.free = true;
    else if (w) d.weight = toNumber(w[1]);
    else if (/^\d+(?:[.,]\d+)?$/.test(t)) d.price = toNumber(t);
    else return null;
  }
  return d;
}


// ============================================================
// OTHER CHARACTERS
// ============================================================

function applyTransfers(text, participantsAnswer, inv, catalog, characters, note, add, out) {
  var ticked = participantsAnswer.split(',').map(function(s) { return nameKey(s); })
                                 .filter(function(s) { return s; });
  var unticked = {};

  var lines = text.split(/\n|;/).map(function(s) { return s.trim(); })
                  .filter(function(s) { return s; });
  for (var l = 0; l < lines.length; l++) {
    var line = lines[l];
    var m = line.match(/^(.+?)\s*(?:->|→|=>|>)\s*(.+?)\s*:\s*(.+)$/);
    if (!m) { out.errors.push('"' + line + '": expected "From -> To: items"'); continue; }

    var from = findCharacter(characters, m[1]);
    var to = findCharacter(characters, m[2]);
    if (!from) { out.errors.push('"' + line + '": no character "' + m[1].trim() + '"'); continue; }
    if (!to) { out.errors.push('"' + line + '": no character "' + m[2].trim() + '"'); continue; }
    if (from === to) { out.errors.push('"' + line + '": giver and receiver are the same'); continue; }
    [from, to].forEach(function(c) {
      if (ticked.length && ticked.indexOf(nameKey(c)) === -1) unticked[c] = true;
    });

    var items = splitEntries(m[3]);
    for (var i = 0; i < items.length; i++) {
      var q = items[i].match(/^(\d+(?:[.,]\d+)?)\s*(?:[xх×]|шт\.?)?\s+(.+)$/i);
      var qty = q ? toNumber(q[1]) : 1;
      var name = (q ? q[2] : items[i]).trim();
      if (!(qty > 0) || !name) { out.errors.push('"' + items[i] + '": expected "2 Item name"'); continue; }

      var isGold = GOLD_NAMES.indexOf(name.toLowerCase()) !== -1;
      var have = isGold ? goldOf(inv, from) : findOwned(inv, from, name, catalog);
      if (!have.found || have.quantity < qty) {
        out.errors.push(isGold
          ? from + ' has ' + (have.quantity || 0) + ' gold, cannot give ' + qty
          : notOwned(inv, from, name, qty, have));
        continue;
      }
      var itemName = isGold ? 'gold' : have.name;
      var cost = isGold ? 1 : have.cost, weight = isGold ? '' : have.weight;
      add({ source: 'transfer', character: from, itemName: itemName, quantity: -qty,
            cost: cost, weight: weight, notes: 'To ' + to + (note ? ': ' + note : '') });
      add({ source: 'transfer', character: to, itemName: itemName, quantity: qty,
            cost: cost, weight: weight, notes: 'From ' + from + (note ? ': ' + note : '') });
    }
  }

  var missed = Object.keys(unticked);
  if (missed.length) out.warnings.push('not ticked in Participants: ' + missed.join(', '));
}


// ============================================================
// HELPERS
// ============================================================

// Entries are separated by new lines and by commas — but not by a comma
// inside (…) or […], or one followed by a digit (decimal: "0,5").
function splitEntries(text) {
  var out = [], cur = '', depth = 0;
  var s = String(text || '');
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    if (c === '(' || c === '[') depth++;
    if ((c === ')' || c === ']') && depth > 0) depth--;
    var isBreak = c === '\n' || (c === ',' && depth === 0 && !/\d/.test(s.charAt(i + 1)));
    if (isBreak) { out.push(cur); cur = ''; } else cur += c;
  }
  out.push(cur);
  return out.map(function(x) { return x.trim(); }).filter(function(x) { return x; });
}

function toNumber(s) {
  s = String(s).replace(',', '.');
  var f = s.match(/^(\d+)\/(\d+)$/);
  return f ? Number(f[1]) / Number(f[2]) : Number(s);
}

function shortNote(text) {
  var s = String(text || '').replace(/\s+/g, ' ').trim();
  return s.length > 80 ? s.substring(0, 79) + '…' : s;
}

function copyInventoryMap(map) {
  var copy = {};
  for (var k in map) {
    copy[k] = { name: map[k].name, quantity: map[k].quantity, cost: map[k].cost, weight: map[k].weight };
  }
  return copy;
}

// The item under the written name, or under its catalog spelling
// ("spell component pouch" -> "Spell component pouch").
function findOwned(inv, charName, name, catalog) {
  var names = [name];
  var ref = catalog[name.toLowerCase()];
  if (ref && ref.name !== name) names.push(ref.name);
  for (var i = 0; i < names.length; i++) {
    var key = charName.toLowerCase() + '|' + names[i].toLowerCase();
    if (key in inv && inv[key].quantity > 0) {
      return { found: true, name: originalName(inv, key, names[i]), quantity: inv[key].quantity,
               cost: inv[key].cost || 0, weight: inv[key].weight || '' };
    }
  }
  return { found: false, quantity: 0 };
}

// Map keys are lower case; the item keeps the spelling it was first written with
function originalName(inv, key, written) { return inv[key].name || written; }

function goldOf(inv, charName) {
  var key = charName.toLowerCase() + '|gold';
  var q = key in inv ? Number(inv[key].quantity) || 0 : 0;
  return { found: q > 0, quantity: roundGold(q) };
}

function notOwned(inv, charName, name, qty, have) {
  var msg = charName + ' has ' + (have.found ? have.quantity : 'no') + ' "' + name + '", needs ' + qty;
  var like = similarItems(inv, charName, name);
  return like.length ? msg + '; similar: ' + like.join(', ') : msg;
}

// Names compared without case, spaces, apostrophes and hyphens:
// "ке'цаль" finds "Ке`цаль".
function nameKey(name) {
  return String(name).toLowerCase().replace(/[\s`'’ʼ-]/g, '');
}

function findCharacter(characters, name) {
  var key = nameKey(name);
  for (var i = 0; i < characters.length; i++) {
    if (nameKey(characters[i]) === key) return characters[i];
  }
  return '';
}

function loadCharacterNames(ss) {
  var sheet = ss.getSheetByName('Database');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var col = data[3].map(function(h) { return String(h).trim(); }).indexOf('char');
  var names = [];
  for (var r = 4; r < data.length; r++) {
    var v = String(data[r][col]).trim();
    if (v) names.push(v);
  }
  return names;
}

// Items the character has whose name is close to the written one
function similarItems(inv, charName, itemName) {
  var prefix = charName.toLowerCase() + '|';
  var want = itemName.toLowerCase();
  var out = [];
  for (var key in inv) {
    if (key.indexOf(prefix) !== 0 || !(inv[key].quantity > 0)) continue;
    var have = key.substring(prefix.length);
    if (have === 'gold' || have === want) continue;
    if (have.indexOf(want) !== -1 || want.indexOf(have) !== -1 ||
        editDistance(have, want) <= Math.max(2, Math.floor(want.length / 5))) {
      out.push('"' + have + '"');
    }
  }
  return out.slice(0, 3);
}

function editDistance(a, b) {
  var prev = [];
  for (var j = 0; j <= b.length; j++) prev.push(j);
  for (var i = 1; i <= a.length; i++) {
    var cur = [i];
    for (var j = 1; j <= b.length; j++) {
      cur.push(Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)));
    }
    prev = cur;
  }
  return prev[b.length];
}
