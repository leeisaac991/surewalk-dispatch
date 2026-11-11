/* Copied unchanged from apps-script/Places.gs (place names, zones, cart rules). Refresh with: ./sync-places.sh */
/* Browser shims: no geocoder here, so only exact place-name matching is used. */
function bigGet_() { return null; }
function bigSet_() {}

/**
 * Place names: short names, nicknames and common misspellings -> one canonical place.
 * Built from ~4,000 pickups/drop-offs in the live sheet (Sep 30 – Oct 17, 2026).
 * Used for grouping a crew's riders by pickup, the possible-duplicate check, and matching
 * driver messages like "picked kins" / "at sanjac" to a ride.
 *
 * Matching ignores case, punctuation, spaces, apostrophes, zip codes and filler words
 * ("the", "hall", "dorm", "residence", "apartments", "austin", "st"...), so only real
 * nickname/typo variants need listing. Keep look-alike places SEPARATE: e.g. San Jac Garage
 * is not San Jacinto, Littlefield Fountain/House are not Littlefield, Jester West is not Jester (Jester East is).
 */
const FILE_VERSION_PLACES = '2026-11-07a';

const PLACES = {
  // Dorms
  'Jester':               ['jester', 'jester center', 'jester circle', 'jester east', 'jester e', 'east jester', 'jcl',
                          'j2', 'j2 dining', 'j2 dining hall', 'jester 2', 'jester j2', 'j1', 'j1 dining', 'j1 dining hall', // J1 / J2: Jester's dining halls
                          'jeast', 'j east', 'j-east'],
  'Jester West':          ['jwest', 'j west', 'j-west', 'jester west'],
  'Kinsolving':           ['kins', 'kinsolving', 'kiins', 'kinsoving', 'kinsolvinfg'],
  'San Jacinto':          ['san jac', 'sanjac', 'san jacinto', 'sacjac', 'san jac dorms', 'san jacinto dorms', 'sjc', 'cypress cafe', 'cypress', 'cypress bend cafe'],
  'Moore-Hill':           ['moore hill', 'moorehill', 'morehill', 'more hill', 'moore'],
  'Somerset Condos':      ['somerset', 'somerset condos', 'somerset condo', 'somerset condominiums', 'somerset condominium'],
  // 1903 Robbins Pl (West Campus), typed many ways: "Robinson PI", with or without the house number.
  'Robbins Place':        ['robbins pl', 'robbins place', 'robinson pl', 'robinson pi', 'robbins pi', 'robinson place',
                           '1903 robbins pl', '1903 robbins place', '1903 robinson pl', '1903 robinson pi', '903 robinson pi', '903 robinson pl', '903 robbins pl'],
  'Stoneleigh Condominiums': ['stoneleigh', 'stoneleigh condominiums', 'stoneleigh condos', 'stoneleigh condo',
                           'stoneleigh condominiums on 2409 leon st', 'stoneleigh condominiums 2409 leon st'],
  'Phi Gamma Delta':      ['phi gamma delta', 'phi gam', 'phigam', 'fiji', 'fiji house', 'phi gamma delta house', 'phi gamma', 'phi gam house'],
  'Avon at 22nd':         ['avon', 'avon at 22nd', 'avon at 22nd st', 'avon at 22nd street', 'avon at 22', 'avon@22', 'avon@22nd',
                           'avon 22', 'avon 22nd', 'avon 22nd st', 'avon on 22nd', 'avon on 22nd st', 'avon apartments'],
  'Littlefield':          ['littlefield', 'little field', 'littlefield dorms'],
  'Duren':                ['duren', 'durin', 'duran', 'durren', 'duren hall', 'durin hall', 'duren dorm', 'duren dorms', 'durin dorms'],
  'Carothers':            ['carothers', 'carother', 'carothers dormitory'],
  'Roberts':              ['roberts', 'robert'],
  'Blanton':              ['blanton', 'blanton residend'],
  'Prather':              ['prather'],
  'Creekside':            ['creekside', 'creek side'],
  'Brackenridge':         ['brackenridge', 'breckenridge'],
  'Andrews':              ['andrews'],
  // Private dorms / apartments
  // The Target on Guadalupe is in the Dobie building: shown and written to the sheet as "Target (Dobie)" (PLACE_REWRITES).
  'Target (Dobie)':       ['target', 'target (dobie)', 'target dobie', 'dobie target', 'target at dobie', 'the target', 'target on guad', 'target on guadalupe', 'ut target', 'target store'],
  'Dobie Twenty21':       ['dobie', 'dobie 21', 'dobie21', 'dobie twenty21', 'dobie twenty 21', 'dobie twenty one', 'dobie center'],
  'Castilian':            ['castilian', 'castillian', 'castillion'],
  '2400 Nueces':          ['2400 nueces', '2400 neuces', '2400 nucess', '2400 nuecues', '2400 nueces complex'],
  '26 West':              ['26 west', '26west'],
  'Legacy on Rio':        ['legacy on rio', 'legacy', 'legacy rio'],
  'The Rambler':          ['rambler'],
  'Skyloft':              ['skyloft', 'sky loft'],
  'The Standard':         ['standard'],
  'The Nine at West Campus': ['nine at west campus', 'nine at west', 'nine west campus', '9 west campus', 'the nine'],
  'The Nine at Rio':      ['nine at rio'],
  'Villas on Rio':        ['villas on rio'],
  'Villas on Guadalupe':  ['villas on guad', 'villas on guadalupe', 'villas on gaud'],
  'Villas on Nueces':     ['villas on nueces', 'villas om nueces'],
  'Villas on 26th':       ['villas on 26th', 'villas on 26', 'villas on w 26th', 'villas on west 26th', 'villas 26th'], // 800 W 26th St
  'Venue on Guadalupe':   ['venue on guad', 'venue on guadalupe'],
  'Hilltop':              ['hilltop', 'hill top'],
  'Lark':                 ['lark'],
  'Rise at West Campus':  ['rise', 'rise at west campus', 'rise at west', 'the rise'],
  'The Sinclair':         ['sinclair', 'sinclaire', 'sinclair aps', 'sinclair 2710 nueces'],
  'Axis West':            ['axis west', 'axis west campus', 'axis west campus 2505 longview'],
  '24 Longview':          ['24 longview', '24longview'],
  // 2616 Salado St = "Texan West Campus" / "The Texan" (riders also say "Texan 26"; a rider objected to that label).
  'Texan West Campus':    ['texan west campus', 'texan west', 'texan wc', 'the texan', 'texan apartments', 'the texan west campus',
                           'texan 26', 'texan26', '2616 salado', 'texan 26 2616 salado', 'texan 2616 salado'],
  'Crest at Pearl':       ['crest at pearl'],
  'Union on 24th':        ['union on 24th', 'union on 24', 'union 24th', 'uniton 24th', 'union of 24'],
  'Union on San Antonio': ['union on san antonio', 'union san antonio'],
  'The Block on 23rd':    ['block on 23rd', 'block on 23', 'block 23rd'],
  'The Block on 25th East': ['block on 25th east', 'block on 25 e', 'block on 25e'],
  'The Block on 25th West': ['block on 25th west', 'block on 25th w', 'block 25th w'],
  'The Block on 28th':    ['block on 28th'],
  'The Block on Leon':    ['block on leon', 'block at leon', 'bock on leon'],
  'The Block on Pearl North': ['block on pearl north', 'block on pearl n', 'block pearl north', 'north block on pearl'],
  'The Block on Pearl South': ['block on pearl south', 'block on pearl s', 'block pearl south', 'south block on pearl'],
  'The Block on Pearl':   ['block on pearl', 'block at pearl', 'block pearl'],
  'Signature 1909':       ['signature 1909'],
  'Inspire on 22nd':      ['inspire on 22nd', 'inspire on 2nnd'],
  '21 Pearl':             ['21 pearl'],
  '21 Rio':               ['21 rio'],
  'Yugo Rio':             ['yugo rio', 'yugo austin rio', 'yugo on rio'],
  'Yugo Waterloo':        ['yugo waterloo', 'yugo austin waterloo', 'waterloo yugo'],
  'Yugo Corner':          ['yugo corner', 'yugo austin corner', 'yugo the corner'],
  'Yugo Nueces':          ['yugo nueces', 'yugo austin nueces'],
  'Quarters Nueces':      ['quarters nueces', 'quarters on nueces', 'quarters at nueces', 'quarter nueces', 'quarters nueces house'],
  'Regents West at 24th': ['regents west at 24th', 'regents west 24th', 'regents west on 24th', 'regents on 24th', 'regents at 24th'],
  'Regents West at 26th': ['regents west at 26th', 'regents west on 26th', 'regents west on 26', 'regents at 26th',
                           'regents w 26th', 'regents on 26th', 'regents west in 26', 'regents west 26th'],
  'GrandMarc':            ['grandmarc', 'grand marc', 'grandmark'],
  'Noble 2500':           ['noble 2500'],
  'Sparq on Rio':         ['sparq on rio'],
  'Tradition on Rio':     ['tradition on rio', 'traditions on rio'],
  'Tanglewood North':     ['tanglewood north'],
  'Moontower':            ['moontower', 'moon tower'],
  '1883 Cameron':         ['1883 cameron', '1883 at cameron', '1883 cameron house'],
  '1883 Montgomery':      ['1883 montgomery', '1883 at montgomery'],
  'Twenty Two 15':        ['twenty two 15', 'twentytwo 15', '2215'],
  'Duval Villas':         ['duval villas', 'duval villa'],
  'Westridge':            ['westridge', 'west ridge', 'westridge condos'],
  'Miller Square & Spanish Trails': ['miller square and spanish trails', 'miller square and spanish trail'],
  // Campus buildings & spots
  // McCombs (GSB / CBA) is right next to PCL: dispatchers treat it as the same pickup spot.
  'PCL':                  ['pcl', 'pcl library', 'perry castaneda library', 'perry castaneda',
                           'mccombs', 'mcombs', 'mccomb', 'mccombs school of business', 'mccombs business school', 'mccombs building',
                           'mccombs speedway', 'mccombs building pcl', 'mccombs pcl', 'pcl mccombs', 'gsb', 'cba',
                           'graduate school of business', 'graduate school of business mccombs next to dobie',
                           'mccombs school of business 2110 speedway austin tx 78705'],
  'Gregory Gym':          ['greg', 'greg gym', 'gregory gym', 'gregory'],
  'Rec Center':           ['rec', 'rec center', 'recsports', 'rec sports', 'ut rec center'],
  'Texas Union':          ['union', 'texas union'],
  'Butler School of Music': ['butler', 'butler school of music', 'butler school music', 'butler music', 'bsom', 'b.s.o.m', 'butler som', 'school of music'],
  'Moody CMA':            ['cma', 'moody cma'],
  // Just "Moody" could be the Moody Center (arena) or DMC (Dealey Center, Moody College): never guessed — flagged.
  'Moody (Center or DMC?)': ['moody'],
  'Winship':              ['winship', 'winship drama'],
  'Art Building':         ['art', 'art building'],
  'Welch':                ['welch'],
  'Burdine':              ['burdine'],
  'Sutton':               ['sutton'],
  'Belmont':              ['belmont', 'bellmont'],
  'Littlefield Fountain': ['littlefield fountain', 'little field fountain'],
  'Littlefield House':    ['littlefield house', 'little field house'],
  'UT Administration':    ['uta', 'ut administration'],
  'DKR Stadium':          ['dkr', 'dkr stadium'],
  'The Bubble':           ['bubble'],
  // Garages (kept separate from the dorms they're named after)
  'Brazos Garage':        ['brazos', 'brazos garage', 'brg'],
  'Manor Garage':         ['manor', 'manor garage', 'manor gargae'],
  'San Jacinto Garage':   ['san jac garage', 'san jacinto garage'],
  'Sunchase Condos':      ['sunchase', 'sunchase condos', 'sunchase condo', 'sunchase condominiums', '3115 helms', '3115 helms st'], // north of 29th
  '21 Rio':               ['21 rio', '21 rio apartments', 'twenty one rio'],
  'Orange Tree Condos':   ['orange tree', 'orange tree condos', 'orange tree condominiums'],
  'St Thomas Condos':     ['st thomas', 'st thomas condo', 'st thomas condos', 'saint thomas condos'],
  'Overlook Condos':      ['overlook', 'overlook condos', 'overlook condos poplar st'],
  'Lenox Condos':         ['lenox', 'lenox condos'],
  'Lonestar Lofts':       ['lonestar lofts', 'lone star lofts'],
  'Miller Square':        ['miller square', 'miller square and spanish trails', 'miller square & spanish trails', 'spanish trails'],
  'Whitaker Fields':      ['whitaker', 'whitaker fields', 'whitaker field', 'whittaker', 'whittaker fields', 'whitaker intramural fields',
                          'witaker', 'witaker fields', 'whitaker feilds', 'wright witaker feilds', 'witaker feilds',
                          'alan wright fields', 'alan wright field', 'alan wright', 'wright fields', 'alan wright intramural fields'], // drop-off only (its new name is Alan Wright Fields; people know Whitaker)
  'Scottish Rite Dormitory': ['srd', 'scottish rite', 'scottish rite dorm', 'scottish rite dormitory', 'scottish rite residence hall', 's r d'], // 210 W 27th
  'Guadalupe Garage':     ['guadalupe garage', 'gug', 'guad garage', 'guadalupe parking garage', 'gug garage'],
  'Health Center Garage': ['health center garage'],
  'East Campus Garage':   ['east campus garage'],
};

const PLACE_FILLER = new Set(['the', 'hall', 'dorm', 'dormitory', 'residence', 'apartments', 'apartment', 'apts',
  'apt', 'building', 'bldg', 'austin', 'tx', 'texas', 'st', 'street', 'unitedstates', 'usa', 'united', 'states']);

// "The Moore-Hill Residence Hall, Austin TX 78705" -> "moorehill"
function placeNorm_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '') // "Castañeda" = "Castaneda"
    .replace(/[’'`]/g, '').replace(/&/g, ' and ').replace(/\b787\d\d\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ').split(' ')
    .filter(w => w && !PLACE_FILLER.has(w)).join('');
}

let PLACE_INDEX_ = null;
function placeIndex_() {
  if (!PLACE_INDEX_) {
    PLACE_INDEX_ = {};
    Object.keys(PLACES).forEach(name => {
      [name].concat(PLACES[name]).forEach(v => { PLACE_INDEX_[placeNorm_(v)] = name; });
    });
  }
  return PLACE_INDEX_;
}

/** Canonical place name ("kins" -> "Kinsolving"), or null if it isn't a known place. */
// Riders add notes in parentheses: "The Nine at West Campus (on Leon and 26th)" is The Nine at West Campus.
const dropNote_ = s => String(s || '').replace(/\([^)]*\)?/g, ' ').replace(/\s+/g, ' ').trim();

function canonPlace_(s) {
  const idx = placeIndex_();
  return idx[placeNorm_(s)] || idx[placeNorm_(dropNote_(s))] || fuzzyPlace_(placeNorm_(dropNote_(s) || s)) || null;
}
// Misspelled place names: "Dell Senton Medical Center", "Jester weat", "kinsolvng". Long names may be 2 letters off,
// medium ones 1; short names and anything with numbers (addresses) are never guessed; a tie between places = no guess.
let FUZZY_CACHE_ = {};
function fuzzyPlace_(n) {
  if (!n || /\d/.test(n) || n.length < 6) return null;
  if (n in FUZZY_CACHE_) return FUZZY_CACHE_[n];
  const max = n.length >= 10 ? 2 : 1, idx = placeIndex_();
  let best = null, bestD = max + 1, tie = false;
  Object.keys(idx).forEach(k => {
    if (k.length < 6 || k[0] !== n[0] || /\d/.test(k) || Math.abs(k.length - n.length) > max) return; // typos keep the first letter ("rise" isn't "nine")
    const d = editDist_(n, k, max);
    if (d < bestD) { bestD = d; best = idx[k]; tie = false; } else if (d === bestD && idx[k] !== best) tie = true;
  });
  return (FUZZY_CACHE_[n] = best && !tie ? best : null);
}
// Edit distance with swaps ("teh" = "the"), giving up early once it's over `max`.
function editDist_(a, b, max) {
  let prev2 = null, prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    let rowMin = i;
    for (let j = 1; j <= b.length; j++) {
      let v = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      if (prev2 && i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) v = Math.min(v, prev2[j - 2] + 1);
      cur.push(v); if (v < rowMin) rowMin = v;
    }
    if (rowMin > max) return max + 1;
    prev2 = prev; prev = cur;
  }
  return prev[b.length];
}

/** Stable comparison key: the canonical name if known, else the normalized text. */
function placeKey_(s) {
  const c = canonPlace_(s);
  return c ? c.toLowerCase() : placeNorm_(dropNote_(s) || s);
}

/*
 * Drop-off -> Location zone (column M), learned from the sheet: Location follows the DROP-OFF (it matched
 * the drop-off 95.6% of the time vs 58% for the pickup). Only places with 3+ past rides where 80%+ agreed are
 * listed (190 places, 91% of past rides; 96.8% agreement with what dispatchers entered). Anything else is left
 * for the dispatcher. Keys are placeKey_() of the drop-off. Add or fix entries freely.
 */
const ZONE_OF = {
  // On
  "2400 nueces": 'On', "andrews": 'On', "blanton": 'On', "brackenridge": 'On', "brazos garage": 'On',
  "butler school of music": 'On', "carothers": 'On', "cpe": 'On', "creekside": 'On', "dkr stadium": 'On',
  "dmc": 'On', "dobie twenty21": 'On', "target (dobie)": 'On', "duren": 'On', "ecg": 'On', "ecga": 'On', "eer": 'On', "etc": 'On',
  "gregory gym": 'On', "hcg": 'On', "health center garage": 'On', "jester": 'On',
  "jester west": 'On', "kinsolving": 'On', "littlefield": 'On', "lot118": 'On', "lot37": 'On', "manor garage": 'On',
  "mccombs": 'On', "moody cma": 'On', "moody (center or dmc?)": 'On', "moore-hill": 'On', "pcl": 'On', "prather": 'On', "rec center": 'On',
  "roberts": 'On', "rowling": 'On', "sag": 'On', "san jacinto": 'On', "san jacinto garage": 'On', "guadalupe garage": 'On', "scottish rite dormitory": 'On', "sunchase condos": 'Off', "whitaker fields": 'Off', "21 rio": 'West', "orange tree condos": 'West', "st thomas condos": 'West', "overlook condos": 'West', "lenox condos": 'West',
  "lonestar lofts": 'West', "miller square": 'Off', "2001comal": 'On', "2105comal": 'On', "smc": 'On',
  "texas union": 'On', "trg": 'On', "utc": 'On',
  // West
  "somerset condos": 'West', // dispatcher: Somerset condos is West Campus
  "avon at 22nd": 'West',    // dispatcher: Avon at 22nd St is West Campus
  "phi gamma delta": 'West', // FIJI house, West Campus
  "stoneleigh condominiums": 'West', // 2409 Leon St (West every time in the sheet)
  "robbins place": 'West',           // 1903 Robbins Pl
  "1010w23rd": 'West', "1101w22nd": 'West', "1883 cameron": 'West', "1883 montgomery": 'West',
  "1883atcameron2707riogrande": 'West', "21 pearl": 'West', "2200pearl": 'West', "2204sangabriel": 'West',
  "2208pearl": 'West', "2212sangabriel": 'West', "2216sangabriel": 'West', "2222pearl": 'West', "2300nueces": 'West',
  "24 longview": 'West', "2414longview": 'West', "2502leon": 'West', "2510leon": 'West', "2510sangabriel": 'West',
  "26 west": 'West', "2606riogrande": 'West', "2616saladounitedstates": 'West', "2706salado": 'West',
  "2707riogrande": 'West', "2707riogrande1883atcameron": 'West', "2710nueces": 'West', "2801riogrande": 'West',
  "2810salado": 'West', "2815riogrande": 'West', "2819riogrande": 'West', "702w25th": 'West',
  "708grahamplace": 'West', "709w22nd": 'West', "714w22nd": 'West', "715w2212": 'West', "910w26th": 'West',
  "910west26th": 'West', "9atrio": 'West', "axis": 'West', "axis west": 'West', "blockon25th": 'West',
  "blockonrio": 'West', "callaway": 'West', "caminoflats": 'West', "castilian": 'West', "centennialcondos": 'West',
  "crest at pearl": 'West', "croixcondos": 'West', "grandmarc": 'West', "hilltop": 'West', "icon": 'West',
  "inspire": 'West', "inspire on 22nd": 'West', "ion": 'West', "lark": 'West', "legacy on rio": 'West',
  "lenoxcondos915w23": 'West', "lonestarlofts": 'West', "mark": 'West', "moontower": 'West', "noble 2500": 'West',
  "nueceshouse": 'West', "overlookcondominiums": 'West', "parkplace": 'West', "pearlcoop": 'West',
  "pointnorth": 'West', "quarters nueces": 'West', "quartersgrayson": 'West', "quartersgraysonhouse": 'West',
  "quartersongrayson": 'West', "quarterssterling": 'West', "regents west at 24th": 'West',
  "regents west at 26th": 'West', "riograndecondominiums": 'West', "rise at west campus": 'West',
  "sangabrielsquare": 'West', "sanremo": 'West', "signature 1909": 'West', "skyloft": 'West', "sparq on rio": 'West',
  "standardat": 'West', "stoneleighcondominiumson2409leon": 'West', "taoscoop": 'West', "texan west campus": 'West',
  "texanpearl": 'West', "the block on 23rd": 'West', "the block on 25th east": 'West',
  "the block on 25th west": 'West', "the block on 28th": 'West', "the block on leon": 'West', "the block on pearl": 'West', "the block on pearl north": 'West', "the block on pearl south": 'West',
  "the nine at rio": 'West', "the nine at west campus": 'West', "the rambler": 'West', "the sinclair": 'West',
  "the standard": 'West', "torre": 'West', "tradition on rio": 'West', "twenty two 15": 'West',
  "union on 24th": 'West', "union on san antonio": 'West', "venue on guadalupe": 'West',
  "villas on guadalupe": 'West', "villas on nueces": 'West', "villas on rio": 'West', "villas on 26th": 'West', "villason24th": 'West',
  "vintagewestcampus": 'West', "waterloo": 'West', "westridge": 'West', "yugo corner": 'West', "yugo nueces": 'West',
  "yugo rio": 'West', "yugo waterloo": 'West',
  // Off
  "1021e45th": 'Off', "106e30th": 'Off', "201e34th": 'Off', "2900cole": 'Off', "2910medicalarts": 'Off',
  "306e30th": 'Off', "306e30th2": 'Off', "3106duval": 'Off', "3111tomgreen": 'Off', "3115helms": 'Off',
  "3121speedway": 'Off', "31stcondominiums": 'Off', "3301speedway": 'Off', "3409speedway": 'Off',
  "3506speedway": 'Off', "3606grooms": 'Off', "403w38th": 'Off', "415w39th": 'Off', "4200avenuea": 'Off',
  "4305duval": 'Off', "4310avenueb": 'Off', "4406bennettave": 'Off', "4504speedway": 'Off', "4505speedway": 'Off',
  "4527nlamarblvd": 'Off', "609e49th": 'Off', "810w32nd": 'Off', "918e37th": 'Off', "duval villas": 'Off',
  "hydeparkcourtnorthcampus": 'Off', "kingston": 'Off', "oakpark": 'Off', "sucasa": 'Off', "tanglewood": 'Off',
  "tanglewood north": 'Off', "windtreecondominiums": 'Off',
};

/** Location zone for a drop-off ('On' | 'West' | 'Off' | 'North'), or '' when it isn't known. */
// Every legal on-campus place is known (the campus map + the sheet's history), so anything else is off campus:
// West when the address or name says West Campus, otherwise Off.
function zoneOf_(dropoff) {
  if (!String(dropoff || '').trim()) return '';
  const bare = dropNote_(dropoff);
  return ZONE_OF[placeKey_(dropoff)] || partZone_(dropoff) || (/\b78712\b/.test(dropoff) ? 'On' : '') || // 78712 = UT's own ZIP code
    zoneFromAddress_(dropoff) || (bare && zoneFromAddress_(bare)) ||
    embeddedAddressZone_(dropoff) || nameZone_(dropoff) || (mentionsCampusPlace_(dropoff) ? 'On' : 'Off'); // "greg gym next to jester"
}
// "NMS, 2506 Speedway, austin TX 78712": a known place named in one part of the text (split at commas / parentheses / " - ").
function partZone_(text) {
  const parts = String(text).split(/\s*[,;()\/]\s*|\s+-\s+|\s+or\s+/i).map(x => x.trim()).filter(Boolean); // also "UA9 or 2609 University Ave", "PCL/FAC"
  if (parts.length < 2) return '';
  for (const p of parts) { const c = canonPlace_(p); if (c && ZONE_OF[c.toLowerCase()]) return ZONE_OF[c.toLowerCase()]; }
  return '';
}
// "Lenox Condos (915 W 23 St)", "Lonestar Lofts, 2408 Leon St": an address anywhere in the text.
function embeddedAddressZone_(text) {
  const re = /\b\d{2,5}\s+(?:[nsew]\.?\s+|west\s+|east\s+)?[a-z0-9]+(?:\s+[a-z]+)?/gi;
  let m;
  while ((m = re.exec(String(text)))) { const z = zoneFromAddress_(m[0]); if (z) return z; }
  return '';
}
// Kinds of places with a usual zone: West Campus in the name, Greek houses and co-ops (West Campus), UT lots (On).
const GREEK_ = /\b(sorority|fraternity|alpha|beta|gamma|delta|epsilon|zeta|theta|iota|kappa|lambda|sigma|tau|upsilon|phi|chi|psi|omega|omicron)\b/i;
function nameZone_(text) {
  const t = String(text);
  if (/west\s*campus/i.test(t)) return 'West'; // "… West Campus" is West, literally in the name
  if (/^\s*(ut\s+)?lot\s*#?\s*\d{1,3}\b/i.test(t)) return 'On';
  if (GREEK_.test(t) || /\bco-?\s?op\b/i.test(t)) return 'West';
  return '';
}

/*
 * Street addresses that aren't in ZONE_OF, by Austin's grid (checked against the sheet's history):
 *   numbered street 29th or north ("621 W 37th")            -> Off   (Off 104 of 110 past rides)
 *   W 19th–28th ("714 W 22nd")                              -> West  (West 61 of 62)
 *   north–south street, house no. 2900+ ("3409 Speedway")   -> Off   (house number ≈ cross street: 3409 ≈ 34th)
 *   West Campus street, house no. 1800–2899 ("2819 Rio Grande") -> West
 * Anything else stays unknown for the dispatcher.
 */
const WEST_CAMPUS_STREETS = ['rio grande', 'nueces', 'san gabriel', 'pearl', 'leon', 'salado', 'longview', 'san antonio',
  'montgomery', 'graham', 'san pedro', 'robbins', 'wichita'];
function zoneFromAddress_(dropoff) {
  const t = String(dropoff || '').toLowerCase().replace(/[.,#]/g, ' ').replace(/\b(austin|tx|texas|787\d\d)\b/g, ' ')
    .replace(/\bdean\s*keaton\b/g, '26th').replace(/\b(mlk|martin luther king( jr)?( blvd)?)\b/g, '19th'); // named numbered streets
  // Robbins Pl (often typed "Robinson PI") is a West Campus street.
  if (/^\s*\d{2,5}\s+(robbins|robinson)\s+p[li]/.test(t)) return 'West';
  let m = /^\s*(\d{2,5})\s+(?:(w|e|west|east)\s+)?(\d{1,2})(?:st|nd|rd|th)?\b/.exec(t);
  if (m) {
    const n = Number(m[3]), num = Number(m[1]);
    if (n >= 29) return 'Off';
    if (n >= 19 && /^w/.test(m[2] || '')) return 'West';
    if (n >= 19 && /^e/.test(m[2] || '') && num >= 700) return 'Off'; // east of campus ("710 E Dean Keaton")
    return '';
  }
  m = /^\s*(\d{3,5})\s+(?:[nsew]\s+)?([a-z][a-z ]*)/.exec(t);
  if (m) {
    const num = Number(m[1]), street = m[2].trim();
    if (num >= 2900) return 'Off';
    if (num >= 1800 && WEST_CAMPUS_STREETS.some(w => street.startsWith(w))) return 'West';
  }
  return '';
}


/*
 * Rough map positions for "Suggested groups" in Pending (riders picked up near each other AND dropped off near
 * each other). Each address is looked up once with Apps Script's built-in geocoder (Maps service — no key needed)
 * and remembered for good in GEO_CACHE; an address it can't place near campus is remembered as 0 and the panel
 * falls back to same pickup place + same drop-off zone. A few new lookups per refresh, so refreshes stay fast.
 */
const GEO_CENTER = [30.2849, -97.7341]; // UT Tower
const GEO_MAX_M = 8000;                  // results farther than this from campus are wrong matches
function geoDist_(a, b) {
  const dy = (a[0] - b[0]) * 110540, dx = (a[1] - b[1]) * 111320 * Math.cos(a[0] * Math.PI / 180);
  return Math.sqrt(dx * dx + dy * dy);
}
/** addresses -> { placeKey_: [lat, lng] | 0 }; at most maxNew new lookups this call. */
function geoPlaces_(addresses, maxNew) {
  const cache = bigGet_('GEO_CACHE') || {}, out = {};
  let added = 0;
  [...new Set(addresses.filter(Boolean))].forEach(a => {
    const k = placeKey_(a);
    if (!k) return;
    if (k in cache) { out[k] = cache[k]; return; }
    if (added >= maxNew) return;
    added++;
    let pos = 0;
    try {
      const r = Maps.newGeocoder().setBounds(30.24, -97.78, 30.33, -97.70).geocode((canonPlace_(a) || a) + ', Austin, TX');
      const loc = r && r.status === 'OK' && r.results && r.results[0] && r.results[0].geometry.location;
      if (loc && geoDist_([loc.lat, loc.lng], GEO_CENTER) < GEO_MAX_M) pos = [+loc.lat.toFixed(5), +loc.lng.toFixed(5)];
    } catch (e) { added--; return; } // quota or network hiccup: not remembered, tried again later
    cache[k] = pos; out[k] = pos;
  });
  if (added) { try { bigSet_('GEO_CACHE', cache); } catch (e) { /* remembered next time */ } }
  return out;
}

/** Same spot: same place name, or (when both are on the map) within CHAIN_M metres. Only remembered positions. */
const CHAIN_M = 250;
function samePlace_(a, b) {
  if (placeKey_(a) === placeKey_(b)) return true;
  const g = geoPlaces_([a, b], 0), pa = g[placeKey_(a)], pb = g[placeKey_(b)];
  return !!(pa && pb && geoDist_(pa, pb) <= CHAIN_M);
}

/*
 * Golf carts stay on main campus: never to/from ECG, ECGA or the garages detached from it (edit this list freely).
 * Matched by place name, so every spelling of these counts ("SAG", "San Antonio Garage", ...).
 */
const CART_NO_GO = ['ECG', 'ECGA', 'East Campus Garage', 'SAG', 'San Antonio Garage', 'TRG', 'Trinity Garage', 'MNG', 'Manor Garage',
  'BRG', 'Brazos Garage', 'HCG', 'Health Center Garage', 'CCG', 'Conference Center Garage', 'GUG', 'Guadalupe Garage', 'NUG', 'Nueces Garage'];
let CART_NO_GO_KEYS_ = null;
function cartNoGo_(place) {
  if (!CART_NO_GO_KEYS_) CART_NO_GO_KEYS_ = new Set(CART_NO_GO.map(placeKey_));
  return CART_NO_GO_KEYS_.has(placeKey_(place));
}
/** Pickups must be on campus: a pickup we know is West / Off / North (an apartment or address) gets a warning. */
// Places we drop off at but never pick up from.
const DROPOFF_ONLY = ['Whitaker Fields', 'Moody Center'];
// The pickup is (or names) a drop-off-only place: "Moody Center", "moody center, 2001 robert dedman dr", "witaker feilds".
// Places riders write that we spell out for the crews — and write back into the sheet (pickup J / drop-off K):
// "Target" → "Target (Dobie)". Add more as { re, to }. Only a cell that is just that place is rewritten.
const PLACE_REWRITES = [
  { re: /^\s*(?:the\s+)?(?:ut\s+)?target(?:\s+(?:store|on\s+guad(?:alupe)?|at\s+guad(?:alupe)?|guad(?:alupe)?|dobie|at\s+dobie))?\s*[.!]*\s*$/i, to: 'Target (Dobie)' },
  { re: /^\s*(?:the\s+)?(?:alan\s+)?wright\s+(?:intramural\s+)?fields?\s*[.!]*\s*$|^\s*alan\s+wright\s*[.!]*\s*$/i, to: 'Whitaker Fields' }, // the name people know
  { re: /^\s*(?:durin|duran|durren|dureen)(?:\s+(?:hall|dorms?|residence hall))?\s*[.!]*\s*$/i, to: 'Duren' }, // misspellings of Duren
];
function placeRewrite_(text) {
  const t = String(text == null ? '' : text).trim(), w = PLACE_REWRITES.find(x => x.re.test(t));
  return w && t !== w.to ? w.to : null;
}
function dropoffOnly_(text) {
  const keys = new Set(DROPOFF_ONLY.map(p => placeKey_(p)));
  const parts = [text].concat(String(text).split(/\s*[,;()]\s*|\s+-\s+/));
  if (parts.some(x => x && keys.has(placeKey_(x)))) return true;
  const t = ' ' + String(text).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
  return DROPOFF_ONLY.some(p => [p].concat(PLACES[p] || []).some(a => { const w = String(a).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); return w.length >= 5 && t.includes(' ' + w + ' '); }));
}
function offCampusPickup_(pickup) {
  if (!String(pickup || '').trim()) return false;
  if (dropoffOnly_(pickup)) return true;
  return zoneOf_(pickup) !== 'On'; // only campus-map / known On places (or text naming one) are legal pickups
}
// "greg gym next to jester", "benches infront of pcl", "between fac and ut tower": the text names an On place.
let CAMPUS_PHRASES_ = null;
const SHORT_OK_ = new Set(['bsom', 'pcl', 'sac', 'sjc', 'j1', 'j2', 'ua9', 'fac', 'gdc', 'eer', 'wcp', 'sjh', 'rlm', 'pma', 'ecj', 'glt', 'kin', 'utc', 'mai', 'unb', 'cma', 'hrc', 'lbj',
  'brg', 'ccg', 'ecg', 'gug', 'hcg', 'mag', 'rhg', 'sag', 'sjg', 'swg', 'trg', 'tsg', 'dhg', 'dkr']);
function mentionsCampusPlace_(text) {
  if (!CAMPUS_PHRASES_) {
    CAMPUS_PHRASES_ = [];
    Object.keys(PLACES).filter(k => ZONE_OF[k.toLowerCase()] === 'On' && !DROPOFF_ONLY.includes(k)).forEach(k =>
      [k].concat(PLACES[k]).forEach(v => {
        const w = String(v).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
        if (w && !/^\d+$/.test(w) && (w.length >= 4 || SHORT_OK_.has(w))) CAMPUS_PHRASES_.push(' ' + w + ' ');
      }));
  }
  const t = ' ' + String(text).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
  return CAMPUS_PHRASES_.some(w => t.includes(w));
}


/*
 * Every building and garage on the UT campus map (Building Index + Parking Garages, AY 2022/23; checked against the 2026 index — added DCG, NUG and the new names): code, name,
 * [existing place it is the same as], [extra names people type]. Merged into PLACES at load: a code or name that
 * already means a place joins it, else it becomes its own place. All are UT property, so pickups there are allowed:
 * Location defaults to On unless the sheet's history already says otherwise.
 * Codes listed together ("cma,cmb") are buildings that share one entry on the map (CMA-B, LLA-F, FC1-8).
 */
const CAMPUS_MAP = [
  ['ACA', 'Academic Annex'], ['ADH', 'Almetris Duren Hall', 'Duren'], ['AFP', 'Athletics Field Pavilion'],
  ['AHG', 'Anna Hiss Gymnasium', '', 'anna hiss'], ['ANB', 'Arno Nowotny Building'], ['', 'Andrews Dormitory', 'Andrews'],
  ['ARC', 'Animal Resources Center'], ['ART', 'Art Building'], ['ASE', 'Aerospace Engineering'],
  ['ATT', 'AT&T Executive Education and Conference Center', '', 'att center', 'at&t center', 'at&t conference center'],
  ['BAT', 'Batts Hall', '', 'batts'], ['BEL', 'L. Theo Bellmont Hall', 'Belmont', 'bellmont'], ['BEN', 'Benedict Hall', '', 'benedict'],
  ['BHD', 'Brackenridge Hall Dormitory', 'Brackenridge'], ['BIO', 'Biological Laboratories', '', 'bio labs'], ['BLD', 'Blanton Dormitory', 'Blanton'],
  ['BMA', 'Blanton Museum of Art', '', 'blanton museum'], ['BMC', 'Dealey Center for New Media', '', 'belo', 'belo center', 'belo center for new media'],
  ['BME', 'Biomedical Engineering Building'], ['BOT', 'Biological Greenhouse'], ['BRB', 'Bernard and Audre Rapoport Building', '', 'rapoport'],
  ['BTL', 'Battle Hall'], ['BUR', 'Burdine Hall', '', 'burdine'], ['BWY', 'Bridgeway Building', '', '2616 wichita'],
  ['CAL', 'Calhoun Hall', '', 'calhoun'], ['CBA', 'McCombs School of Business'], ['CCF', 'Caven Lacrosse and Sports Center'],
  ['CCJ', 'John B. Connally Center for Justice'], ['CDA', 'Comal Child Development Center Annex'], ['CDL', 'Collections Deposit Library'],
  ['CEE', 'Continuing Engineering Education'], ['CLK', 'Caven Clark Field Support Building'],
  ['CMA,CMB', 'Jesse H. Jones Communication Center'], ['CML', 'Child Development Center'], ['COM', 'Computation Center'],
  ['CPE', 'Chemical and Petroleum Engineering'], ['CRB', 'Central Receiving Building'], ['CRD', 'Carothers Dormitory', 'Carothers'],
  ['CRH', 'Creekside Residence Hall', 'Creekside'], ['CS3', 'Chilling Station No. 3'], ['CS4', 'Chilling Station No. 4'],
  ['CS5', 'Chilling Station No. 5'], ['CS6', 'Chilling Station No. 6'],
  ['DCP', 'Denton A. Cooley Pavilion', '', 'cooley pavilion'], ['DEV', 'Development Office Building'],
  ['DFA', 'E. William Doty Fine Arts Building', '', 'fine arts', 'doty fine arts'], ['DFF', 'UFCU Disch-Falk Field', '', 'disch falk', 'disch-falk', 'falk disch', 'falk-disch', 'disch falk field', 'falk disch field',
    'disch', 'baseball field', 'baseball stadium', 'ut baseball field', 'ufcu disch falk'], // the baseball field: on campus
  ['DHG', 'Dobie Hall Garage', '', 'dobie garage'], ['DCG', 'Dobie Center Garage', '', 'dobie twenty one garage', 'dobie 21 garage', 'dobie twenty21 garage'], ['', 'Dell Seton Medical Center', '', 'dell seton'] /* "DMC" here means the Dealey Center (riders' usage), not this */, ['DTO', 'Dobie Twenty One', 'Dobie Twenty21'],
  ['EAS', 'Edgar A. Smith Building'], ['ECJ', 'Ernest Cockrell Jr. Hall', '', 'cockrell'],
  ['EER,EERC', 'Engineering Education and Research Center', '', 'eer lawn'], ['EPS', 'E. P. Schoch Building', '', 'schoch'],
  ['ERC', 'Erwin Special Events Center', '', 'erwin center'], ['ESS', 'Computer Science Annex'], ['ETC', 'Engineering Teaching Center II', '', 'etc ii', 'etc 2', 'etc2', 'etcii', 'etc two', 'engineering teaching center 2', 'engineering teaching center two', 'etc building'],
  ['FAC', 'Peter T. Flawn Academic Center', '', 'flawn', 'flawn academic center'],
  ['FC1,FC2,FC3,FC4,FC5,FC6,FC7,FC8', 'Facilities Complex'], ['FDF', 'Frank Denius Fields', '', 'denius fields'],
  ['FDH', 'J. Frank Dobie House'], ['FNT', 'Faulkner Nano Science and Technology Building'], ['FPC', 'OFPC Field Staff Office'],
  ['GAR', 'Garrison Hall', '', 'garrison'], ['GDC', 'Gates Dell Complex', '', 'gates dell', 'gates'], ['GEA', 'Mary E. Gearing Hall', '', 'gearing'],
  ['GEB', 'Dorothy L. Gebauer Building', '', 'gebauer'], ['GOL', 'Goldsmith Hall', '', 'goldsmith'], ['GRE', 'Gregory Gymnasium', 'Gregory Gym'],
  ['GWB', 'Gordon White Building'], ['GSB', 'Graduate School of Business Building'],
  ['HDB', 'Health Discovery Building'], ['HLB', 'Health Learning Building'], ['HMA', 'Hogg Memorial Auditorium', '', 'hogg auditorium'],
  ['HRC', 'Harry Ransom Center', '', 'ransom center'], ['HRH', 'Homer Rainey Hall'], ['HSM', 'William Randolph Hearst Building'],
  ['HTB', 'Health Transformation Building'], ['IPF', 'Indoor Practice Facility'], ['INT', 'International Office'],
  ['JCD', 'Jester Dormitory', 'Jester'], ['JES', 'Beauford H. Jester Center', 'Jester'], ['JGB', 'Jackson Geological Sciences Building', '', 'jackson geo', 'jackson school'],
  ['JHH', 'John W. Hargis Hall'], ['JON', 'Jesse H. Jones Hall'], ['KIN', 'Kinsolving Dormitory', 'Kinsolving'],
  ['LBJ', 'Lyndon B. Johnson Library', '', 'lbj library'], ['LCH', 'Littlefield Carriage House'], ['LDH', 'Longhorn Dining Facility'],
  ['LFH', 'Littlefield Home', 'Littlefield House'], ['LLA,LLB,LLC,LLD,LLE,LLF', 'Living Learning Centers', '', 'living learning center'],
  ['LTD', 'Littlefield Dormitory', 'Littlefield'], ['LTH', 'Laboratory Theatre Building'],
  ['MAI', 'Main Building', '', 'ut tower', 'the tower', 'tower'], ['MBB', 'Moffett Molecular Biology Building'], ['MEZ', 'Mezes Hall', '', 'mezes'],
  ['MFH', 'Mithoff Track and Soccer Field House', '', 'mithoff track/soccer field house', 'mithoff track soccer field house', 'mithoff field house'], ['MHD', 'Moore-Hill Dormitory', 'Moore-Hill'],
  ['MMS', 'Mike A. Myers Track and Soccer Stadium', '', 'myers stadium'], ['MNC', 'Moncrief-Neuhaus Athletic Center'],
  ['MRH', 'Music Building and Recital Hall', 'Butler School of Music', 'music building'], ['MSB', 'Mail Services Building'], ['MTC', 'Material Transfers Center'],
  ['NEZ', 'North End Zone Building'], ['NHB', 'Norman Hackerman Building', '', 'hackerman'], ['NMS', 'Neural and Molecular Science Building'],
  ['NUR', 'Nursing School'], ['PAC', 'Performing Arts Center'], ['PAI', 'T.S. Painter Hall', '', 'painter', 'painter hall'],
  ['PAR', 'Parlin Hall', '', 'parlin'], ['PAT', 'J.T. Patterson Laboratories Building', '', 'patterson'], ['PCL', 'Perry-Castaneda Library', 'PCL'],
  ['PHD', 'Prather Hall Dormitory', 'Prather'], ['PHR', 'Pharmacy Building', '', 'pharmacy'], ['POB', "Peter O'Donnell Jr. Building", '', 'odonnell'],
  ['PPA', 'Hal C. Weaver Power Plant Annex'], ['PPE', 'Hal C. Weaver Power Plant Expansion'], ['PPL', 'Hal C. Weaver Power Plant'],
  ['RHD', 'Roberts Hall Dormitory', 'Roberts'], ['RRH', 'Robert B. Rowling Hall', '', 'rowling', 'rowling hall'], ['RLM', 'Robert Lee Moore Hall'],
  ['RLP', 'Patton Hall'], ['RSC', 'Recreational Sports Center', 'Rec Center'],
  ['SAC', 'Student Activity Center', '', 'wcp', 'william c powers', 'powers student activity center', 'william c. powers jr. student activity center', 'william c powers jr student activity center'], ['SBS', 'Red and Charline McCombs Field'],
  ['SEA', 'Sarah M. and Charles E. Seay Building', '', 'seay'], ['SER', 'Service Building'], ['SJH', 'San Jacinto Residence Hall', 'San Jacinto'],
  ['SRH', 'Sid Richardson Hall'], ['SSB', 'Student Services Building', '', 'student services'], ['SSW', 'School of Social Work Building', '', 'social work'],
  ['STD', 'DKR-Texas Memorial Stadium', 'DKR Stadium'], ['SUT', 'Sutton Hall', 'Sutton'], ['SW7', '2617 Speedway'],
  ['SZB', 'George I. Sanchez Building', '', 'sanchez'], ['TCC', 'Joe C. Thompson Conference Center', '', 'thompson conference center'],
  ['TMM', 'Texas Memorial Museum'], ['TNH', 'Townes Hall', '', 'law school'], ['TSC', 'Jamail Texas Swimming Center', '', 'swim center'],
  ['TTC', 'Texas Tennis Center'], ['UA9', '2609 University Avenue', '', '2609 university ave', '2609 university', '2609 univ ave', 'ua9 building'], ['UIL', 'University Interscholastic League'],
  ['UNB', 'Union Building', 'Texas Union'], ['UPB', 'University Police Building', '', 'utpd'], ['USS', 'University Sign Shop'],
  ['UTA', 'UT Administration Building', 'UT Administration'], ['UTC', 'University Teaching Center'], ['UTX', 'Etter-Harbin Alumni Center', '', 'alumni center', 'etter harbin', 'etter harbin alumni center', 'etter', 'ehac', 'harbin'],
  ['WAG', 'Waggener Hall'], ['WAT', 'Arthur P. Watson House'], ['WCH', 'Will C. Hogg Building'], ['WEL', 'Robert A. Welch Hall', 'Welch'],
  ['WIN', 'F. Loren Winship Drama Building', 'Winship'], ['WMB', 'West Mall Office Building'], ['WRW', 'W. R. Woolrich Laboratories', '', 'woolrich'],
  ['WWH', 'Walter Webb Hall'],
  ['GLT', 'Gary L. Thomas Energy Engineering Building', '', 'gary l thomas', 'gary thomas'], // newer than the 2022 map
  ['', 'Dell Medical School', '', 'dell medical center', 'dell med', 'dell medical', 'dell med school', 'dell medical school campus',
    'dell med center', 'dell medical school building'], // on campus (its HDB / HLB / HTB buildings are separate entries)
  ['', 'Kinsolving Dormitory', 'Kinsolving', 'kins dining', 'kins dining hall', 'kinsolving dining', 'kinsolving dining hall', 'kins cafe'],
  // Other campus spots riders use as pickups (from the sheet's history)
  ['', 'Littlefield Fountain', '', 'fountain', 'little field fountain', 'littlefield fountain'],
  ['', 'Littlefield Home', 'Littlefield House', 'littlefield cafe', 'littlefield patio cafe'],
  ['PMA', 'Robert Lee Moore Hall', '', 'physics math astronomy', 'physics, math & astronomy building', 'physics math and astronomy building', 'physics math & astronomy', 'pma building'], ['', 'International Office', '', 'texas global'],
  ['', 'The Bubble', '', 'bubble'], ['', 'Clark Field', '', 'caven clark field'],
  ['', 'Performing Arts Center', '', 'bass concert hall', 'bass hall', 'bass', 'across bass concert hall', 'texas performing arts', 'texas performing arts building'],
  ['', 'Nursing School', '', 'school of nursing', 'nursing building'], ['', 'Mike A. Myers Track and Soccer Stadium', '', 'mike a myers soccer field', 'mike a myers stadium', 'myers'],
  ['', 'Moody Center', '', 'moody center arena'], ['VAC', 'Visual Arts Center', '', 'visual arts building', 'visual arts', 'visual art building', 'vac building'], ['', 'Townes Hall', '', 'ut law school', 'law school'],
  ['', 'Lyndon B. Johnson Library', '', 'lbj presidential library', 'lbj public affairs', 'lbj school'],
  ['DMC', 'Dealey Center for New Media', '', 'dealy center', 'dealey center', 'dealy center of new media', 'dealey center for new media', 'kut',
    'dealey media center', 'dealy media center', 'dealey new media', 'dealey center for new media',
    'g. b. dealey center for new media', 'gb dealey center for new media', 'g.b. dealey center', 'gb dealey center', 'dealey',
    'moody auditorium', 'dmc 2.106'], // "DMC" = the G. B. Dealey Center for New Media (Moody College, 300 W. Dean Keeton)
  ['', 'DKR-Texas Memorial Stadium', 'DKR Stadium', 'dkr gate 6', 'stadium'], ['', 'Etter-Harbin Alumni Center', '', 'texas exes center', 'texas exes'],
  ['', 'Sutton Hall', 'Sutton', 'sutten', 'sutten hall'], ['', 'San Jacinto Residence Hall', 'San Jacinto', 'san jacinto resident hall', 'cypress bend cafe', 'cypress bend'],
  ['', 'Harry Ransom Center', '', 'harry randsom center', 'ransom'], ['', 'Robert B. Rowling Hall', '', 'rowlling'],
  ['', 'Art Building', '', 'art and art history building', 'art building and museum'], ['', 'Aerospace Engineering', '', 'aerospace building'],
  ['', 'Facilities Complex', '', 'facilities complex building 1'],
  ['', 'Clark Field', '', 'clarke field'], ['', 'Inner Campus Drive'], ['', 'Texas Memorial Museum', '', 'natural history museum'],
  ['ECGA', 'East Campus Graduate Apartments', '', 'east campus grad apartments'],
  ['', 'Chemical and Petroleum Engineering', '', 'chemical engineering building', 'chem eng building', 'new chemical eng building', 'petroleum engineering building', 'petroleum engineering', 'petroleum building', 'pete building', 'chemical and petroleum engineering building'],
  ['', 'Gregory Gymnasium', 'Gregory Gym', 'eddie reese outdoor pool', 'outdoor pool', 'gregory pool'], ['', 'Almetris Duren Hall', 'Duren', 'durren'],
  ['', "Peter O'Donnell Jr. Building", '', 'peter odonnell', 'peter o donnell'], ['', 'San Jacinto Residence Hall', 'San Jacinto', 'san jancinto', 'san j', 'san j bus stop'],
  ['', 'Townes Hall', '', 'school of law'], ['', 'Goldsmith Hall', '', 'school of architecture'],
  ['', 'Music Building and Recital Hall', 'Butler School of Music', 'bates recital hall', 'bates hall'], ['', 'Waggener Hall', '', 'wagner hall', 'waggener'],
  ['', 'Homer Rainey Hall', '', 'rainey hall'], ['', 'DKR-Texas Memorial Stadium', 'DKR Stadium', 'drk'],
  ['', 'Creekside Residence Hall', 'Creekside', 'waller creek', 'waller creek residence halls'], ['', 'Union Building', 'Texas Union', 'texas union on 24th', 'union on 24th st'],
  // Parking garages
  ['BRG', 'Brazos Garage'], ['CCG', 'Conference Center Garage'], ['ECG', 'East Campus Garage'], ['GUG', 'Guadalupe Garage'],
  ['HCG', 'Health Center Garage'], ['MAG', 'Manor Garage'], ['RHG', 'Rowling Hall Garage'], ['SAG', 'San Antonio Garage'],
  ['SJG', 'San Jacinto Garage'], ['SWG', 'Speedway Garage'], ['TRG', 'Trinity Garage'], ['TSG', '27th Street Garage'], ['NUG', 'Nueces Garage'],
];
(function addCampusMap_() {
  const norm = placeNorm_;
  const owner = s => { const n = norm(s); return n ? Object.keys(PLACES).find(k => norm(k) === n || PLACES[k].some(v => norm(v) === n)) : null; };
  CAMPUS_MAP.forEach(([codes, name, join, ...more]) => {
    const codeList = codes ? codes.split(',') : [];
    const target = join || owner(name) || codeList.map(owner).find(Boolean) || name;
    if (!PLACES[target]) PLACES[target] = [];
    codeList.concat([name], more).forEach(a => {
      const v = a.toLowerCase(), o = owner(v);
      if (!o || o === target) { if (!PLACES[target].includes(v)) PLACES[target].push(v); }
    });
    const key = target.toLowerCase();
    if (!ZONE_OF[key]) ZONE_OF[key] = codeList.concat([name], more).map(x => ZONE_OF[x.toLowerCase()] || ZONE_OF[norm(x)]).find(Boolean) || 'On';
  });
  PLACE_INDEX_ = null;
})();
