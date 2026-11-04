/* Macademia — Decision Trees 1.0 pilot (based on editor v11.2 / 1.0c). Based on Microsoft's MIT-licensed Office starter.
 * See LICENSE in the project root. */
/* global console, document, Excel, Office */
                                              
                            
                    
                                                               
                                                            
                 
                      
                   
                      
                             
 
                 
                                                       
                                                
                        
                      
                     
                         
                                 
                                                
 
const KEY = "MacDecisionTree.editor.v1.";
const PREFIX = "MDT_Editor_";
let current               = null;
let currentSheet = "";
let stage = "starting";
let busy = false;
let editingId = "";
let formSnapshot = "";
let pendingSelection                                                  = null;
let selectionScheduled = false;
const watchedSheets = new Set        ();
const pendingChanges = new Map                     ();
let drainScheduled = false;
function el                                     (id        )    { return document.getElementById(id)     ; }
function input(id        )                   { return el                  (id); }
function clone   (v   )    { return JSON.parse(JSON.stringify(v))     ; }
function children(m       , id        )             { return m.nodes.filter(n => n.parent === id); }
function root(m       )           {
  const list = m.nodes.filter(n => n.parent === null);
  if (list.length !== 1) throw new Error("A tree must have exactly one root.");
  return list[0];
}
function chosen(m       , id = el                   ("node").value)           {
  const n = m.nodes.find(n => n.id === id);
  if (!n) throw new Error("Load the active tree again and select a node.");
  return n;
}
function column(i        )         {
  let text = "";
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + (n - 1) % 26) + text;
  return text;
}
function addr(n          , offset        , value = false)         {
  const col = n.col + (value ? 3 : offset === -2 && n.card ? 1 : 0);
  return `${column(col)}${n.row + (n.card && offset > 0 ? offset - 1 : offset)}`;
}
const COLORS = { chance: "#4776CC", decision: "#29955B", outcome: "#F3C84B" };
const GLYPHS                       = { chance: "●", decision: "■", outcome: "▲" };
const PAPER = "#F6F3EE", CARD = "#FFFEFB", INK = "#25282D";
function cardRange(n          )         {
  return `${column(n.col)}${n.row - 2}:${column(n.col + 4)}${n.row + (n.kind === "outcome" ? 1 : 3)}`;
}
function probabilityText(value      )         {
  return typeof value === "number" ? `${Number((value * 100).toFixed(6))}%` : value;
}
function expr(text        , title        )       {
  const v = text.trim();
  if (v.startsWith("=") && v.length > 1) return v;
  if (!v || !Number.isFinite(Number(v))) throw new Error(`${title} must be a number or Excel formula beginning with =.`);
  return Number(v);
}
function payoffExpr(text        )       {
  return text.trim() === "" ? 0 : expr(text, "Net payoff");
}
function probabilityExpr(text        )       {
  const v = text.trim();
  if (v.endsWith("%")) {
    const number = Number(v.slice(0, -1));
    if (!v.slice(0, -1).trim() || !Number.isFinite(number)) throw new Error("Enter a probability such as 80% or 0.8.");
    return number / 100;
  }
  const result = expr(v, "Probability");
  return typeof result === "number" && result > 1 && result <= 100 ? result / 100 : result;
}
// Entered probabilities stay fixed. Every Auto branch shares the unused mass.
function ensureRemainders(m       )       {
  const legacy = m.probabilityMode !== "shared-auto";
  for (const n of m.nodes) {
    if (m.nodes.find(p => p.id === n.parent)?.kind !== "chance") {
      n.remainder = false;
      continue;
    }
    // Preserve old manual entries on upgrade; Reset group to Auto starts a
    // fresh equal split when the previous defaults should no longer be fixed.
    if (legacy) n.probabilityFixed = m.probabilityMode === "remainder" ? !n.remainder : true;
    n.remainder = !n.probabilityFixed;
  }
  m.probabilityMode = "shared-auto";
}
function setRemainder(m       , parent          , id        )       {
  const list = children(m, parent.id);
  if (parent.kind !== "chance" || !list.some(n => n.id === id)) throw new Error("Choose a branch leaving a chance node.");
  const n = chosen(m, id);
  n.probabilityFixed = false;
  n.remainder = true;
}
function remainderFormula(m       , n          )         {
  const siblings = children(m, n.parent );
  const count = siblings.filter(c => !c.probabilityFixed).length;
  if (!count) throw new Error("No automatic branch is available.");
  const other = siblings.filter(c => c.probabilityFixed).map(c => addr(c, 1, true));
  if (!other.length) return `=1/${count}`;
  const args = other.join(",");
  return `=IF(AND(COUNT(${args})=${other.length},MIN(${args})>=0,MAX(${args})<=1,SUM(${args})<=1.000000001),MAX(0,1-SUM(${args}))/${count},NA())`;
}
function displayedProbability(m       , n          )       {
  if (!n.remainder) return n.probability;
  const list = children(m, n.parent );
  const fixed = list.filter(c => c.probabilityFixed);
  if (fixed.every(c => typeof c.probability === "number")) {
    const total = fixed.reduce((sum, c) => sum + Number(c.probability), 0);
    return Math.max(0, 1 - total) / list.filter(c => !c.probabilityFixed).length;
  }
  return n.probability;
}
function chanceGroup(m       , n          )                       {
  if (n.kind === "chance") return n;
  return m.nodes.find(p => p.id === n.parent && p.kind === "chance");
}
function normalizedFormula(value         )         {
  return typeof value === "string" ? value.replace(/[\s$]/g, "").toUpperCase() : `VALUE:${String(value)}`;
}
function probabilitySignature(m       )         {
  return JSON.stringify(m.nodes.filter(n => m.nodes.find(p => p.id === n.parent)?.kind === "chance")
    .map(n => [n.id, !!n.probabilityFixed, n.probabilityFixed ? n.probability : null]));
}
function probabilityNotice(m       )         {
  for (const parent of m.nodes.filter(n => n.kind === "chance")) {
    const list = children(m, parent.id), fixed = list.filter(n => n.probabilityFixed);
    if (!list.length || fixed.some(n => typeof n.probability !== "number")) continue;
    const total = fixed.reduce((sum, n) => sum + Number(n.probability), 0);
    if (total > 1 + 1e-9) return `${parent.label}: Fixed probabilities exceed 100%. Lower an entered probability or switch a branch to Auto. Its value remains #N/A until corrected.`;
    if (fixed.length === list.length && Math.abs(total - 1) > 1e-6) {
      return `${parent.label}: all branches are Fixed, but they do not total 100%. Switch at least one branch to Auto or correct an entered value.`;
    }
  }
  return "";
}
function probabilityValidity(m       , parent          , absolute = false)         {
  const refs = children(m, parent.id).map(n => {
    const address = addr(n, 1, true);
    if (!absolute) return address;
    const match = /^([A-Z]+)(\d+)$/.exec(address) ; return `$${match[1]}$${match[2]}`;
  });
  if (!refs.length) return "FALSE";
  const args = refs.join(",");
  return `IFERROR(AND(COUNT(${args})=${refs.length},MIN(${args})>=0,MAX(${args})<=1,ABS(SUM(${args})-1)<0.000001),FALSE)`;
}
function installProbabilityWarnings(sheet                 , m       )       {
  // Native Excel rules and formulas continue to flag and clear errors even when
  // the task pane is closed or an input formula changes on another worksheet.
  for (const parent of m.nodes.filter(n => n.kind === "chance")) {
    const list = children(m, parent.id); if (!list.length) continue;
    const valid = probabilityValidity(m, parent), invalid = `=NOT(${probabilityValidity(m, parent, true)})`;
    for (const n of list) {
      const row = n.card ? `${addr(n, 1)}:${column(n.col + 4)}${n.row}` : addr(n, 1, true);
      const area = sheet.getRange(row); area.conditionalFormats.clearAll();
      const warning = area.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      warning.custom.rule.formula = invalid;
      warning.custom.format.fill.color = "#FCE0DF"; warning.custom.format.font.color = "#9C1C24";
      const caption = m.view === "presentation" ? "Probability" : n.remainder ? "Probability (Auto)" : "Probability (Fixed)";
      sheet.getRange(addr(n, 1)).formulas = [[`=IF(${valid},"${caption}","! ${caption}")`]];
    }
    if (parent.card) {
      const footer = sheet.getRange(`${addr(parent, 4)}:${column(parent.col + 4)}${parent.row + 3}`); footer.merge();
      const total = `IFERROR(TEXT(SUM(${list.map(n => addr(n, 1, true)).join(",")}),"0.0%"),"invalid")`;
      const hiddenCount = parent.collapsed ? descendants(m, parent.id).length : 0;
      const normal = hiddenCount ? `"[+${hiddenCount}] nodes hidden · Expand in sidebar"` : '""';
      sheet.getRange(addr(parent, 4)).formulas = [[`=IF(${valid},${normal},"! Total "&${total}&"; set a branch Auto or fix inputs")`]];
      footer.format.font.size = 9; footer.format.shrinkToFit = true;
      footer.format.indentLevel = 1; footer.format.font.color = INK;
      footer.conditionalFormats.clearAll();
      const warning = footer.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      warning.custom.rule.formula = invalid; warning.custom.format.fill.color = "#FCE0DF";
      warning.custom.format.font.color = "#9C1C24";
    }
  }
}
async function refreshProbabilityAlerts(ctx                      , sheet                 , m       )                {
  if (typeof document === "undefined") return;
  const reads = m.nodes.filter(n => m.nodes.find(p => p.id === n.parent)?.kind === "chance").map(n => {
    const range = sheet.getRange(addr(n, 1, true)); range.load("values"); return { n, range };
  }); await ctx.sync();
  const messages           = [];
  for (const parent of m.nodes.filter(n => n.kind === "chance")) {
    const group = reads.filter(r => r.n.parent === parent.id); if (!group.length) continue;
    const values = group.map(r => r.range.values[0][0]);
    const numeric = values.every(v => typeof v === "number" && Number.isFinite(v));
    const total = numeric ? (values            ).reduce((sum, v) => sum + v, 0) : NaN;
    if (numeric && (values            ).every(v => v >= 0 && v <= 1) && Math.abs(total - 1) < 1e-6) continue;
    const allFixed = group.every(r => r.n.probabilityFixed);
    const totalText = numeric ? `${Number((total * 100).toFixed(4))}%` : "an invalid value";
    messages.push(`${parent.label} (${parent.id}): probabilities total ${totalText}. ${allFixed ? "Every branch is Fixed, so none can balance the remainder. " : "Check the Fixed entries; they may exceed 100% or contain an invalid input. "}Select a branch you want calculated, check “Auto: share the remaining percentage”, then Apply changes. Keep Fixed entries at or below 100% in total. Alternatively, correct the entered probabilities so this group totals 100%.`);
  }
  const host = el("probability-alert"); host.textContent = messages.length ? "! Probability check\n\n" + messages.join("\n\n") : "";
  host.classList.toggle("hidden", !messages.length);
}

function changeContains(change        , target        )          {
  function point(value        )                                      {
    const match = /^([A-Z]+)([0-9]+)$/i.exec(value.replace(/\$/g, ""));
    if (!match) return null;
    let col = 0;
    for (const letter of match[1].toUpperCase()) col = col * 26 + letter.charCodeAt(0) - 64;
    return { row: Number(match[2]), col };
  }
  const cell = point(target);
  if (!cell) return false;
  return change.split(",").some(area => {
    const pair = area.trim().slice(area.lastIndexOf("!") + 1).split(":");
    const first = point(pair[0]), last = point(pair[1] || pair[0]);
    return !!first && !!last && cell.row >= Math.min(first.row, last.row)
      && cell.row <= Math.max(first.row, last.row) && cell.col >= Math.min(first.col, last.col)
      && cell.col <= Math.max(first.col, last.col);
  });
}
async function writeProbabilityCells(ctx                      , sheet                 , m       )                {
  ensureRemainders(m);
  stage = "updating shared Auto probabilities";
  for (const n of m.nodes) {
    if (m.nodes.find(p => p.id === n.parent)?.kind !== "chance") continue;
    const cell = sheet.getRange(addr(n, 1, true));
    if (n.remainder) cell.formulas = [[remainderFormula(m, n)]];
    cell.numberFormat = [["0.0%"]];
    const valueArea = n.card ? sheet.getRange(`${addr(n, 1, true)}:${column(n.col + 4)}${n.row}`) : cell;
    valueArea.format.fill.color = m.view === "presentation" ? CARD : n.remainder ? "#E9EDF5" : "#FFF0CC";
    cell.format.font.bold = true;
    sheet.getRange(addr(n, 1)).values = [[m.view === "presentation" ? "Probability" : n.remainder ? "Probability (Auto)" : "Probability (Fixed)"]];
  }
  installProbabilityWarnings(sheet, m);
}
async function watchSheet(ctx                      , sheet                 )                {
  if (watchedSheets.has(sheet.id)) return;
  stage = "enabling live probability editing";
  sheet.onChanged.add(onWorksheetChanged);
  sheet.onSelectionChanged.add(onWorksheetSelectionChanged);
  await ctx.sync();
  watchedSheets.add(sheet.id);
}
function onWorksheetChanged(event                                 )       {
  if (event.changeType !== "RangeEdited") return;
  if (!busy && event.worksheetId === currentSheet) invalidateReport("Worksheet edited. Generate a fresh PDF snapshot.");
  const changes = pendingChanges.get(event.worksheetId) || new Set        ();
  changes.add(event.address);
  pendingChanges.set(event.worksheetId, changes);
  scheduleDrain();
}
function scheduleDrain()       {
  if (busy || drainScheduled || !pendingChanges.size) return;
  drainScheduled = true;
  setTimeout(() => { void drainChanges(); }, 0);
}
async function drainChanges()                {
  drainScheduled = false;
  if (busy || !pendingChanges.size) return;
  const [sheetId, changes] = pendingChanges.entries().next().value                         ;
  pendingChanges.delete(sheetId);
  await perform(async ctx => {
    stage = "reading a probability edit";
    const sheet = ctx.workbook.worksheets.getItem(sheetId);
    const setting = ctx.workbook.settings.getItemOrNullObject(KEY + sheetId);
    setting.load("value");
    await ctx.sync();
    if (setting.isNullObject) return "";
    const m = JSON.parse(String(setting.value))         ;
    if (m.probabilityMode !== "shared-auto") return "";
    const relevant = m.nodes.some(n => m.nodes.find(p => p.id === n.parent)?.kind === "chance"
      && [...changes].some(change => changeContains(change, addr(n, 1, true))));
    if (!relevant) return "";
    const before = clone(m), signature = probabilitySignature(m);
    await capture(ctx, sheet, m);
    // Formula/format writes also generate events. Stop here for our own writes,
    // so the live handler settles without recursion or repeated redraws.
    if (signature === probabilitySignature(m)) { await refreshProbabilityAlerts(ctx, sheet, m); return ""; }
    await writeProbabilityCells(ctx, sheet, m);
    ctx.workbook.settings.add(KEY + sheetId, JSON.stringify(m));
    ctx.workbook.settings.add(KEY + sheetId + ".undo", JSON.stringify(before));
    await ctx.sync();
    if (currentSheet === sheetId) {
      if (!unappliedChanges()) populate(m, sheetId);
      await refreshProbabilityAlerts(ctx, sheet, m);
    }
    return probabilityNotice(m) || "Entered probabilities are Fixed. The remaining Auto branches now share the unused percentage equally. Save with Cmd+S.";
  }, true);
}
function parseProbabilityList(text        , count        )                                                   {
  const parts = text.trim().split(/[\/,;\s]+/).filter(Boolean);
  if (parts.length !== count) throw new Error(`Enter ${count} entries, one per branch in the displayed order.`);
  const automatic = parts.map((v, i) => /^(\*|x|auto)$/i.test(v) ? i : -1).filter(i => i >= 0);
  const probabilities = parts.map(v => {
    if (/^(\*|x|auto)$/i.test(v)) return 0;
    const raw = v.replace(/%$/, "");
    if (!raw) throw new Error("Enter a percentage number or * for the remainder.");
    const number = Number(raw);
    if (!Number.isFinite(number) || number < 0 || number > 100) throw new Error("List entries are percentages from 0 to 100, or * for the remainder.");
    return number / 100;
  });
  const total = probabilities.reduce((sum, v) => sum + v, 0);
  if (total > 1 + 1e-9) throw new Error("These probabilities exceed 100%. Reduce the entered percentages.");
  if (!automatic.length && Math.abs(total - 1) > 1e-9) throw new Error("The list must total 100%, or use * for the probability to calculate automatically.");
  if (automatic.length) for (const i of automatic) probabilities[i] = Math.max(0, 1 - total) / automatic.length;
  return { probabilities, automatic };
}
function descendants(m       , id        )             {
  const list             = [];
  function visit(parent        ) { for (const n of children(m, parent)) { list.push(n); visit(n.id); } }
  visit(id); return list;
}
function layout(m       )       {
  if (m.nodes.length > 60) throw new Error("This editor supports up to 60 nodes. Copying or inserting this branch would exceed that limit.");
  if (new Set(m.nodes.map(n => n.id)).size !== m.nodes.length) throw new Error("Duplicate node identifiers.");
  const seen = new Set        ();
  function validate(n          , depth        , hidden         ) {
    if (seen.has(n.id)) throw new Error("The tree contains a cycle.");
    if (depth > 6) throw new Error("This editor supports six branch levels. Inserting or copying here would exceed that limit.");
    seen.add(n.id); n.hidden = hidden; n.card = true;
    const list = children(m, n.id);
    if (n.kind === "outcome" && list.length) throw new Error("An outcome cannot have child branches.");
    if (!list.length) n.collapsed = false;
    for (const c of list) validate(c, depth + 1, hidden || !!n.collapsed);
  }
  validate(root(m), 0, false);
  if (seen.size !== m.nodes.length) throw new Error("Some nodes are disconnected.");
  let leaf = 0, maxDepth = 0;
  function place(n          , depth        )         {
    n.col = 1 + depth * 7; maxDepth = Math.max(maxDepth, depth);
    const list = children(m, n.id).filter(c => !c.hidden);
    if (list.length) { const rows = list.map(c => place(c, depth + 1)); n.row = Math.round((rows[0] + rows[rows.length - 1]) / 2); }
    else { n.row = 11 + leaf * 8; leaf++; }
    return n.row;
  }
  place(root(m), 0);
  const hidden = m.nodes.filter(n => n.hidden);
  // Dedicated hidden columns retain native formulas for every collapsed descendant.
  // Start after BOTH the visible tree and its document header (through column P).
  m.hiddenCol = hidden.length ? Math.max(17, 1 + maxDepth * 7 + 7) : undefined;
  hidden.forEach((n, i) => { n.col = m.hiddenCol ; n.row = 11 + i * 8; });
  m.rows = Math.max(...m.nodes.map(n => n.row)) + 8;
  m.cols = Math.max(16, 1 + maxDepth * 7 + 6, m.hiddenCol === undefined ? 0 : m.hiddenCol + 6);
}
function reorderBranch(m       , n          , delta        )       {
  if (!n.parent) throw new Error("The root has no sibling branches to reorder.");
  const list = children(m, n.parent), index = list.findIndex(c => c.id === n.id), next = index + delta;
  if (next < 0 || next >= list.length) throw new Error(delta < 0 ? "This is already the first branch." : "This is already the last branch.");
  const a = m.nodes.indexOf(n), b = m.nodes.indexOf(list[next]); [m.nodes[a], m.nodes[b]] = [m.nodes[b], m.nodes[a]];
}
function duplicateBranch(m       , n          )         {
  if (!n.parent) throw new Error("Choose a branch below the root to duplicate.");
  const originals = [n, ...descendants(m, n.id)], ids = new Map(originals.map(c => [c.id, `n${m.nextId++}`]));
  const copies = originals.map(c => ({ ...clone(c), id: ids.get(c.id) , parent: c.id === n.id ? n.parent : ids.get(c.parent ) , hidden: false, row: 0, col: 0 }));
  copies[0].label = `${n.label.slice(0, 73)} (copy)`;
  if (m.nodes.find(p => p.id === n.parent)?.kind === "chance") { copies[0].probabilityFixed = false; copies[0].remainder = true; }
  m.nodes.splice(m.nodes.indexOf(n) + 1, 0, ...copies);
  ensureRemainders(m); layout(m); return copies[0].id;
}
function insertBefore(m       , n          , kind                       )         {
  if (!n.parent) throw new Error("Choose a branch below the root to insert a node along it.");
  const wrapper           = { id: `n${m.nextId++}`, parent: n.parent, kind, label: n.label, probability: n.probability,
    probabilityFixed: n.probabilityFixed, remainder: n.remainder, payoff: 0, row: 0, col: 0 };
  m.nodes.splice(m.nodes.indexOf(n), 0, wrapper);
  n.parent = wrapper.id; n.probability = 1; n.probabilityFixed = false; n.remainder = kind === "chance";
  ensureRemainders(m); layout(m); return wrapper.id;
}
function nodeAtSelection(m       , address        )                       {
  const cell = address.slice(address.lastIndexOf("!") + 1).split(/[:,]/)[0].replace(/\$/g, "");
  const hit = m.nodes.filter(n => !n.hidden && changeContains(cardRange(n), cell));
  return hit.length === 1 ? hit[0] : undefined;
}
function formFingerprint()         {
  return JSON.stringify([input("label").value, el                   ("kind").value, input("payoff").value,
    input("probability").value, input("remainder").checked, input("probability-list").value, branchDrafts]);
}
function unappliedChanges()          { return !!current && !!formSnapshot && formFingerprint() !== formSnapshot; }
function requireApplied()          {
  if (!unappliedChanges()) return true;
  el("status").textContent = "You have unapplied sidebar edits. Apply changes or Discard edits before switching nodes or using tree tools.";
  if (editingId) el                   ("node").value = editingId;
  return false;
}
function onWorksheetSelectionChanged(event                                          )       {
  if (event.worksheetId !== currentSheet) return;
  pendingSelection = { worksheetId: event.worksheetId, address: event.address }; scheduleSelection();
}
function scheduleSelection()       {
  if (busy || selectionScheduled || !pendingSelection) return;
  selectionScheduled = true;
  setTimeout(() => {
    selectionScheduled = false; if (busy || !pendingSelection) return;
    const next = pendingSelection; pendingSelection = null;
    if (!current || next.worksheetId !== currentSheet) return;
    const n = nodeAtSelection(current, next.address);
    if (!n || n.id === editingId || !requireApplied()) return;
    el                   ("node").value = n.id; void load(false, true);
  }, 0);
}
function addOutcome(m       , parent          , label        , probability        )           {
  const n           = { id: `n${m.nextId++}`, parent: parent.id, kind: "outcome", label, probability, payoff: 0, row: 0, col: 0 };
  m.nodes.push(n); return n;
}
function makeTree(sample         )        {
  const m        = { version: 1, nextId: 2, objective: "max", rows: 0, cols: 0, probabilityMode: "shared-auto",
    nodes: [{ id: "n1", parent: null, kind: "decision", label: "Choose an option", probability: 1, payoff: 0, row: 0, col: 0 }] };
  const first = addOutcome(m, m.nodes[0], sample ? "Take the gamble" : "Option A", 0.5);
  const second = addOutcome(m, m.nodes[0], sample ? "Certain payoff" : "Option B", 0.5);
  if (sample) {
    first.kind = "chance"; second.payoff = 30000;
    const success = addOutcome(m, first, "Success", 0.6); success.payoff = 100000; success.probabilityFixed = true;
    addOutcome(m, first, "Failure", 0.4).payoff = -20000;
  }
  layout(m); ensureRemainders(m); return m;
}
function populate(m       , sheetId        , selectId         )       {
  const previous = selectId || el                   ("node").value;
  const switchedSheet = currentSheet !== sheetId;
  current = clone(m); currentSheet = sheetId;
  const dropdown = el                   ("node"); dropdown.replaceChildren();
  function append(n          , depth        ) {
    const option = document.createElement("option"); option.value = n.id;
    option.textContent = `${"  ".repeat(depth)}${n.id}: ${n.label} (${n.kind}${n.remainder ? "; auto probability" : ""})`;
    if (n.collapsed) option.textContent += ` [+${descendants(m, n.id).length} hidden]`;
    dropdown.appendChild(option); children(m, n.id).filter(c => !c.hidden).forEach(c => append(c, depth + 1));
  }
  append(root(m), 0);
  dropdown.value = m.nodes.some(n => n.id === previous && !n.hidden) ? previous : root(m).id;
  editingId = dropdown.value;
  if (switchedSheet || !input("report-title").value.trim()) input("report-title").value = root(m).label;
  invalidateReport("Tree loaded or edited. Generate a fresh PDF snapshot.");
  el                   ("view").value = m.view || "edit";
  el                   ("objective").value = m.objective; fillForm();
}
function fillForm()       {
  if (!current) return;
  const n = chosen(current);
  input("label").value = n.label; el                   ("kind").value = n.kind;
  input("payoff").value = String(n.payoff); input("probability").value = probabilityText(displayedProbability(current, n));
  const parent = current.nodes.find(c => c.id === n.parent);
  input("remainder").checked = !!n.remainder;
  input("remainder").disabled = parent?.kind !== "chance";
  input("probability").disabled = parent?.kind !== "chance";
  el("probability-note").textContent = parent?.kind === "chance"
    ? n.remainder ? "Auto shares the unused percentage equally. Type a probability to fix this branch, then Apply changes."
      : "This probability is fixed. Check Auto and Apply changes to let it share the remaining percentage again."
    : "Probability applies only to branches leaving a chance node.";
  el("selected").textContent = `${n.kind.charAt(0).toUpperCase() + n.kind.slice(1)} · ${children(current, n.id).length} ${children(current, n.id).length === 1 ? "branch" : "branches"}`;
  el("node-title").textContent = n.label;
  el("incoming").style.display = parent?.kind === "chance" ? "block" : "none";
  for (const mode of ["edit", "presentation"]) el("view-" + mode).setAttribute("aria-pressed", String((current.view || "edit") === mode));
  const group = chanceGroup(current, n);
  el("probability-group").style.display = group ? "block" : "none";
  el("probability-order").textContent = group
    ? children(current, group.id).map((c, i) => `${i + 1}. ${c.label} — ${c.remainder ? "Auto" : "Fixed"}`).join("\n") : "";
  input("probability-list").value = "";
  renderBranchEditor(current, n);
  el("collapse").textContent = n.collapsed ? "Expand branch" : "Collapse branch";
  el("collapse-note").textContent = n.collapsed ? `${descendants(current, n.id).length} descendants hidden. Their full calculations remain live.` : "Collapse hides descendants without changing their values.";
  input("show-strategy").checked = current.showStrategy !== false;
  refreshActionAvailability();
  refreshSensitivity(current);
  formSnapshot = formFingerprint();
}
function refreshActionAvailability()       {
  if (!current || !editingId) return;
  const n = chosen(current, editingId), siblings = n.parent ? children(current, n.parent) : [];
  const index = siblings.findIndex(c => c.id === n.id);
  el                   ("collapse").disabled = !children(current, n.id).length;
  el                   ("move-up").disabled = !n.parent || index === 0;
  el                   ("move-down").disabled = !n.parent || index === siblings.length - 1;
  el                   ("duplicate").disabled = !n.parent;
  el                   ("insert-node").disabled = !n.parent;
  el                   ("expand-all").disabled = !current.nodes.some(c => c.collapsed);
}
                                                                                                                                            
let branchDrafts                = [];
function renderBranchEditor(m       , n          )       {
  const host = el("branch-rows"); host.replaceChildren();
  branchDrafts = n.kind === "chance" ? children(m, n.id).map(c => ({ id: c.id, probability: probabilityText(displayedProbability(m, c)),
    payoff: String(c.payoff), fixed: !!c.probabilityFixed, probabilityEdited: false, payoffEdited: false })) : [];
  el("branches").style.display = branchDrafts.length ? "block" : "none";
  for (const draft of branchDrafts) {
    const node = chosen(m, draft.id), row = document.createElement("div"); row.className = "branch-row";
    const label = document.createElement("div"); label.className = "branch-name"; label.textContent = node.label; row.appendChild(label);
    const controls = document.createElement("div"); controls.className = "branch-fields";
    function field(title        , control             ) { const wrap = document.createElement("label"); wrap.textContent = title; wrap.appendChild(control); controls.appendChild(wrap); }
    const probability = document.createElement("input"); probability.value = draft.probability; probability.id = "bp-" + draft.id;
    probability.setAttribute("aria-label", node.label + " probability");
    const mode = document.createElement("select"); mode.id = "bm-" + draft.id; mode.setAttribute("aria-label", node.label + " probability mode");
    for (const value of ["fixed", "auto"]) { const option = document.createElement("option"); option.value = value; option.textContent = value === "fixed" ? "Fixed" : "Auto"; mode.appendChild(option); }
    mode.value = draft.fixed ? "fixed" : "auto";
    const payoff = document.createElement("input"); payoff.value = draft.payoff; payoff.setAttribute("aria-label", node.label + " payoff");
    probability.oninput = () => { draft.probability = probability.value; draft.fixed = true; draft.probabilityEdited = true; mode.value = "fixed"; refreshBranchPreview(); };
    mode.onchange = () => { draft.fixed = mode.value === "fixed"; draft.probabilityEdited = true; draft.probability = probability.value; refreshBranchPreview(); };
    payoff.oninput = () => { draft.payoff = payoff.value; draft.payoffEdited = true; };
    field("Probability", probability); field("Mode", mode); field("Payoff", payoff);
    row.appendChild(controls); host.appendChild(row);
  }
  refreshBranchPreview();
}
function refreshBranchPreview()       {
  if (!branchDrafts.length) return;
  try {
    const fixed = branchDrafts.filter(d => d.fixed).map(d => probabilityExpr(d.probability));
    if (fixed.some(v => typeof v !== "number")) { el("branch-total").textContent = "Formula inputs are evaluated by Excel when applied."; return; }
    const numbers = fixed            ;
    if (numbers.some(v => v < 0 || v > 1)) throw new Error("Each probability must be between 0% and 100%.");
    const total = numbers.reduce((a, b) => a + b, 0), automatic = branchDrafts.filter(d => !d.fixed);
    if (total > 1 + 1e-9) throw new Error("Fixed probabilities exceed 100%. Reduce an entry or switch it to Auto.");
    for (const draft of automatic) {
      draft.probability = probabilityText(Math.max(0, 1 - total) / automatic.length);
      input("bp-" + draft.id).value = draft.probability;
    }
    el("branch-total").textContent = automatic.length ? `Total 100% · ${automatic.length} Auto ${automatic.length === 1 ? "branch shares" : "branches share"} the remainder.`
      : Math.abs(total - 1) < 1e-6 ? "Total 100% · All branches Fixed." : `Total ${Number((total * 100).toFixed(4))}% · Switch a branch to Auto or correct an entry.`;
  } catch (e) { el("branch-total").textContent = (e         ).message; }
}
function applyBranchDrafts(m       , drafts               )       {
  for (const draft of drafts) {
    const n = m.nodes.find(c => c.id === draft.id); if (!n) continue;
    if (draft.payoffEdited) n.payoff = payoffExpr(draft.payoff);
    if (draft.probabilityEdited && m.nodes.find(c => c.id === n.parent)?.kind === "chance") {
      n.probabilityFixed = draft.fixed;
      if (draft.fixed) {
        n.probability = probabilityExpr(draft.probability);
        if (typeof n.probability === "number" && (n.probability < 0 || n.probability > 1)) throw new Error("Probability must be between 0% and 100%.");
      }
    }
  }
  ensureRemainders(m);
}
async function perform(work                                                , quiet = false)                {
  if (busy) return; busy = true;
  const locked = Array.from(document.querySelectorAll                                                          ("input,button,select"));
  const wasDisabled = new Map(locked.map(c => [c, c.disabled]));
  locked.forEach(c => { c.disabled = true; });
  if (!quiet) el("status").textContent = "Working...";
  try {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) throw new Error("This session requires ExcelApi 1.9 or later.");
    const message = await Excel.run(work);
    if (message) el("status").textContent = message;
  } catch (error) {
    console.error(error);
    const d = error                                                            ;
    el("status").textContent = [`Stopped while ${stage}.`, d.message || String(error), d.code ? `Code: ${d.code}` : "",
      d.debugInfo ? JSON.stringify(d.debugInfo, null, 2) : ""].filter(Boolean).join("\n\n");
  } finally {
    busy = false; locked.forEach(c => { if (c.isConnected) c.disabled = wasDisabled.get(c) || false; });
    if (current) {
      const selected = current.nodes.find(n => n.id === el                   ("node").value);
      const hasChanceParent = !!selected && current.nodes.some(n => n.id === selected.parent && n.kind === "chance");
      input("probability").disabled = !hasChanceParent; input("remainder").disabled = !hasChanceParent;
    }
    refreshActionAvailability();
    scheduleDrain(); scheduleSelection();
  }
}
async function active(ctx                      )                                                    {
  stage = "loading the active tree";
  const sheet = ctx.workbook.worksheets.getActiveWorksheet(); sheet.load("id"); await ctx.sync();
  const setting = ctx.workbook.settings.getItemOrNullObject(KEY + sheet.id); setting.load("value"); await ctx.sync();
  if (setting.isNullObject) throw new Error("Select an editor tree sheet, or click New tree. The old fixed demo is separate.");
  const model = JSON.parse(String(setting.value))         ;
  if (model.version !== 1) throw new Error("This model uses a different version.");
  layout(clone(model)); return { sheet, model };
}
async function capture(ctx                      , sheet                 , m       )                {
  stage = "reading worksheet inputs";
  const existingSharedMode = m.probabilityMode === "shared-auto";
  ensureRemainders(m);
  const reads = m.nodes.map(n => {
    const name = sheet.getRange(addr(n, -2)), payoff = sheet.getRange(addr(n, 2, true));
    const parent = m.nodes.find(c => c.id === n.parent);
    const probability = parent?.kind === "chance" ? sheet.getRange(addr(n, 1, true)) : null;
    name.load("values"); payoff.load("values,formulas"); probability?.load("values,formulas");
    return { n, name, payoff, probability };
  });
  await ctx.sync();
  // Typing over a generated Auto formula explicitly fixes that branch. Inspect
  // all formulas before updating flags, so a multi-cell paste is one edit.
  const typed = existingSharedMode ? reads.filter(r => r.n.remainder && r.probability
    && normalizedFormula(r.probability.formulas[0][0]) !== normalizedFormula(remainderFormula(m, r.n))) : [];
  for (const r of typed) { r.n.probabilityFixed = true; r.n.remainder = false; }
  const clearedPayoffs                = [];
  for (const r of reads) {
    r.n.label = String(r.name.values[0][0] || r.n.id).slice(0, 80);
    for (const [key, cell] of [["payoff", r.payoff], ["probability", r.probability]]         ) {
      if (!cell) continue;
      const v = cell.values[0][0];
      const formula = cell.formulas[0][0];
      // A cleared payoff means zero, just as Excel's references to blank cells do.
      // A formula returning text/empty string is still an error, not a cleared input.
      if (key === "payoff" && (v === "" || v === null || v === undefined)
        && !(typeof formula === "string" && formula.startsWith("="))) {
        r.n.payoff = 0; clearedPayoffs.push(cell); continue;
      }
      if (key === "probability" && r.n.remainder) {
        // Cache the displayed result for switching back to manual; never retain
        // a generated formula's old cell addresses when the layout changes.
        if (typeof v === "number" && Number.isFinite(v)) r.n.probability = v;
        continue;
      }
      if (typeof v !== "number" || !Number.isFinite(v)) throw new Error(`${key} for ${r.n.label} (${r.n.id}, ${addr(r.n, key === "payoff" ? 2 : 1, true)}) must evaluate to a number. Enter a number or correct its formula.`);
      if (key === "probability" && (v < 0 || v > 1)) throw new Error(`Probability for ${r.n.label} must be between 0 and 1.`);
      r.n[key] = typeof formula === "string" && formula.startsWith("=") ? formula : v;
    }
  }
  // Write explicit zeros only after every input passes validation.
  for (const cell of clearedPayoffs) cell.values = [[0]];
  if (clearedPayoffs.length) await ctx.sync();
}
function evaluationFormula(m       , n          )         {
  const own = addr(n, 2, true), list = children(m, n.id);
  if (!list.length) return n.kind === "outcome" ? `=${own}` : "=NA()";
  const values = list.map(c => addr(c, 3, true));
  if (n.kind === "decision") return `=${own}+${m.objective === "max" ? "MAX" : "MIN"}(${values.join(",")})`;
  const ps = list.map(c => addr(c, 1, true)), args = ps.join(",");
  const terms = list.map((c, i) => `${ps[i]}*${values[i]}`);
  return `=IF(AND(COUNT(${args})=${list.length},MIN(${args})>=0,MAX(${args})<=1,ABS(SUM(${args})-1)<0.000001),${own}+${terms.join("+")},NA())`;
}
function choiceFormula(m       , n          )         {
  const list = children(m, n.id);
  if (!list.length) return '="Add branches"';
  const values = list.map(c => addr(c, 3, true));
  const best = `${m.objective === "max" ? "MAX" : "MIN"}(${values.join(",")})`;
  const ties = values.map(v => `IF(${v}=${best},1,0)`).join("+");
  let choice = '""';
  for (const c of [...list].reverse()) choice = `IF(${addr(c, 3, true)}=${best},"Best: "&${addr(c, -2)},${choice})`;
  return `=IFERROR(IF((${ties})>1,"Equal best value: "&(${ties})&" options",${choice}),"Check inputs")`;
}
function strategyFormula(m       , n          )         {
  const parent = m.nodes.find(p => p.id === n.parent);
  if (!parent || parent.kind !== "decision") return '=""';
  const list = children(m, parent.id), values = list.map(c => addr(c, 3, true));
  const best = `${m.objective === "max" ? "MAX" : "MIN"}(${values.join(",")})`;
  const ties = values.map(v => `IF(${v}=${best},1,0)`).join("+");
  return `=IFERROR(IF(${addr(n, 3, true)}=${best},IF((${ties})>1,"TIED BEST","BEST"),"Alternate"),"Check inputs")`;
}
function outline(range             , color        )       {
  for (const side of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]         ) {
    const border = range.format.borders.getItem(side); border.style = "Continuous"; border.color = color;
  }
}
async function markSelection(ctx                      , sheet                 , m       , id        )                {
  if (m.layoutVersion !== 8) return;
  for (const n of m.nodes.filter(n => !n.hidden)) outline(sheet.getRange(cardRange(n)), m.view !== "presentation" && n.id === id ? "#365AA5" : "#8B8985");
  await ctx.sync();
}
async function draw(ctx                      , sheet                 , m       , previous        , selectId         )                {
  layout(m); ensureRemainders(m); m.layoutVersion = 8; m.rendererVersion = 9;
  const presenting = m.view === "presentation";
  const selectedId = selectId && m.nodes.some(n => n.id === selectId && !n.hidden) ? selectId : root(m).id;
  stage = "preparing the v9 tree canvas";
  sheet.load("id"); sheet.shapes.load("items/name"); await ctx.sync();
  // Persist inputs before touching the canvas. A failed drawing can be retried
  // without reading a partially cleared worksheet as fresh model data.
  const recovery = ctx.workbook.settings.add(KEY + sheet.id + ".recovery", JSON.stringify({ model: m, previous: previous || m }));
  ctx.workbook.settings.add(KEY + sheet.id, JSON.stringify(previous || m));
  await ctx.sync();
  for (const shape of sheet.shapes.items) if (shape.name.startsWith(PREFIX)) shape.delete();
  // Finish deletion and unmerge before resizing: old merged rows must not affect anchors.
  const canvas = sheet.getRangeByIndexes(0, 0, Math.max(m.rows, previous?.rows || 0), Math.max(m.cols, previous?.cols || 0));
  canvas.columnHidden = false; canvas.unmerge(); canvas.clear(); await ctx.sync();
  canvas.format.columnWidth = 48; canvas.format.rowHeight = 18;
  canvas.format.fill.color = PAPER; canvas.format.font.color = INK;
  canvas.format.font.name = "Aptos"; canvas.format.font.size = 11;
  canvas.format.wrapText = false; canvas.format.shrinkToFit = false;
  canvas.format.verticalAlignment = "Center"; sheet.showGridlines = false;
  const money = '$#,##0;($#,##0)';
  function text(cell        , value        ) { sheet.getRange(cell).values = [[/^[=+@-]/.test(value) ? "'" + value : value]]; }
  function merged(address        )              { const r = sheet.getRange(address); r.merge(); return r; }
  function number(n          , offset        , value      , format        , editable         ) {
    const address = addr(n, offset, true), row = n.row + offset - 1;
    const range = merged(`${address}:${column(n.col + 4)}${row}`);
    const cell = sheet.getRange(address);
    if (typeof value === "string" && value.startsWith("=")) cell.formulas = [[value]]; else cell.values = [[value]];
    cell.numberFormat = [[format]]; range.format.font.bold = true; range.format.shrinkToFit = true;
    range.format.horizontalAlignment = "Right";
    range.format.fill.color = presenting ? CARD : editable ? "#FFF0CC" : "#E9EDF5";
  }
  function cellSymbol(area             , anchor        , kind      , size        ) {
    // Real cell content uses the same zoom/layout engine as the card labels.
    // Floating shape coordinates are deliberately not used for node symbols.
    sheet.getRange(anchor).values = [[GLYPHS[kind]]];
    area.format.font.name = "Arial"; area.format.font.size = size;
    area.format.font.bold = false; area.format.font.color = COLORS[kind];
    area.format.horizontalAlignment = "Center"; area.format.verticalAlignment = "Center";
    area.format.wrapText = false; area.format.shrinkToFit = false;
  }
  const header = sheet.getRange("B2:P4"); header.format.fill.color = "#E6E3DF"; outline(header, "#8B8985");
  cellSymbol(sheet.getRange("B2"), "B2", "chance", 12);
  cellSymbol(sheet.getRange("B3"), "B3", "decision", 12);
  cellSymbol(sheet.getRange("B4"), "B4", "outcome", 12);
  merged("C2:I4"); text("C2", "Decision analysis"); sheet.getRange("C2").format.font.size = 20; sheet.getRange("C2").format.font.bold = true;
  merged("J2:L4"); text("J2", "Expected value");
  const total = merged("M2:P4"); sheet.getRange("M2").formulas = [[`=${addr(root(m), 3, true)}`]]; sheet.getRange("M2").numberFormat = [[money]];
  total.format.font.size = 18; total.format.font.bold = true; total.format.shrinkToFit = true;
  merged("B6:P6"); text("B6", (m.objective === "max" ? "Maximize net payoff" : "Minimize value / cost") + "   |   Square: decision   Circle: chance   Triangle: outcome");
  sheet.getRange("B6").format.font.size = 10;
  merged("B7:P7"); text("B7", m.showStrategy !== false ? "BEST / TIED BEST = choices at a decision node. [+N] = collapsed descendants; full calculations remain live." : "[+N] = collapsed descendants. Full calculations remain live.");
  sheet.getRange("B7").format.font.size = 10;
  for (const n of m.nodes) {
    const card = sheet.getRange(cardRange(n)); card.format.fill.color = CARD;
    const heading = sheet.getRange(`${column(n.col)}${n.row - 2}:${column(n.col + 4)}${n.row - 1}`);
    heading.format.fill.color = "#F0EEEA";
    const badgeAddress = `${column(n.col)}${n.row - 2}`;
    const badge = merged(`${badgeAddress}:${column(n.col)}${n.row - 1}`);
    cellSymbol(badge, badgeAddress, n.kind, 21);
    const label = merged(`${addr(n, -2)}:${column(n.col + 4)}${n.row - 1}`);
    label.format.horizontalAlignment = "Center";
    label.format.verticalAlignment = "Center";
    label.format.indentLevel = 0;
    label.format.shrinkToFit = false;
    label.format.wrapText = true; label.format.font.bold = true; label.format.font.size = n.label.length > 60 ? 9 : n.label.length > 38 ? 10 : 12;
    text(addr(n, -2), n.label);
    heading.format.borders.getItem("EdgeBottom").style = "Continuous";
    heading.format.borders.getItem("EdgeBottom").color = "#B6B2AA";
    for (const offset of [1, 2, 3]) {
      const captions = merged(`${addr(n, offset)}:${column(n.col + 2)}${n.row + offset - 1}`);
      captions.format.shrinkToFit = true; captions.format.indentLevel = 1;
    }
    if (m.nodes.find(c => c.id === n.parent)?.kind === "chance") {
      text(addr(n, 1), presenting ? "Probability" : n.remainder ? "Probability (Auto)" : "Probability (Fixed)");
      number(n, 1, n.remainder ? remainderFormula(m, n) : n.probability, "0.0%", !n.remainder);
    }
    if (m.showStrategy !== false && m.nodes.find(p => p.id === n.parent)?.kind === "decision") {
      text(addr(n, 1), "Decision branch"); number(n, 1, strategyFormula(m, n), "General", false);
      sheet.getRange(addr(n, 1, true)).format.font.color = "#244D36";
    }
    text(addr(n, 2), n.kind === "outcome" ? "Net payoff" : "Payoff at node"); number(n, 2, n.payoff, money, true);
    text(addr(n, 3), n.kind === "outcome" ? "" : "Expected value");
    number(n, 3, evaluationFormula(m, n), n.kind === "outcome" ? ";;;" : money, false);
    if (n.kind === "outcome") sheet.getRange(`${addr(n, 3)}:${column(n.col + 4)}${n.row + 2}`).format.fill.color = PAPER;
    if ((n.kind === "decision" && m.showStrategy !== false) || n.collapsed) {
      const footer = merged(`${addr(n, 4)}:${column(n.col + 4)}${n.row + 3}`);
      const hiddenCount = n.collapsed ? descendants(m, n.id).length : 0;
      const formula = n.kind === "decision" && m.showStrategy !== false
        ? choiceFormula(m, n) + (hiddenCount ? `&" · [+${hiddenCount}]"` : "")
        : `="[+${hiddenCount}] nodes hidden · Expand in sidebar"`;
      sheet.getRange(addr(n, 4)).formulas = [[formula]];
      footer.format.shrinkToFit = true; footer.format.font.size = 10; footer.format.indentLevel = 1;
    }
    outline(card, !presenting && n.id === selectedId ? "#365AA5" : "#8B8985");
  }
  canvas.format.rowHeight = 18;
  // Connectors live entirely on borders of the two blank gutter columns.
  // Headers are two rows tall; their midpoint is an exact worksheet boundary.
  // Using that same grid for cards and routes avoids shape zoom/point drift.
  function horizontal(row        , firstCol        , lastCol        )       {
    if (lastCol < firstCol) return;
    const range = sheet.getRange(`${column(firstCol)}${row}:${column(lastCol)}${row}`);
    const edge = range.format.borders.getItem("EdgeTop");
    edge.style = "Continuous"; edge.color = "#4C4D50";
  }
  for (const parent of m.nodes.filter(n => !n.hidden && !n.collapsed)) {
    const list = children(m, parent.id).filter(n => !n.hidden); if (!list.length) continue;
    stage = `joining card borders from ${parent.label}`;
    const parentPort = parent.row - 1;
    const busCol = parent.col + 6;
    horizontal(parentPort, parent.col + 5, busCol - 1);
    const ports = [parentPort, ...list.map(n => n.row - 1)];
    const firstRow = Math.min(...ports), lastRow = Math.max(...ports);
    if (lastRow > firstRow) {
      const trunk = sheet.getRange(`${column(busCol)}${firstRow}:${column(busCol)}${lastRow - 1}`);
      const edge = trunk.format.borders.getItem("EdgeLeft");
      edge.style = "Continuous"; edge.color = "#4C4D50";
    }
    for (const child of list) horizontal(child.row - 1, busCol, child.col - 1);
  }
  if (m.hiddenCol !== undefined) {
    sheet.getRange(`${column(m.hiddenCol)}1:${column(m.hiddenCol + 4)}${m.rows}`).columnHidden = true;
  }
  stage = "adding probability warning flags";
  installProbabilityWarnings(sheet, m);
  stage = "saving the finished tree";
  ctx.workbook.settings.add(KEY + sheet.id, JSON.stringify(m));
  rememberEditor(ctx, sheet.id, selectedId);
  if (previous) ctx.workbook.settings.add(KEY + sheet.id + ".undo", JSON.stringify(previous));
  sheet.getRange(addr(chosen(m, selectedId), -2)).select();
  await ctx.sync();
  recovery.delete(); await ctx.sync();
  populate(m, sheet.id, selectedId); await refreshProbabilityAlerts(ctx, sheet, m); await watchSheet(ctx, sheet);
}
async function modify(change                                          , allowDirty = false)                {
  if (!allowDirty && !requireApplied()) return;
  const id = el                   ("node").value;
  await perform(async ctx => {
    const { sheet, model } = await active(ctx);
    if (sheet.id !== currentSheet) throw new Error("The active worksheet changed. Click Load active tree first.");
    await capture(ctx, sheet, model); const before = clone(model);
    stage = "applying the selected edit";
    const next = change(model, chosen(model, id)); layout(model);
    await draw(ctx, sheet, model, before, typeof next === "string" ? next : id);
    return probabilityNotice(model) || "Tree updated. Fixed probabilities stay unchanged; Auto branches share the rest equally. Save with Cmd+S.";
  });
}
// Session metadata lives in workbook settings, beside the full tree model.
// Closing the task pane clears JavaScript memory but never removes these records.
                                                                                             
let resumeSheetId = "";
function parseEditorSession(value         )                       {
  try {
    const data = JSON.parse(String(value))                 ;
    return data?.version === 1 && typeof data.sheetId === "string" && typeof data.nodeId === "string"
      && data.sheetId.length > 0 && data.nodeId.length > 0 ? data : null;
  } catch { return null; }
}
function rememberEditor(ctx                      , sheetId        , nodeId        , restoredTitle         )       {
  const title = restoredTitle || (typeof document !== "undefined" && currentSheet === sheetId ? input("report-title").value.trim().slice(0, 80) : undefined);
  const session                = { version: 1, sheetId, nodeId, ...(title ? { reportTitle: title } : {}) };
  ctx.workbook.settings.add(KEY + sheetId + ".session", JSON.stringify(session));
  ctx.workbook.settings.add(KEY + "lastSession", JSON.stringify(session));
}
function resumeNode(model       , session                      )         {
  const n = model.nodes.find(n => n.id === session?.nodeId && !n.hidden);
  return n?.id || root(model).id;
}
async function restoreOnOpen()                {
  let shouldLoad = false;
  await perform(async ctx => {
    stage = "looking for a saved editor session";
    const sheet = ctx.workbook.worksheets.getActiveWorksheet(); sheet.load("id"); await ctx.sync();
    const stored = ctx.workbook.settings.getItemOrNullObject(KEY + sheet.id);
    const last = ctx.workbook.settings.getItemOrNullObject(KEY + "lastSession"); stored.load("value"); last.load("value"); await ctx.sync();
    shouldLoad = !stored.isNullObject;
    const session = last.isNullObject ? null : parseEditorSession(last.value);
    resumeSheetId = session?.sheetId || "";
    el("resume-last").classList.toggle("hidden", shouldLoad || !resumeSheetId);
    return shouldLoad ? "Restoring this workbook's tree…" : resumeSheetId
      ? "This sheet is not a tree. Click Resume last tree, or select a tree sheet and Load active tree."
      : "Create a tree, or select a saved tree sheet and Load active tree. Save working trees as .xlsx workbooks.";
  });
  if (shouldLoad) await load();
}
async function resumeLastTree()                {
  if (!requireApplied() || !resumeSheetId) return;
  let activated = false;
  await perform(async ctx => {
    stage = "finding the last edited tree";
    const sheets = ctx.workbook.worksheets; sheets.load("items/id"); await ctx.sync();
    const sheet = sheets.items.find(s => s.id === resumeSheetId);
    if (!sheet) throw new Error("The last tree sheet was deleted. Select another tree sheet and click Load active tree.");
    sheet.activate(); await ctx.sync(); activated = true; return "Opening the last edited tree…";
  });
  if (activated) await load();
}
async function saveTreeWorkbook()                {
  if (!requireApplied()) return;
  await perform(async ctx => {
    const { sheet, model } = await active(ctx);
    if (sheet.id !== currentSheet) throw new Error("Load this tree before saving its editor session.");
    stage = "checking the workbook before saving";
    ctx.workbook.load("name,readOnly");
    const pending = ctx.workbook.settings.getItemOrNullObject(KEY + sheet.id + ".recovery"); pending.load("value"); await ctx.sync();
    if (ctx.workbook.readOnly) throw new Error("This workbook is read-only. Save an editable .xlsx copy using File > Save As, then reopen that copy.");
    if (/\.(csv|tsv|txt|xls)$/i.test(ctx.workbook.name)) throw new Error("Save this file as an Excel Workbook (.xlsx) using File > Save As first. CSV/text and old .xls formats are not supported for editable trees.");
    if (!pending.isNullObject) throw new Error("A drawing recovery is pending. Click Load active tree to recover it before saving.");
    stage = "saving tree inputs and the editor session";
    await capture(ctx, sheet, model); await writeProbabilityCells(ctx, sheet, model);
    const nodeId = resumeNode(model, { version: 1, sheetId: sheet.id, nodeId: editingId });
    ctx.workbook.settings.add(KEY + sheet.id, JSON.stringify(model)); rememberEditor(ctx, sheet.id, nodeId); await ctx.sync();
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.11")) {
      return "Editor checkpoint updated in this workbook. Press Cmd+S to save the file before closing Excel. Reopen the .xlsx and open Macademia — Decision Trees to resume.";
    }
    stage = "asking Excel to save the workbook";
    ctx.workbook.save(Excel.SaveBehavior.prompt); await ctx.sync();
    // A cancelled Save As must never be reported as a successful save.
    ctx.workbook.load("isDirty,previouslySaved"); await ctx.sync();
    if (!ctx.workbook.previouslySaved || ctx.workbook.isDirty) return "Editor checkpoint is ready, but Excel still reports unsaved changes. Finish Save As or press Cmd+S before closing. Use .xlsx.";
    return "Workbook saved. Reopen this .xlsx, then open Macademia — Decision Trees: your selected node, view, collapsed branches and formulas will be restored. An internet connection is required to open the hosted editor.";
  });
}

async function create(sample         )                {
  if (!requireApplied()) return;
  await perform(async ctx => {
    stage = "creating a new tree sheet";
    const m = makeTree(sample), sheet = ctx.workbook.worksheets.add(); sheet.activate(); await draw(ctx, sheet, m);
    return sample ? "Sample created. Select a node to edit. Save with Cmd+S." : "Tree created. Select Option A or B, enter a payoff, and click Save node. Save with Cmd+S.";
  });
}
async function interruptedV8Canvas(ctx                      , sheet                 , m       )                   {
  // Recognize the exact pre-patch failure: the v8 header was written, then the
  // first merged formula failed before ANY node input was written.
  const heading = sheet.getRange("C2"); heading.load("values");
  const reads = m.nodes.map(n => {
    const label = sheet.getRange(addr(n, -2)), payoff = sheet.getRange(addr(n, 2, true));
    label.load("values"); payoff.load("values"); return { label, payoff };
  });
  await ctx.sync();
  return heading.values[0][0] === "Decision analysis" && reads.every(r => r.label.values[0][0] === "" && r.payoff.values[0][0] === "");
}
async function load(useCell = false, quiet = false)                {
  if (!requireApplied()) return;
  await perform(async ctx => {
    const { sheet, model } = await active(ctx);
    const pending = ctx.workbook.settings.getItemOrNullObject(KEY + sheet.id + ".recovery"); pending.load("value");
    const sessionSetting = ctx.workbook.settings.getItemOrNullObject(KEY + sheet.id + ".session"); sessionSetting.load("value");
    await ctx.sync();
    const session = sessionSetting.isNullObject ? null : parseEditorSession(sessionSetting.value);
    if (!pending.isNullObject) {
      const saved = JSON.parse(String(pending.value))                                     ;
      await draw(ctx, sheet, saved.model, saved.previous, resumeNode(saved.model, session));
      return "Interrupted drawing recovered from its saved input checkpoint. Save with Cmd+S.";
    }
    if (await interruptedV8Canvas(ctx, sheet, model)) {
      await draw(ctx, sheet, model, clone(model));
      return "Tree restored from the workbook's last saved tree model after the v8 drawing error. Review the inputs, then save with Cmd+S.";
    }
    const upgrading = model.probabilityMode !== "shared-auto" || model.layoutVersion !== 8 || model.rendererVersion !== 9;
    await capture(ctx, sheet, model);
    const oldLayout = clone(model);
    let id = currentSheet === sheet.id ? el                   ("node").value : resumeNode(model, session);
    if (useCell) {
      const selection = ctx.workbook.getSelectedRange(); selection.load("rowIndex,columnIndex"); await ctx.sync();
      const matches = model.nodes.filter(n => !n.hidden && selection.rowIndex >= n.row - 3 && selection.rowIndex <= n.row + 3
        && selection.columnIndex >= n.col && selection.columnIndex <= n.col + 4);
      if (matches.length !== 1) throw new Error("Select a node's label, input, or nearby cell, then click Use selected cell.");
      id = matches[0].id;
    }
    if (upgrading) {
      await draw(ctx, sheet, model, oldLayout, id);
      return "Tree upgraded. Card editing, branch tools, collapse controls and sensitivity analysis are ready.";
    }
    await writeProbabilityCells(ctx, sheet, model);
    const restoring = currentSheet !== sheet.id;
    rememberEditor(ctx, sheet.id, resumeNode(model, { version: 1, sheetId: sheet.id, nodeId: id }), currentSheet === sheet.id ? undefined : session?.reportTitle);
    ctx.workbook.settings.add(KEY + sheet.id, JSON.stringify(model)); await ctx.sync(); populate(model, sheet.id, id);
    if (restoring && session?.reportTitle) input("report-title").value = session.reportTitle;
    el("resume-last").classList.add("hidden");
    await markSelection(ctx, sheet, model, chosen(model, model.nodes.some(n => n.id === id) ? id : root(model).id).id);
    await refreshProbabilityAlerts(ctx, sheet, model); await watchSheet(ctx, sheet);
    return quiet ? "" : "Tree loaded. Select any card to edit it. Fixed probabilities stay entered; Auto shares the remainder.";
  }, quiet);
}
                                                 
                                                                  
                                                                                                                     
                                                                       
                                                                                                                   
                             
                                                                                                                         
                                                                                                      
 
function uniqueCuts(values          )           {
  const result           = [];
  for (const value of values.sort((a, b) => a - b)) {
    const previous = result[result.length - 1];
    if (!result.length || Math.abs(value - previous) > 1e-11 * Math.max(1, Math.abs(value), Math.abs(previous))) result.push(value);
  }
  return result;
}
function pieceAt(curve           , x        )          {
  const part = curve.find(p => x >= p.lo && x <= p.hi);
  if (!part) throw new Error("The sensitivity value is outside the calculated domain.");
  return part;
}
function curveValue(curve           , x        )         { const p = pieceAt(curve, x); return p.a * x + p.b; }
function sameWinners(a          , b          )          { return [...a].sort().join("|") === [...b].sort().join("|"); }
function curveCuts(curves             , lo        , hi        )           {
  const cuts = [lo, hi];
  for (const curve of curves) for (const piece of curve) {
    if (piece.lo > lo && piece.lo < hi) cuts.push(piece.lo);
    if (piece.hi > lo && piece.hi < hi) cuts.push(piece.hi);
  }
  return uniqueCuts(cuts);
}
function bestIds(values                                 , objective               )           {
  if (values.some(v => !Number.isFinite(v.value))) throw new Error("A sensitivity result is not a finite number. Check the input values and range.");
  const best = (objective === "max" ? Math.max : Math.min)(...values.map(v => v.value));
  const tolerance = 64 * Number.EPSILON * Math.max(1, ...values.map(v => Math.abs(v.value)));
  return values.filter(v => Math.abs(v.value - best) <= tolerance).map(v => v.id);
}
function probabilitySetup(m       , selected          )                                                                               {
  const parent = m.nodes.find(n => n.id === selected.parent);
  if (!parent || parent.kind !== "chance") throw new Error("Choose a branch leaving a chance node to test its probability.");
  const others = children(m, parent.id).filter(n => n.id !== selected.id);
  const automatic = others.filter(n => !n.probabilityFixed);
  if (!automatic.length) throw new Error("This test needs another Auto branch to absorb the remainder. Set a different branch in this chance group to Auto, apply that edit, then rerun. Other Fixed entries will not be changed.");
  const fixed = others.filter(n => n.probabilityFixed);
  if (fixed.some(n => typeof n.probability !== "number")) throw new Error("Read the evaluated Excel inputs before running sensitivity analysis.");
  const fixedTotal = fixed.reduce((sum, n) => sum + Number(n.probability), 0);
  if (fixedTotal < -1e-9 || fixedTotal > 1 + 1e-9) throw new Error("Other Fixed probabilities must total between 0% and 100%.");
  return { parent, automatic, fixedTotal, max: Math.max(0, 1 - fixedTotal) };
}
function analyzeSensitivity(m       , request                    )                    {
  const decision = chosen(m, request.decisionId), selected = chosen(m, request.inputId);
  const options = children(m, decision.id);
  if (decision.kind !== "decision" || options.length < 2) throw new Error("Choose a decision with at least two alternatives.");
  if (![request.lo, request.hi].every(Number.isFinite) || request.hi <= request.lo) throw new Error("The upper test value must be greater than the lower value.");
  if (m.nodes.some(n => typeof n.payoff !== "number" || !Number.isFinite(n.payoff))) throw new Error("Sensitivity needs the current evaluated payoffs from Excel.");
  for (const parent of m.nodes.filter(n => n.kind === "chance")) {
    const list = children(m, parent.id);
    if (!list.length || list.some(n => typeof n.probability !== "number" || !Number.isFinite(n.probability) || n.probability < 0 || n.probability > 1)
      || Math.abs(list.reduce((sum, n) => sum + Number(n.probability), 0) - 1) > 1e-6) throw new Error(`Correct probabilities at ${parent.label} before running sensitivity.`);
  }
  let lo = request.lo, hi = request.hi, setup                                                 ;
  const currentValue = Number(selected[request.field]);
  if (request.field === "probability") {
    if (lo < 0 || hi > 1) throw new Error("Probability test values must be between 0% and 100%.");
    setup = probabilitySetup(m, selected); hi = Math.min(hi, setup.max);
    if (hi <= lo) throw new Error(`There is no feasible interval above the lower value. Other Fixed entries limit this probability to ${probabilityText(setup.max)}.`);
  }
  const lower = Math.min(lo, currentValue), upper = Math.max(hi, currentValue);
  const constant = (value        )            => [{ lo: lower, hi: upper, a: 0, b: value }];
  const ownCurve = (n          )            => request.field === "payoff" && n.id === selected.id
    ? [{ lo: lower, hi: upper, a: 1, b: 0 }] : constant(Number(n.payoff));
  const memo = new Map                   ();
  function crossings(curves             , from        , to        )           {
    const base = curveCuts(curves, from, to), cuts = [...base];
    for (let k = 0; k < base.length - 1; k++) {
      const left = base[k], right = base[k + 1], mid = left + (right - left) / 2;
      const lines = curves.map(c => pieceAt(c, mid));
      for (let i = 0; i < lines.length; i++) for (let j = i + 1; j < lines.length; j++) {
        const denominator = lines[i].a - lines[j].a;
        if (denominator === 0) continue;
        const x = (lines[j].b - lines[i].b) / denominator;
        if (x > left && x < right && Number.isFinite(x)) cuts.push(x);
      }
    }
    return uniqueCuts(cuts);
  }
  function combine(curves             , weights                            , own           )            {
    const cuts = curveCuts([...curves, own], lower, upper), result            = [];
    for (let i = 0; i < cuts.length - 1; i++) {
      const x = cuts[i] + (cuts[i + 1] - cuts[i]) / 2, cost = pieceAt(own, x);
      let a = cost.a, b = cost.b;
      curves.forEach((c, j) => {
        const line = pieceAt(c, x), weight = weights[j];
        if (weight.a !== 0 && line.a !== 0) throw new Error("This input creates a nonlinear dependency that this version cannot analyze.");
        a += weight.a * line.b + weight.b * line.a; b += weight.b * line.b;
      });
      result.push({ lo: cuts[i], hi: cuts[i + 1], a, b });
    }
    return result;
  }
  function rollback(n          )            {
    const cached = memo.get(n.id); if (cached) return cached;
    const list = children(m, n.id), own = ownCurve(n);
    let curve           ;
    if (!list.length) {
      if (n.kind !== "outcome") throw new Error(`${n.label} has no branches. Complete the tree before running sensitivity.`);
      curve = own;
    } else if (n.kind === "decision") {
      const curves = list.map(rollback), cuts = crossings(curves, lower, upper); curve = [];
      for (let i = 0; i < cuts.length - 1; i++) {
        const x = cuts[i] + (cuts[i + 1] - cuts[i]) / 2;
        const values = curves.map(c => curveValue(c, x)), best = (m.objective === "max" ? Math.max : Math.min)(...values);
        const line = pieceAt(curves[values.indexOf(best)], x), cost = pieceAt(own, x);
        curve.push({ lo: cuts[i], hi: cuts[i + 1], a: line.a + cost.a, b: line.b + cost.b });
      }
    } else {
      const weights = list.map(c => {
        if (setup && n.id === setup.parent.id) {
          if (c.id === selected.id) return { a: 1, b: 0 };
          if (setup.automatic.some(auto => auto.id === c.id)) return { a: -1 / setup.automatic.length, b: (1 - setup.fixedTotal) / setup.automatic.length };
        }
        return { a: 0, b: Number(c.probability) };
      });
      curve = combine(list.map(rollback), weights, own);
    }
    memo.set(n.id, curve); return curve;
  }
  const curves = options.map(n => combine([rollback(n)], [{ a: 0, b: 1 }], ownCurve(decision)));
  const at = (x        ) => options.map((n, i) => ({ id: n.id, value: curveValue(curves[i], x) }));
  const cuts = crossings(curves, lo, hi), bands                    = [];
  for (let i = 0; i < cuts.length - 1; i++) {
    const winners = bestIds(at(cuts[i] + (cuts[i + 1] - cuts[i]) / 2), m.objective), last = bands[bands.length - 1];
    if (last && sameWinners(last.winners, winners)) last.hi = cuts[i + 1];
    else bands.push({ lo: cuts[i], hi: cuts[i + 1], winners });
  }
  const boundaries                        = [];
  for (let i = 0; i < cuts.length; i++) {
    const x = cuts[i], winners = bestIds(at(x), m.objective);
    const below = i ? bestIds(at(cuts[i - 1] + (x - cuts[i - 1]) / 2), m.objective) : [];
    const above = i < cuts.length - 1 ? bestIds(at(x + (cuts[i + 1] - x) / 2), m.objective) : [];
    if ((below.length && !sameWinners(winners, below)) || (above.length && !sameWinners(winners, above))) {
      boundaries.push({ x, winners, below, above, switches: !!below.length && !!above.length && !sameWinners(below, above) });
    }
  }
  const currentValues = at(currentValue);
  return { request: { ...request, lo, hi }, current: currentValue, currentWinners: bestIds(currentValues, m.objective), currentValues,
    bands, boundaries, feasibleMax: setup?.max, clipped: hi !== request.hi };
}
async function sensitivitySnapshot(ctx                      , sheet                 , source       )                 {
  const m = clone(source); ensureRemainders(m);
  const reads = m.nodes.map(n => {
    const label = sheet.getRange(addr(n, -2)), payoff = sheet.getRange(addr(n, 2, true));
    const probability = m.nodes.find(p => p.id === n.parent)?.kind === "chance" ? sheet.getRange(addr(n, 1, true)) : null;
    label.load("values"); payoff.load("values,formulas"); probability?.load("values,formulas"); return { n, label, payoff, probability };
  });
  await ctx.sync();
  const overwritten = reads.filter(r => r.probability && r.n.remainder
    && normalizedFormula(r.probability.formulas[0][0]) !== normalizedFormula(remainderFormula(m, r.n)));
  overwritten.forEach(r => { r.n.probabilityFixed = true; r.n.remainder = false; });
  for (const r of reads) {
    r.n.label = String(r.label.values[0][0] || r.n.label);
    const raw = r.payoff.values[0][0], formula = r.payoff.formulas[0][0];
    const value = (raw === "" || raw === null || raw === undefined) && !(typeof formula === "string" && formula.startsWith("=")) ? 0 : raw;
    if (typeof value !== "number" || !Number.isFinite(value)) throw new Error(`Correct the payoff at ${r.n.label} (${r.n.id}, ${addr(r.n, 2, true)}). Enter a number or fix its worksheet formula.`);
    r.n.payoff = value;
    if (r.probability && r.n.probabilityFixed) {
      const p = r.probability.values[0][0];
      if (typeof p !== "number" || !Number.isFinite(p) || p < 0 || p > 1) throw new Error(`Correct the Fixed probability at ${r.n.label} (${r.n.id}, ${addr(r.n, 1, true)}). Enter a value between 0% and 100%.`);
      r.n.probability = p;
    }
  }
  for (const parent of m.nodes.filter(n => n.kind === "chance")) {
    const list = children(m, parent.id), auto = list.filter(n => !n.probabilityFixed);
    const total = list.filter(n => n.probabilityFixed).reduce((sum, n) => sum + Number(n.probability), 0);
    if (total > 1 + 1e-9) throw new Error(`Fixed probabilities exceed 100% at ${parent.label}.`);
    for (const n of auto) n.probability = Math.max(0, 1 - total) / auto.length;
  }
  return m;
}
let sensitivityDefaults = "";
let sensitivitySource = "";
function sensitivitySignature(m       )         {
  return JSON.stringify([m.objective, m.nodes.map(n => [n.id, n.parent, n.kind, n.label, n.payoff, n.probabilityFixed, n.probability])]);
}
function refreshSensitivity(m       )       {
  if (sensitivitySource && sensitivitySource !== sensitivitySignature(m)) {
    el("sensitivity-result").textContent = "Tree inputs changed. Rerun the what-if analysis for updated results.";
    sensitivitySource = "";
  }
  function fill(id        , list            , preferred        ) {
    const select = el                   (id), before = select.value; select.replaceChildren();
    for (const n of list) { const option = document.createElement("option"); option.value = n.id; option.textContent = `${n.label} (${n.id})${n.hidden ? " · collapsed" : ""}`; select.appendChild(option); }
    select.value = list.some(n => n.id === before) ? before : list.some(n => n.id === preferred) ? preferred : list[0]?.id || "";
  }
  fill("sensitivity-decision", m.nodes.filter(n => n.kind === "decision" && children(m, n.id).length > 1), root(m).id);
  const field = el                   ("sensitivity-field").value                    ;
  const candidates = m.nodes.filter(n => field === "payoff" || m.nodes.find(p => p.id === n.parent)?.kind === "chance");
  let preferred = editingId;
  if (field === "payoff" && preferred === root(m).id) preferred = children(m, preferred)[0]?.id || preferred;
  if (field === "probability" && !candidates.some(n => n.id === preferred)) preferred = candidates.find(n => n.parent === editingId)?.id || candidates[0]?.id || "";
  fill("sensitivity-input", candidates, preferred); sensitivityRangeDefaults();
}
function sensitivityRangeDefaults(force = false)       {
  if (!current) return;
  const field = el                   ("sensitivity-field").value                    ;
  const n = current.nodes.find(n => n.id === el                   ("sensitivity-input").value);
  const key = `${currentSheet}|${field}|${n?.id || ""}`;
  if (!force && key === sensitivityDefaults) return;
  sensitivityDefaults = key;
  if (field === "probability") { input("sensitivity-low").value = "0%"; input("sensitivity-high").value = "100%"; }
  else { const value = typeof n?.payoff === "number" ? n.payoff : 0, span = Math.max(100000, Math.abs(value)); input("sensitivity-low").value = String(value - span); input("sensitivity-high").value = String(value + span); }
  el("sensitivity-result").replaceChildren();
}
function sensitivityNumber(text        , field                  )         {
  if (!text.trim()) throw new Error("Enter both ends of the test range.");
  const cleaned = text.replace(/[$,\s]/g, "");
  if (field === "payoff" && !cleaned) throw new Error("Enter a numeric payoff range.");
  const value = field === "probability" ? probabilityExpr(text) : Number(cleaned);
  if (typeof value !== "number" || !Number.isFinite(value)) throw new Error("Test ranges must be numbers; use percentages for probabilities.");
  return value;
}
function sensitivityFormat(value        , field                  )         {
  return field === "probability" ? probabilityText(value) : new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 2 }).format(value);
}
function showSensitivity(m       , result                   )       {
  const host = el("sensitivity-result"); host.replaceChildren();
  const name = (id        ) => `${chosen(m, id).label} (${id})`;
  const names = (ids          ) => ids.map(name).join(" / ");
  const format = (n        ) => sensitivityFormat(n, result.request.field);
  function p(text        , strong = false) { const item = document.createElement("p"); item.textContent = text; if (strong) item.className = "sensitivity-headline"; host.appendChild(item); }
  p("Snapshot results — rerun after changing any input or formula assumption.");
  p(`Current best: ${names(result.currentWinners)}${result.currentWinners.length > 1 ? " — tied" : ""}`, true);
  p(`Decision: ${name(result.request.decisionId)}. Input: ${name(result.request.inputId)} · ${result.request.field}.`);
  const decision = chosen(m, result.request.decisionId);
  if (result.request.field === "payoff" && result.request.inputId === decision.id) p("This payoff is shared by every alternative at this decision, so it changes their totals equally.");
  else if (!descendants(m, decision.id).some(n => n.id === result.request.inputId)) p("This input is outside the selected decision's branches and cannot change this decision. Choose a different decision or input to test.");
  p(`Current input: ${format(result.current)}. Tested: ${format(result.request.lo)} to ${format(result.request.hi)}.`);
  if (result.clipped) p(`Other Fixed entries limit this probability to ${format(result.feasibleMax )}; the tested range was capped there.`);
  for (const option of result.currentValues) p(`${name(option.id)}: ${sensitivityFormat(option.value, "payoff")} expected value now.`);
  if (!result.boundaries.length) p("No decision change or new break-even point within this tested range.", true);
  else {
    const nearest = [...result.boundaries].sort((a, b) => Math.abs(a.x - result.current) - Math.abs(b.x - result.current))[0];
    const delta = nearest.x - result.current;
    const distance = result.request.field === "probability" ? `${Number((Math.abs(delta) * 100).toFixed(6))} percentage points` : sensitivityFormat(Math.abs(delta), "payoff");
    p(`Nearest break-even: ≈ ${format(nearest.x)} (${Math.abs(delta) < 1e-10 ? "at the current input" : `${delta < 0 ? "decrease" : "increase"} of ${distance}`}).`, true);
    for (const boundary of result.boundaries) {
      p(`≈ ${format(boundary.x)}: ${names(boundary.winners)} tie${boundary.switches ? "; preferred choice changes across this point" : "; tie or range-end boundary"}.`);
      if (boundary.below.length) p(`Just below: ${names(boundary.below)}.`);
      if (boundary.above.length) p(`Just above: ${names(boundary.above)}.`);
    }
  }
  const table = document.createElement("table"), header = table.createTHead().insertRow();
  for (const title of ["Input range", "Best choice"]) { const th = document.createElement("th"); th.scope = "col"; th.textContent = title; header.appendChild(th); }
  const body = table.createTBody();
  for (const band of result.bands) { const row = body.insertRow(); row.insertCell().textContent = `${format(band.lo)} – ${format(band.hi)}`; row.insertCell().textContent = names(band.winners) + (band.winners.length > 1 ? " (tie)" : ""); }
  host.appendChild(table);
  p("Ranges describe values between boundary points; ties at the boundaries are listed separately. Displayed thresholds are rounded.");
  p("Snapshot only. Other input/formula values are held at their current evaluated numbers; other Auto branches in the tested chance group share the remainder. Nested decisions are re-optimized. Rerun after changing assumptions. Source inputs were not changed.");
}
async function runSensitivity()                {
  if (!requireApplied()) return;
  const request                     = { decisionId: el                   ("sensitivity-decision").value, inputId: el                   ("sensitivity-input").value,
    field: el                   ("sensitivity-field").value                    , lo: 0, hi: 0 };
  try { request.lo = sensitivityNumber(input("sensitivity-low").value, request.field); request.hi = sensitivityNumber(input("sensitivity-high").value, request.field); }
  catch (e) { el("status").textContent = (e         ).message; el("sensitivity-result").textContent = (e         ).message; return; }
  await perform(async ctx => {
    try {
    const { sheet, model } = await active(ctx);
    if (sheet.id !== currentSheet) throw new Error("Load this tree before running sensitivity analysis.");
    stage = "reading a sensitivity snapshot"; const snapshot = await sensitivitySnapshot(ctx, sheet, model);
    stage = "calculating decision break-even points"; const result = analyzeSensitivity(snapshot, request);
    sensitivitySource = sensitivitySignature(model);
    showSensitivity(snapshot, result); return "Sensitivity analysis complete. The tree's source inputs were not changed.";
    } catch (error) { el("sensitivity-result").textContent = (error         ).message || String(error); throw error; }
  });
}

// Release 1.0a: local-only reports. No network upload, formula rewrite or canvas mutation.
                                                                                            
                                                                                                         
                                                                                                                                           
let reportUrl = "";
function invalidateReport(message        )       {
  if (!reportUrl) return;
  URL.revokeObjectURL(reportUrl); reportUrl = "";
  el("report-result").textContent = message;
}
function reportValues(m       )                      {
  const values = new Map                (), visiting = new Set        ();
  function evaluate(n          )         {
    if (values.has(n.id)) return values.get(n.id) ;
    if (visiting.has(n.id)) throw new Error(`Circular branch at ${n.label} (${n.id}).`);
    visiting.add(n.id);
    if (typeof n.payoff !== "number" || !Number.isFinite(n.payoff)) throw new Error(`Correct the payoff at ${n.label} (${n.id}, ${addr(n, 2, true)}).`);
    const list = children(m, n.id);
    if (n.kind !== "outcome" && !list.length) throw new Error(`Add branches to ${n.label} (${n.id}), or change it to an Outcome.`);
    if (n.kind === "outcome" && list.length) throw new Error(`Outcome ${n.label} (${n.id}) has child branches. Change its type first.`);
    const downstream = list.map(evaluate);
    let result = n.payoff;
    if (n.kind === "decision") result += m.objective === "max" ? Math.max(...downstream) : Math.min(...downstream);
    if (n.kind === "chance") {
      const ps = list.map(c => c.probability);
      if (ps.some(p => typeof p !== "number" || !Number.isFinite(p) || p < 0 || p > 1)
        || Math.abs(ps.reduce        ((sum, p) => sum + Number(p), 0) - 1) >= 1e-6)
        throw new Error(`Probabilities at ${n.label} (${n.id}) must total 100%. Correct the Fixed entries or switch a branch to Auto.`);
      result += downstream.reduce((sum, value, i) => sum + Number(ps[i]) * value, 0);
    }
    if (!Number.isFinite(result)) throw new Error(`The calculated value at ${n.label} is too large. Check its inputs.`);
    visiting.delete(n.id); values.set(n.id, result); return result;
  }
  evaluate(root(m));
  if (values.size !== m.nodes.length) throw new Error("The tree contains disconnected nodes. Reload it before exporting.");
  return values;
}
function reportPlan(source       , overview         )             {
  // Evaluate while original worksheet addresses are intact, then lay out a fully expanded copy.
  const values = reportValues(source), m = clone(source);
  m.nodes.forEach(n => { n.collapsed = false; }); layout(m);
  const minY = Math.min(...m.nodes.map(n => (n.row - 3) * 18));
  const cards               = m.nodes.map(n => ({ node: n, x: (n.col - 1) * 48, y: (n.row - 3) * 18 - minY,
    width: 240, height: n.kind === "outcome" ? 72 : 108 }));
  const width = Math.max(...cards.map(c => c.x + c.width)), height = Math.max(...cards.map(c => c.y + c.height));
  const tiles               = [];
  if (overview) tiles.push({ x: 0, y: 0, width, height, column: 1, row: 1 });
  else {
    // Each horizontal strip contains two depth columns. Cuts lie only in gutters.
    for (let x = 0, column = 1; x < width; x += 672, column++) {
      const endX = Math.min(x + 624, width), strip = cards.filter(c => c.x >= x && c.x < endX);
      const intervals = strip.map(c => ({ start: c.y, end: c.y + c.height })).sort((a, b) => a.start - b.start);
      const merged                                   = [];
      for (const interval of intervals) {
        const last = merged[merged.length - 1];
        if (last && interval.start <= last.end + 12) last.end = Math.max(last.end, interval.end);
        else merged.push({ ...interval });
      }
      let y = 0, row = 1;
      while (y < height) {
        let end = Math.min(y + 440, height);
        const crossed = merged.find(i => i.start < end && i.end > end);
        if (crossed) end = crossed.start > y + 12 ? crossed.start - 6 : crossed.end + 6;
        end = Math.min(end, height);
        if (end <= y) throw new Error("Could not find a safe page boundary.");
        // Each strip is independently banded; node IDs identify cross-page connections.
        // Blank sections are omitted.
        const intersects = strip.some(c => c.y < end && c.y + c.height > y);
        if (intersects) tiles.push({ x, y, width: endX - x, height: end - y, column, row });
        y = end; row++;
      }
    }
  }
  return { model: m, cards, values, width, height, tiles };
}
function reportMoney(value        )         {
  const absolute = Math.abs(value).toLocaleString("en-US", { maximumFractionDigits: 0 });
  return value < 0 ? `($${absolute})` : `$${absolute}`;
}
function reportBest(m       , values                     , parent          )           {
  const list = children(m, parent.id); if (!list.length) return [];
  const best = m.objective === "max" ? Math.max(...list.map(n => values.get(n.id) )) : Math.min(...list.map(n => values.get(n.id) ));
  // Match the worksheet's exact-equality policy markers.
  return list.filter(n => values.get(n.id) === best).map(n => n.id);
}
function canvasText(ctx                          , text        , x        , y        , width        , maxLines        , size        , bold = false, centered = false)       {
  ctx.font = `${bold ? "600 " : ""}${size}px Arial, sans-serif`; ctx.fillStyle = INK;
  const words = text.replace(/\s+/g, " ").trim().split(" "), lines           = []; let line = "";
  for (const word of words) {
    const proposed = line ? line + " " + word : word;
    if (ctx.measureText(proposed).width <= width) { line = proposed; continue; }
    if (line) { lines.push(line); line = ""; }
    // Only split words that cannot fit by themselves (e.g. a long identifier).
    for (const char of Array.from(word)) {
      if (line && ctx.measureText(line + char).width > width) { lines.push(line); line = char; } else line += char;
    }
  }
  if (line) lines.push(line);
  if (lines.length > maxLines) {
    lines.length = maxLines; let last = lines[maxLines - 1];
    while (last && ctx.measureText(last + "…").width > width) last = last.slice(0, -1);
    lines[maxLines - 1] = last + "…";
  }
  ctx.textAlign = centered ? "center" : "left"; ctx.textBaseline = "middle";
  const step = size * 1.2;
  lines.forEach((value, i) => ctx.fillText(value, centered ? x + width / 2 : x, y + (i - (lines.length - 1) / 2) * step));
}
function reportCanvas(plan            , tile            , title        , index        , date        , overview         )                    {
  const pageWidth = overview ? Math.max(792, tile.width + 48) : 792;
  const pageHeight = overview ? Math.max(612, tile.height + 160) : 612;
  // Bound both memory and the longest canvas dimension, even for the 60-node limit.
  const density = Math.min(3, 16000 / Math.max(pageWidth, pageHeight), Math.sqrt(24000000 / (pageWidth * pageHeight)));
  const canvas = document.createElement("canvas"); canvas.width = Math.ceil(pageWidth * density); canvas.height = Math.ceil(pageHeight * density);
  const ctx = canvas.getContext("2d"); if (!ctx) throw new Error("This Excel web view cannot create a PDF image. Use Prepare Excel print instead.");
  ctx.scale(density, density); ctx.fillStyle = PAPER; ctx.fillRect(0, 0, pageWidth, pageHeight);
  canvasText(ctx, title, 24, 34, pageWidth - 48, 2, 18, true);
  canvasText(ctx, `${plan.model.objective === "max" ? "Maximize expected net payoff" : "Minimize expected value / cost"}  ·  Tree value ${reportMoney(plan.values.get(root(plan.model).id) )}`, 24, 66, pageWidth - 48, 1, 11);
  canvasText(ctx, "■ Decision    ● Chance    ▲ Outcome    |    BEST / TIED BEST: preferred decision alternatives", 24, 84, pageWidth - 48, 1, 10);
  if (!overview) {
    const here = new Set(plan.cards.filter(c => c.x >= tile.x && c.x < tile.x + tile.width && c.y >= tile.y && c.y < tile.y + tile.height).map(c => c.node.id));
    const references = plan.cards.filter(c => here.has(c.node.id) && c.node.parent && !here.has(c.node.parent)).map(c => `${c.node.id} ← ${c.node.parent}`);
    canvasText(ctx, references.length ? "Continued connections: " + references.join("; ") : "Start of this tree section", 24, 104, pageWidth - 48, 2, 9);
  }
  const scale = Math.min(1, (pageWidth - 48) / tile.width, (pageHeight - 160) / tile.height);
  ctx.save(); ctx.beginPath(); ctx.rect(24, 124, pageWidth - 48, pageHeight - 160); ctx.clip();
  ctx.translate(24, 124); ctx.scale(scale, scale); ctx.translate(-tile.x, -tile.y);
  ctx.lineWidth = 1; ctx.strokeStyle = "#4C4D50";
  const byId = new Map(plan.cards.map(c => [c.node.id, c]));
  for (const parent of plan.cards) {
    const list = children(plan.model, parent.node.id).map(n => byId.get(n.id) ); if (!list.length) continue;
    const fromX = parent.x + parent.width, fromY = parent.y + 18, bus = fromX + 48;
    ctx.beginPath(); ctx.moveTo(fromX, fromY); ctx.lineTo(bus, fromY);
    ctx.moveTo(bus, Math.min(fromY, ...list.map(c => c.y + 18))); ctx.lineTo(bus, Math.max(fromY, ...list.map(c => c.y + 18)));
    for (const child of list) { ctx.moveTo(bus, child.y + 18); ctx.lineTo(child.x, child.y + 18); } ctx.stroke();
  }
  for (const c of plan.cards) {
    if (c.x + c.width < tile.x || c.x > tile.x + tile.width || c.y + c.height < tile.y || c.y > tile.y + tile.height) continue;
    const n = c.node, parent = plan.model.nodes.find(p => p.id === n.parent), best = parent?.kind === "decision" ? reportBest(plan.model, plan.values, parent) : [];
    ctx.fillStyle = CARD; ctx.fillRect(c.x, c.y, c.width, c.height); ctx.fillStyle = "#F0EEEA"; ctx.fillRect(c.x, c.y, c.width, 36);
    ctx.strokeStyle = "#8B8985"; ctx.strokeRect(c.x, c.y, c.width, c.height); ctx.beginPath();ctx.moveTo(c.x, c.y + 36);ctx.lineTo(c.x + c.width, c.y + 36);ctx.stroke();
    ctx.fillStyle = COLORS[n.kind]; ctx.beginPath(); const sx = c.x + 18, sy = c.y + 18;
    if (n.kind === "chance") ctx.arc(sx, sy, 6, 0, Math.PI * 2);
    else if (n.kind === "decision") ctx.rect(sx - 6, sy - 6, 12, 12);
    else {ctx.moveTo(sx, sy - 7);ctx.lineTo(sx + 7, sy + 6);ctx.lineTo(sx - 7, sy + 6);ctx.closePath();} ctx.fill();
    canvasText(ctx, n.id, c.x + 2, c.y + 30, 32, 1, 7, false, true);
    canvasText(ctx, n.label, c.x + 36, c.y + 18, 196, 2, 11, true, true);
    const rows                     = [];
    if (parent?.kind === "chance") rows.push(["Probability", probabilityText(n.probability)]);
    else if (best.length && plan.model.showStrategy !== false) rows.push(["Decision branch", best.includes(n.id) ? best.length > 1 ? "TIED BEST" : "BEST" : "Alternate"]);
    rows.push([n.kind === "outcome" ? "Net payoff" : "Payoff at node", reportMoney(Number(n.payoff))]);
    if (n.kind !== "outcome") rows.push(["Expected value", reportMoney(plan.values.get(n.id) )]);
    rows.forEach(([caption, value], i) => {
      const y = c.y + 45 + i * 18;
      canvasText(ctx, caption, c.x + 8, y, 138, 1, 10);
      ctx.font = "600 10px Arial, sans-serif"; ctx.textAlign = "right"; ctx.fillStyle = INK; ctx.fillText(value, c.x + 232, y, 86);
    });
    if (n.kind === "decision" && plan.model.showStrategy !== false) {
      const ids = reportBest(plan.model, plan.values, n), label = ids.length === 1 ? `Best: ${byId.get(ids[0]) .node.label}` : `${ids.length} equal best options`;
      canvasText(ctx, label, c.x + 8, c.y + 99, 224, 1, 9);
    }
  }
  ctx.restore();
  canvasText(ctx, `Macademia — Decision Trees 1.0  ·  ${date}  ·  ${plan.model.nodes.length} nodes  ·  Snapshot of evaluated inputs`, 24, pageHeight - 22, pageWidth - 220, 1, 9);
  canvasText(ctx, overview ? "Full-tree overview" : `Page ${index + 1}/${plan.tiles.length} · Map C${tile.column} / R${tile.row}`, pageWidth - 190, pageHeight - 22, 166, 1, 9);
  if (!overview) canvasText(ctx, "Connectors may continue on another map page. Node boxes are never split.", 24, pageHeight - 40, pageWidth - 48, 1, 9);
  return canvas;
}
                                                                                                              
function imagePdf(images            )             {
  // Minimal PDF 1.4 writer. All text/shapes are rasterized together to preserve
  // Unicode labels and consistent alignment across Mac fonts. No external package.
  const parts               = [], offsets           = [0]; let size = 0;
  function append(data                     ) { const bytes = typeof data === "string" ? Uint8Array.from(Array.from(data).map(c => c.charCodeAt(0))) : data; parts.push(bytes); size += bytes.length; }
  function object(id        , value        ) { offsets[id] = size; append(`${id} 0 obj\n${value}\nendobj\n`); }
  append("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n"); object(1, "<< /Type /Catalog /Pages 2 0 R >>");
  object(2, `<< /Type /Pages /Count ${images.length} /Kids [${images.map((_, i) => `${3 + i * 3} 0 R`).join(" ")}] >>`);
  images.forEach((im, i) => {
    const id = 3 + i * 3;
    object(id, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${im.pageWidth} ${im.pageHeight}] /Resources << /XObject << /Im0 ${id + 2} 0 R >> >> /Contents ${id + 1} 0 R >>`);
    const commands = `q ${im.pageWidth} 0 0 ${im.pageHeight} 0 0 cm /Im0 Do Q\n`;
    object(id + 1, `<< /Length ${commands.length} >>\nstream\n${commands}endstream`);
    offsets[id + 2] = size;
    append(`${id + 2} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${im.width} /Height ${im.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${im.bytes.length} >>\nstream\n`);
    append(im.bytes); append("\nendstream\nendobj\n");
  });
  const xref = size, count = 3 + images.length * 3;
  append(`xref\n0 ${count}\n0000000000 65535 f \n`);
  for (let i = 1; i < count; i++) append(`${String(offsets[i]).padStart(10, "0")} 00000 n \n`);
  append(`trailer\n<< /Size ${count} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
  const result = new Uint8Array(size); let cursor = 0; for (const part of parts) { result.set(part, cursor); cursor += part.length; } return result;
}
async function generateReport()                {
  if (!requireApplied()) return;
  invalidateReport("Generating a fresh snapshot…");
  el("report-result").textContent = "Reading current worksheet inputs…";
  await perform(async ctx => {
    try {
      const { sheet, model } = await active(ctx);
      if (sheet.id !== currentSheet) throw new Error("Load this tree before exporting it.");
      stage = "validating report inputs"; const snapshot = await sensitivitySnapshot(ctx, sheet, model);
      const overview = el                   ("report-format").value === "overview", plan = reportPlan(snapshot, overview);
      const title = input("report-title").value.trim() || root(snapshot).label, images             = [], date = new Date().toLocaleDateString();
      for (let i = 0; i < plan.tiles.length; i++) {
        stage = `rendering PDF page ${i + 1} of ${plan.tiles.length}`; el("report-result").textContent = stage;
        const tile = plan.tiles[i], canvas = reportCanvas(plan, tile, title, i, date, overview), raw = atob(canvas.toDataURL("image/jpeg", 0.94).split(",")[1]);
        const bytes = Uint8Array.from(Array.from(raw).map(c => c.charCodeAt(0)));
        images.push({ bytes, width: canvas.width, height: canvas.height, pageWidth: overview ? Math.max(792, tile.width + 48) : 792, pageHeight: overview ? Math.max(612, tile.height + 160) : 612 });
        canvas.width = 1; canvas.height = 1; // release each large drawing surface before the next page
        await new Promise      (resolve => { setTimeout(resolve, 0); });
      }
      const bytes = imagePdf(images), blob = new Blob([bytes            ], { type: "application/pdf" });
      if (reportUrl) URL.revokeObjectURL(reportUrl); reportUrl = URL.createObjectURL(blob);
      const link = document.createElement("a"); link.href = reportUrl; link.download = (title.replace(/[^a-zA-Z0-9 _-]/g, "").slice(0, 70) || "Decision-analysis") + ".pdf";
      link.className = "download-report"; link.textContent = `Save PDF · ${images.length} ${images.length === 1 ? "page" : "pages"}`;
      el("report-result").replaceChildren(link);
      const note = document.createElement("p"); note.textContent = "Includes the entire tree, including collapsed branches. This is a static snapshot; regenerate after changes. If Excel blocks this download, use Prepare Excel print below."; el("report-result").appendChild(note);
      return "PDF ready. Click Save PDF in Export / PDF. Source inputs and your working layout were not changed.";
    } catch (e) { el("report-result").textContent = (e         ).message; throw e; }
  });
}
async function checkTree()                {
  if (!requireApplied()) return;
  await perform(async ctx => {
    stage = "checking the complete tree"; const { sheet, model } = await active(ctx);
    if (sheet.id !== currentSheet) throw new Error("Load the active tree first.");
    const snapshot = await sensitivitySnapshot(ctx, sheet, model), values = reportValues(snapshot);
    return `Input check passed: ${snapshot.nodes.length} nodes, including collapsed branches. Tree value ${reportMoney(values.get(root(snapshot).id) )}. All chance groups total 100%. This checks evaluated inputs and tree structure; save with Cmd+S.`;
  });
}
async function prepareExcelPrint()                {
  if (!requireApplied()) return;
  await perform(async ctx => {
    const { sheet, model } = await active(ctx);
    if (sheet.id !== currentSheet) throw new Error("Load the active tree first.");
    stage = "preparing the full-tree Excel print view";
    const snapshot = await sensitivitySnapshot(ctx, sheet, model); reportValues(snapshot);
    await capture(ctx, sheet, model); const previous = clone(model);
    model.nodes.forEach(n => { n.collapsed = false; }); model.view = "presentation";
    await draw(ctx, sheet, model, previous);
    // Verified ExcelApi 1.9 properties; native File > Print remains user-controlled.
    const lastCol = Math.max(15, ...model.nodes.filter(n => !n.hidden).map(n => n.col + 4));
    const lastRow = Math.max(...model.nodes.filter(n => !n.hidden).map(n => n.row + (n.kind === "outcome" ? 1 : 3)));
    const page = sheet.pageLayout; page.setPrintArea(`A1:${column(lastCol)}${lastRow + 1}`);
    page.orientation = "Landscape"; page.printGridlines = false; page.printHeadings = false;
    page.leftMargin = 18; page.rightMargin = 18; page.topMargin = 18; page.bottomMargin = 18;
    page.centerHorizontally = true; page.zoom = { horizontalFitToPages: 1, verticalFitToPages: 1 }; await ctx.sync();
    return "Full tree prepared in Presentation view. In Excel choose File > Print, print the active sheet, then PDF > Save as PDF. It fits one page; large trees may look small. Undo last edit restores the previous tree view (print settings remain).";
  });
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) return;
  document.body.style.cssText = "margin:0;background:#e7e4df;color:#25282d;display:block;";
  // Logo uses standard circle/square/triangle geometry, matching the native Excel node symbols.
  const logo = "data:image/svg+xml," + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="8" r="6" fill="#4776CC" stroke="#25282D"/><rect x="3" y="19" width="11" height="11" fill="#29955B" stroke="#25282D"/><path d="M24 18 L31 30 L17 30 Z" fill="#F3C84B" stroke="#25282D"/></svg>');
  document.body.innerHTML = `
    <style>
      *{box-sizing:border-box}body{font:14px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
      main{padding:0 16px 20px;max-width:620px;margin:auto}h1{font-size:19px;margin:0;letter-spacing:-.4px}
      .brand{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid #aaa69f;box-shadow:0 1px #fff}
      .product-name{font-size:12px;letter-spacing:.2px;color:#575a60}.brand img{width:30px;height:30px}.version{margin-left:auto;font-size:11px;color:#555}
      h2{font-size:16px;margin:0 0 12px}p{margin:8px 0}button,input,select{font:inherit;color:inherit}
      button{padding:9px 12px;border:1px solid #77746f;border-radius:3px;background:#e7e4df;box-shadow:inset 1px 1px #fff,inset -1px -1px #b0ada7;cursor:pointer}
      button:hover{background:#f4f1eb}button:active{box-shadow:inset 1px 1px #918d85}
      button.primary{background:#365aa5;color:white;border-color:#243f79;box-shadow:inset 1px 1px #879bc9,inset -1px -1px #233e78;font-weight:650}
      button.primary:hover{background:#2d4f98}button:disabled{opacity:.55;cursor:wait}
      button:focus-visible,input:focus-visible,select:focus-visible,summary:focus-visible{outline:2px solid #365aa5;outline-offset:2px}
      input,select{width:100%;min-width:0;border:1px solid #99958d;border-radius:2px;background:#fffefa;padding:8px;box-shadow:inset 1px 1px #d8d4cc}
      input:disabled,select:disabled{background:#e2dfd9;color:#666}label{display:block;font-weight:550;margin:10px 0 4px}
      .check{display:flex;align-items:flex-start;gap:8px;font-weight:400}.check input{width:auto;margin-top:4px}
      .node-picker{display:flex;gap:8px;align-items:center}.node-picker select{flex:1}.node-picker button{flex:none;padding:8px}.field-pair{display:grid;grid-template-columns:1fr 1fr;gap:10px}
      .row{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}.row button{flex:1}
      .section{padding:12px 0;border-top:1px solid #b4b0a8;box-shadow:0 -1px #fdfbf7}
      .titlebar{margin:15px -16px 0;padding:10px 16px;background:#586e9b;color:#fff;font-size:17px;font-weight:650;overflow-wrap:anywhere}
      .segmented{display:flex;margin:14px 0}.segmented button{width:50%;border-radius:0}.segmented button:first-child{border-radius:3px 0 0 3px}.segmented button:last-child{border-radius:0 3px 3px 0}
      .segmented button[aria-pressed="true"]{background:#c2cad9;box-shadow:inset 1px 1px #8c97ac;font-weight:650}
      .download-report{display:block;padding:12px;background:#365aa5;color:#fff;text-align:center;font-weight:650;border:1px solid #243f79;border-radius:3px;text-decoration:none}#report-result{font-size:12px;overflow-wrap:anywhere}
      .muted{font-size:12px;color:#575a60}.hidden{display:none!important}
      #status{font-size:12px;white-space:pre-wrap;overflow-wrap:anywhere;padding:10px;border:1px solid #b5b0a7;background:#f4f1e9;max-height:190px;overflow:auto}
      .branch-row{padding:10px 0;border-top:1px solid #bdb8b0}.branch-name{font-weight:650;overflow-wrap:anywhere}
      .branch-fields{display:grid;grid-template-columns:minmax(70px,1fr) 80px minmax(80px,1fr);gap:8px}
      .branch-fields label{font-size:11px;font-weight:500}.branch-fields input,.branch-fields select{margin-top:3px;font-size:13px;padding:7px 5px}
      #sensitivity-result{font-size:12px;overflow-wrap:anywhere}#sensitivity-result table{width:100%;border-collapse:collapse;margin:10px 0}#sensitivity-result th,#sensitivity-result td{padding:7px 4px;border-bottom:1px solid #aaa69f;text-align:left;vertical-align:top}.sensitivity-headline{font-weight:700;font-size:14px}
      #probability-alert{white-space:pre-wrap;max-height:240px;overflow:auto;padding:12px;margin:12px 0;background:#fce0df;color:#841820;border:2px solid #a1272d;border-radius:3px;font-size:12px;font-weight:550}
      #branch-total{font-size:12px;line-height:1.5}details{margin:12px 0}summary{cursor:pointer;font-weight:600;padding:5px 0}
      .actions{display:flex;gap:8px;margin:14px 0}.actions #save-node{flex:2}.actions #add{flex:1}
      @media(max-width:330px){main{padding:0 12px 16px}.titlebar{margin-left:-12px;margin-right:-12px}.branch-fields{grid-template-columns:1fr 80px}.branch-fields label:last-child{grid-column:1/-1}}
    </style><main>
      <header class="brand"><img src="${logo}" alt="Blue chance circle, green decision square, yellow outcome triangle"><div><h1>Macademia</h1><div class="product-name">Decision Trees</div></div><span class="version">1.0 · Pilot</span></header>
      <div class="row"><button id="load">Load active tree</button><button id="new">New tree</button></div>
      <div class="row"><button id="save-workbook" class="primary">Save tree &amp; workbook</button><button id="resume-last" class="hidden">Resume last tree</button></div>
      <div id="status" role="status" aria-live="polite">Create a tree or load your existing tree sheet.</div>
      <div id="probability-alert" class="hidden" role="status" aria-live="polite"></div>
      <details class="section"><summary>Quick start / 1.0</summary><p>1. Create a sample or load a tree. Select a card to edit it.</p><p>2. Enter payoffs and probabilities, then Apply changes. Fixed entries stay entered; Auto branches divide the remainder.</p><p>3. Check inputs, test decision break-even points, and export the full tree below. Save the workbook with Cmd+S.</p><p class="muted">Macademia — Decision Trees runs from its hosted editor. No Node.js or Terminal is needed. Workbook inputs and calculations stay in Excel; PDF generation runs locally in this pane. Internet is required to open the editor.</p></details>
      <div class="segmented" role="group" aria-label="Worksheet view"><button id="view-edit" aria-pressed="true">Edit</button><button id="view-presentation" aria-pressed="false">Presentation</button></div>
      <select id="view" class="hidden" aria-label="Worksheet view"><option value="edit">Edit</option><option value="presentation">Presentation</option></select>
      <label for="node">Selected node</label><div class="node-picker"><select id="node"></select><button id="use-cell">Use cell</button></div>
      <div class="titlebar" id="node-title">Select a node</div><p id="selected" class="muted"></p>
      <section class="section"><h2>Node</h2><label for="label">Node label</label><input id="label" maxlength="80">
      <div class="field-pair"><div><label for="kind">Node type</label><select id="kind"><option value="decision">Decision · square</option><option value="chance">Chance · circle</option><option value="outcome">Outcome · triangle</option></select></div><div><label for="payoff">Payoff at node</label><input id="payoff" value="0"></div></div>
      <p class="muted">Costs are negative. Blank means $0. This payoff is added once to downstream value.</p></section>
      <section class="section" id="incoming" style="display:none"><h2>Incoming probability</h2><label for="probability">Probability</label><input id="probability" value="50%" disabled>
      <label class="check"><input type="checkbox" id="remainder" disabled>Auto: share the remaining percentage</label><p id="probability-note" class="muted"></p></section>
      <section class="section" id="branches" style="display:none"><h2>Branches / outcomes</h2><div id="branch-rows"></div><p id="branch-total"></p><p class="muted">Typing a probability makes it Fixed. Auto branches share the rest. Apply changes to update the sheet.</p></section>
      <div class="actions"><button id="save-node" class="primary">Apply changes</button><button id="add">Add branch</button></div>
      <div class="row"><button id="collapse">Collapse branch</button><button id="expand-all">Expand all</button></div><p id="collapse-note" class="muted"></p>
      <details class="section" id="report-panel"><summary>Export / PDF</summary>
        <label for="report-title">Report title</label><input id="report-title" maxlength="80" placeholder="Decision analysis">
        <label for="report-format">PDF layout</label><select id="report-format"><option value="overview">Single-page overview · full-size poster</option><option value="pages">Readable pages · Letter landscape</option></select>
        <p class="muted">Exports every branch, including collapsed descendants. The overview is one large page for zooming; readable pages keep cards intact. PDFs contain high-resolution images, so text is not selectable.</p>
        <div class="row"><button id="report-generate" class="primary">Generate PDF</button><button id="check-tree">Check inputs</button></div>
        <div id="report-result" role="status" aria-live="polite"></div>
        <details><summary>If Excel blocks the download</summary><p class="muted">Use native Mac printing instead. This expands all branches and switches the worksheet to Presentation. Undo last edit restores your working view.</p><button id="report-print">Prepare Excel print</button><p class="muted">Then choose File → Print → PDF → Save as PDF. This fallback fits the entire tree on one page.</p></details>
      </details>
      <details class="section" id="sensitivity-panel"><summary>What would change this decision?</summary>
        <label for="sensitivity-decision">Decision to test</label><select id="sensitivity-decision"></select>
        <label for="sensitivity-field">Input to vary</label><select id="sensitivity-field"><option value="payoff">Payoff at a node</option><option value="probability">Branch probability</option></select>
        <label for="sensitivity-input">Node / incoming branch</label><select id="sensitivity-input"></select>
        <div class="field-pair"><div><label for="sensitivity-low">Test from</label><input id="sensitivity-low" value="-100000"></div><div><label for="sensitivity-high">Test to</label><input id="sensitivity-high" value="100000"></div></div>
        <p class="muted">Probability ranges are capped by the other Fixed entries. At least one other branch must be Auto. Payoffs accept numbers such as -4000 or $80,000.</p>
        <div class="row"><button id="sensitivity-run" class="primary">Find break-even points</button></div>
        <p class="muted">Read-only what-if: source inputs stay unchanged. Other formula values are held at their current evaluated numbers.</p>
        <div id="sensitivity-result" role="status" aria-live="polite"></div>
      </details>
      <details><summary>Branch tools</summary>
        <div class="row"><button id="move-up">Move up</button><button id="move-down">Move down</button></div>
        <button id="duplicate">Duplicate branch + subtree</button>
        <p class="muted">A copy beside a chance branch starts as Auto; existing Fixed probabilities stay entered. Formula inputs keep their references.</p>
        <label for="insert-kind">Insert before selected node</label><select id="insert-kind"><option value="decision">Decision</option><option value="chance">Chance</option></select>
        <div class="row"><button id="insert-node">Insert node</button></div>
        <p class="muted">Inserts a zero-payoff node with one child. The original incoming probability transfers to it, preserving the current value. Rename it, then add alternatives.</p>
      </details>
      <button id="discard-edits">Discard sidebar edits</button>
      <details id="probability-group" style="display:none"><summary>Bulk probabilities</summary><p id="probability-order" class="muted" style="white-space:pre-wrap"></p>
      <label for="probability-list">Percentages in branch order</label><input id="probability-list" placeholder="70 / * / *">
      <p class="muted">Each * or X shares the remainder equally. These actions apply immediately; apply any pending form changes first.</p>
      <div class="row"><button id="apply-probabilities">Apply list</button><button id="reset-probabilities">Reset group to Auto</button></div></details>
      <details class="section"><summary>Tree tools</summary><label class="check"><input type="checkbox" id="show-strategy" checked>Show BEST / TIED BEST decision branches</label><p class="muted">Markers identify the best choice at each decision, including ties. All chance outcomes stay possible.</p><label for="objective">Decision objective</label><select id="objective"><option value="max">Maximize net payoff</option><option value="min">Minimize value / cost</option></select>
      <div class="row"><button id="redraw">Refresh / redraw</button><button id="undo">Undo last edit</button></div>
      <div class="row"><button id="sample">New sample</button><button id="delete">Delete branch + descendants</button></div>
      <p class="muted">Card selection keeps unapplied sidebar edits until you Apply or Discard them. Undo restores the previous saved edit.</p></details>
      <p class="muted">Apply sidebar changes, then Save tree &amp; workbook (or Cmd+S). Reopen the .xlsx and open Macademia — Decision Trees to restore your editing session automatically. CSV cannot preserve an editable tree.</p>
      <p class="muted">Keep formula assumptions and notes on separate sheets. The dedicated canvas is rebuilt on edits; do not move its cells or shapes manually. Presentation is a visual mode, not sheet protection.</p>
    </main>`;
  el("save-workbook").onclick = () => { void saveTreeWorkbook(); };
  el("resume-last").onclick = () => { void resumeLastTree(); };
  input("report-title").oninput = () => { invalidateReport("Report title changed. Generate the PDF again."); };
  el("report-format").onchange = () => { invalidateReport("PDF layout changed. Generate the PDF again."); };
  el("report-generate").onclick = () => { void generateReport(); };
  el("check-tree").onclick = () => { void checkTree(); };
  el("report-print").onclick = () => { void prepareExcelPrint(); };
  el("sensitivity-run").onclick = () => { void runSensitivity(); };
  el("sensitivity-field").onchange = () => { if (current) { sensitivityDefaults = ""; refreshSensitivity(current); } };
  el("sensitivity-input").onchange = () => { sensitivityRangeDefaults(true); };
  el("sensitivity-decision").onchange = () => { el("sensitivity-result").replaceChildren(); };
  el("discard-edits").onclick = () => {
    if (!current) return;
    el                   ("node").value = editingId; fillForm();
    el("status").textContent = "Sidebar edits discarded. Worksheet inputs were not changed.";
  };
  el("show-strategy").onchange = () => {
    if (!requireApplied()) { input("show-strategy").checked = current?.showStrategy !== false; return; }
    const show = input("show-strategy").checked;
    void modify(m => { m.showStrategy = show; });
  };
  el("collapse").onclick = () => { void modify((m, n) => {
    if (!children(m, n.id).length) throw new Error("This node has no descendants to collapse.");
    n.collapsed = !n.collapsed;
  }); };
  el("expand-all").onclick = () => { void modify(m => { m.nodes.forEach(n => { n.collapsed = false; }); }); };
  el("move-up").onclick = () => { void modify((m, n) => { reorderBranch(m, n, -1); }); };
  el("move-down").onclick = () => { void modify((m, n) => { reorderBranch(m, n, 1); }); };
  el("duplicate").onclick = () => { void modify((m, n) => duplicateBranch(m, n)); };
  el("insert-node").onclick = () => {
    const kind = el                   ("insert-kind").value                         ;
    void modify((m, n) => insertBefore(m, n, kind));
  };
  el("new").onclick = () => { void create(false); }; el("sample").onclick = () => { void create(true); };
  el("load").onclick = () => { void load(); }; el("use-cell").onclick = () => { void load(true); };
  el("node").onchange = () => { void load(); };
  el("remainder").onchange = () => {
    if (current) input("probability").value = probabilityText(displayedProbability(current, chosen(current)));
  };
  input("probability").oninput = () => { input("remainder").checked = false; };
  for (const mode of ["edit", "presentation"]         ) el("view-" + mode).onclick = () => {
    void modify(m => { m.view = mode; });
  };
  el("view").onchange = () => {
    const view = el                   ("view").value                           ;
    void modify(m => { m.view = view; });
  };
  el("objective").onchange = () => { const objective = el                   ("objective").value                 ; void modify(m => { m.objective = objective; }); };
  el("save-node").onclick = () => {
    const drafts = clone(branchDrafts);
    const edits = { label: input("label").value.trim(), kind: el                   ("kind").value        , payoff: input("payoff").value, probability: input("probability").value, remainder: input("remainder").checked };
    void modify((m, n) => {
      if (!edits.label) throw new Error("Give the node a label.");
      const list = children(m, n.id);
      if (edits.kind === "outcome" && list.length) throw new Error("Delete this node's child branches before converting it to an outcome.");
      const oldKind = n.kind; n.label = edits.label.slice(0, 80); n.kind = edits.kind; n.payoff = payoffExpr(edits.payoff);
      const parent = m.nodes.find(c => c.id === n.parent);
      if (parent?.kind === "chance") {
        if (edits.remainder) setRemainder(m, parent, n.id);
        else {
          n.probability = probabilityExpr(edits.probability);
          if (typeof n.probability === "number" && (n.probability < 0 || n.probability > 1)) throw new Error("Probability must be between 0% and 100%.");
          n.probabilityFixed = true; n.remainder = false;
        }
      }
      if (n.kind !== "outcome" && !list.length) { addOutcome(m, n, "Branch 1", 0.5); addOutcome(m, n, "Branch 2", 0.5); }
      else if (n.kind === "chance" && oldKind !== "chance") list.forEach(c => { c.probability = 1 / list.length; c.probabilityFixed = false; });
      applyBranchDrafts(m, drafts);
    }, true);
  };
  el("apply-probabilities").onclick = () => {
    const list = input("probability-list").value;
    void modify((m, n) => {
      const parent = chanceGroup(m, n);
      if (!parent) throw new Error("Select a chance node or a branch leaving a chance node.");
      const branches = children(m, parent.id);
      if (!branches.length) throw new Error("Add branches to this chance node first.");
      const parsed = parseProbabilityList(list, branches.length);
      branches.forEach((c, i) => { c.probability = parsed.probabilities[i]; c.probabilityFixed = !parsed.automatic.includes(i); });
      ensureRemainders(m);
    });
  };
  el("reset-probabilities").onclick = () => { void modify((m, n) => {
    const parent = chanceGroup(m, n);
    if (!parent) throw new Error("Select a chance node or one of its branches first.");
    children(m, parent.id).forEach(c => { c.probabilityFixed = false; });
    ensureRemainders(m);
  }); };
  el("add").onclick = () => { void modify((m, n) => {
    if (n.kind === "outcome") throw new Error("Change this outcome to Decision or Chance and click Apply changes first.");
    n.collapsed = false;
    const added = addOutcome(m, n, `Branch ${children(m, n.id).length + 1}`, n.kind === "chance" ? 0 : 1);
    if (n.kind === "chance") setRemainder(m, n, added.id);
    return added.id;
  }); };
  el("delete").onclick = () => { void modify((m, n) => {
    if (n.parent === null) throw new Error("The root cannot be deleted. Start a New tree instead.");
    const removed = new Set        ();
    function remove(id        ) { removed.add(id); children(m, id).forEach(c => remove(c.id)); }
    remove(n.id); m.nodes = m.nodes.filter(c => !removed.has(c.id)); return n.parent;
  }); };
  el("redraw").onclick = () => { void modify(() => {}); };
  el("undo").onclick = () => { if (!requireApplied()) return; void perform(async ctx => {
    const { sheet, model } = await active(ctx);
    const undo = ctx.workbook.settings.getItemOrNullObject(KEY + sheet.id + ".undo"); undo.load("value"); await ctx.sync();
    if (undo.isNullObject) throw new Error("There is no edit to undo for this tree.");
    const previous = JSON.parse(String(undo.value))         ; await draw(ctx, sheet, previous, model);
    return "Previous edit restored. Undo again toggles back. Save with Cmd+S.";
  }); };
  void restoreOnOpen();
});
