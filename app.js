/* ===========================
   Fiche Projet CLB — app.js
   - Scores Q1→Q5 selon règles fournies
   - Alertes dynamiques
   - Historique financements (table)
   - Sauvegarde/chargement localStorage
   - Génération PDF via jsPDF + autotable
   =========================== */

const LS_KEY = "ficheProjet_v1";

// Seuils labels (Q3 / Q5) modifiables ici
const Q3_LABEL_THRESHOLDS = [
  { max: 2, label: "faible" },
  { max: 5, label: "limitée" },
  { max: 8, label: "modérée" },
  { max: Infinity, label: "forte" },
];

const Q5_LABEL_THRESHOLDS = [
  { max: 1, label: "faible" },
  { max: 3, label: "modérée" },
  { max: 5, label: "forte" },
  { max: Infinity, label: "très forte" },
];

// Mapping pays/zone Q1 : points + vigilance (on prend le MAX parmi les zones cochées)
const ZONE_SCORE = {
  z_fr: { pts: 2, vig: "faible" },           // France/Europe/Canada
  z_ueuk: { pts: 2, vig: "faible" },
  z_canada: { pts: 2, vig: "faible" },
  z_usau: { pts: 3, vig: "modérée" },        // USA/Australie/Japon…
  z_asie: { pts: 3, vig: "modérée" },        // Japon/Corée/Taiwan/Singapour
  z_chineinde: { pts: 4, vig: "élevée" },    // Chine/Inde/Autres
  z_autres: { pts: 4, vig: "élevée" },
};

const fundingSources = ["Institutionnel", "Industriel", "Fondation", "Europe", "Autre"];

const $ = (sel) => document.querySelector(sel);

const form = $("#ficheForm");
const alertsEl = $("#alerts");
const lastSavedEl = $("#lastSaved");

const q1ScoreEl = $("#q1Score");
const q1VigEl = $("#q1Vigilance");
const q2ScoreEl = $("#q2Score");
const q2LabelEl = $("#q2Label");
const q3ScoreEl = $("#q3Score");
const q3LabelEl = $("#q3Label");
const q4ScoreEl = $("#q4Score");
const q5ScoreEl = $("#q5Score");
const q5LabelEl = $("#q5Label");

const autresWrap = $("#autresWrap");
const zAutres = $("#z_autres");

const pfAutreWrap = $("#pfAutreWrap");
const pfAutre = $("#pf_autre");

const cpWrap = $("#cpWrap");
const finCp = $("#fin_cp");

const aapWrap = $("#aapWrap");
const aap = $("#aap");
const q6Section = $("#q6Section");

const fundingTableBody = $("#fundingTable tbody");

$("#btnAddFunding").addEventListener("click", addFundingRow);
$("#btnSave").addEventListener("click", saveLocal);
$("#btnLoad").addEventListener("click", loadLocal);
$("#btnReset").addEventListener("click", resetAll);
$("#btnPdf").addEventListener("click", generatePdf);

form.addEventListener("input", () => {
  handleConditionalFields();
  recomputeAll();
});

form.addEventListener("change", () => {
  handleConditionalFields();
  recomputeAll();
});

function handleConditionalFields() {
  autresWrap.style.display = zAutres?.checked ? "block" : "none";
  pfAutreWrap.style.display = pfAutre?.checked ? "block" : "none";
  cpWrap.style.display = finCp?.checked ? "block" : "none";
  aapWrap.style.display = aap?.checked ? "grid" : "none";

  const showQ6 = !!(form.elements.act_data?.checked || form.elements.act_ech?.checked);
  if (q6Section) q6Section.style.display = showQ6 ? "block" : "none";

  if (!showQ6) {
    ["trf_ech", "trf_data", "trf_multi"].forEach((name) => {
      if (form.elements[name]) form.elements[name].checked = false;
    });
    ["trf_zone", "trf_dest", "trf_desc", "trf_com"].forEach((name) => {
      if (form.elements[name]) form.elements[name].value = "";
    });
  }
}

/* ---------------------------
   Helpers: read form values
---------------------------- */
function getFormData() {
  const fd = new FormData(form);
  const obj = {};
  for (const [k, v] of fd.entries()) obj[k] = v;

  // checkboxes
  const checkNames = [
    "pj",
    "p_acad","p_indus","p_multi",
    "z_fr","z_ueuk","z_canada","z_usau","z_asie","z_chineinde","z_autres",
    "pub","educ","clin","auteur",
    "act_essai","act_prom","act_data","act_ech","act_drci","act_clin","act_cher","act_st","act_achat","act_eq","act_heberg","act_pf",
    "pf_pic","pf_par","pf_cyto","pf_geno","pf_bioinfo","pf_preclin","pf_autre",
    "fin_exist","fin_ext","fin_valide","fin_couvre","fin_cp",
    "aap",
    "bv_sig","bv_mol","bv_aut","licence",
    "trf_ech","trf_data","trf_multi"
  ];
  checkNames.forEach((n) => (obj[n] = !!form.elements[n]?.checked));

  // numbers
  obj.fin_montant = toNumber(obj.fin_montant);
  obj.fin_frais = toNumber(obj.fin_frais);

  // funding table rows
  obj.fundingHistory = readFundingTable();

  return obj;
}

function toNumber(v) {
  if (v === undefined || v === null || v === "") return 0;
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function getCheckedCountByPrefix(prefix) {
  const inputs = form.querySelectorAll(`input[type="checkbox"][name^="${prefix}"]`);
  let c = 0;
  inputs.forEach(i => { if (i.checked) c++; });
  return c;
}

function labelFromThresholds(value, thresholds) {
  for (const t of thresholds) {
    if (value <= t.max) return t.label;
  }
  return thresholds[thresholds.length - 1].label;
}

/* ---------------------------
   Q1: score exact
---------------------------- */
function computeQ1(d) {
  // 1) Type partenaire (max 5)
  let type = 0;
  if (d.p_acad) type += 2;
  if (d.p_indus) type += 2;
  if (d.p_multi) type += 1;
  type = Math.min(type, 5);

  // 2) Nombre partenaires extérieurs (max 3)
  let nb = 0;
  if (d.nbPart === "1") nb = 1;
  else if (d.nbPart === "2-3") nb = 2;
  else if (d.nbPart === "4+") nb = 3;

  // 3) Pays (max 4) = MAX des zones cochées
  const zoneKeys = Object.keys(ZONE_SCORE);
  let maxPts = 0;
  let maxVig = "—";
  zoneKeys.forEach((k) => {
    if (d[k]) {
      const z = ZONE_SCORE[k];
      if (z.pts > maxPts) {
        maxPts = z.pts;
        maxVig = z.vig;
      }
    }
  });

  const total = type + nb + maxPts;
  return { total, type, nb, pays: maxPts, vigilance: maxPts ? maxVig : "—" };
}

/* ---------------------------
   Q2: score exact /5
---------------------------- */
function computeQ2(d) {
  let s = 0;
  if (d.pub) s += 1;
  if (d.educ) s += 1;
  if (d.clin) s += 1;
  if (d.auteur) s += 1;
  if ((d.journal || "Aucun") !== "Aucun") s += 1;

  let label = "—";
  if (s <= 1) label = "Impact faible";
  else if (s <= 3) label = "Impact modéré";
  else if (s <= 4) label = "Impact élevé";
  else label = "Impact très élevé";

  return { score: s, label };
}

/* ---------------------------
   Q3: activités + plateformes (0.5)
---------------------------- */
function computeQ3(d) {
  const actCount = getCheckedCountByPrefix("act_");
  const pfCount = getCheckedCountByPrefix("pf_");
  // pf_autre_txt doesn't count; only checkbox counts
  const score = actCount * 1 + pfCount * 0.5;

  const label = labelFromThresholds(score, Q3_LABEL_THRESHOLDS);
  return { score, label, actCount, pfCount };
}

/* ---------------------------
   Q4: score exact /6
---------------------------- */
function computeQ4(d) {
  let s = 0;
  // 4 cases clés
  if (d.fin_ext) s += 1;
  if (d.fin_valide) s += 1;
  if (d.fin_couvre) s += 1;
  if (d.fin_cp) s += 1;

  // montant
  if (d.fin_montant > 0 && d.fin_montant < 30000) s += 1;
  else if (d.fin_montant >= 30000) s += 2;

  return { score: s };
}

/* ---------------------------
   Q5: score exact
---------------------------- */
function computeQ5(d) {
  // 1) Brevet = MAX
  let brevet = 0;
  if (d.bv_sig) brevet = Math.max(brevet, 1);
  if (d.bv_aut) brevet = Math.max(brevet, 2);
  if (d.bv_mol) brevet = Math.max(brevet, 3);

  // 2) Contrat licence
  let contrat = d.licence ? brevet : 0;
  const score = brevet + contrat;
  const label = labelFromThresholds(score, Q5_LABEL_THRESHOLDS);

  return { score, label, brevet, contrat };
}

/* ---------------------------
   Alertes
---------------------------- */
function computeAlerts(d, q1, q2, q3, q4, q5) {
  const alerts = [];

  // A) gros enjeu scientifique
  if (q2.label === "Impact élevé" || q2.label === "Impact très élevé") {
    alerts.push({ level: "warn", text: "Gros enjeu scientifique : anticiper stratégie de publication, rôles auteurs, besoins biostat/méthodo." });
  }

  // B) chef de projet recommandé
  const projectComplex =
    (q3.label === "forte") ||
    d.p_multi ||
    d.p_indus ||
    d.trf_multi;
  if (!d.fin_cp && projectComplex) {
    alerts.push({ level: "warn", text: "Projet complexe : un chef de projet/produit est recommandé (coordination, jalons, budget)." });
  }

  // C) valorisation
  const anyBrevet = d.bv_sig || d.bv_mol || d.bv_aut;
  if (q5.label === "forte" || q5.label === "très forte" || anyBrevet) {
    alerts.push({ level: "warn", text: "Potentiel de valorisation : se rapprocher de la cellule valorisation (brevet/licence/contrats)." });
  }

  // D) MR004 si transfert de données
  if (d.trf_data) {
    alerts.push({ level: "danger", text: "Transfert/échange de données : compléter la fiche MR004 (déclaration DPD) avant échange." });
  }

  // E) transfert échantillons
  if (d.trf_ech) {
    alerts.push({ level: "warn", text: "Transfert d’échantillons : vérifier circuit CRB/CMT. Si données associées, compléter aussi MR004." });
  }

  // F) hors UE non adéquat (données)
  const riskyZone = (d.trf_zone === "Chine/Inde" || d.trf_zone === "Autres");
  if (d.trf_data && riskyZone) {
    alerts.push({ level: "danger", text: "Données hors UE (risque élevé) : validation DPD + clauses contractuelles renforcées." });
  }

  // G) alerte contrat
  if (d.p_indus || d.act_st) {
    alerts.push({ level: "warn", text: "Prévoir contractualisation (collaboration/sous-traitance/consortium)." });
  }

  return alerts;
}

function renderAlerts(alerts) {
  alertsEl.innerHTML = "";
  if (!alerts.length) {
    const div = document.createElement("div");
    div.className = "alert ok";
    div.textContent = "Aucune alerte particulière détectée à ce stade.";
    alertsEl.appendChild(div);
    return;
  }
  alerts.forEach(a => {
    const div = document.createElement("div");
    div.className = `alert ${a.level}`;
    div.textContent = a.text;
    alertsEl.appendChild(div);
  });
}

/* ---------------------------
   Recompute all & update UI
---------------------------- */
function recomputeAll() {
  const d = getFormData();

  const q1 = computeQ1(d);
  const q2 = computeQ2(d);
  const q3 = computeQ3(d);
  const q4 = computeQ4(d);
  const q5 = computeQ5(d);

  q1ScoreEl.textContent = String(q1.total);
  q1VigEl.textContent = q1.vigilance;

  q2ScoreEl.textContent = String(q2.score);
  q2LabelEl.textContent = q2.label;

  q3ScoreEl.textContent = String(q3.score);
  q3LabelEl.textContent = q3.label;

  q4ScoreEl.textContent = String(q4.score);

  q5ScoreEl.textContent = String(q5.score);
  q5LabelEl.textContent = q5.label;

  const alerts = computeAlerts(d, q1, q2, q3, q4, q5);
  renderAlerts(alerts);
}

/* ---------------------------
   Funding table
---------------------------- */
function addFundingRow(prefill = {}) {
  const tr = document.createElement("tr");

  tr.innerHTML = `
    <td><input type="month" class="fund-date" value="${escapeHtml(prefill.date || "")}"></td>
    <td>
      <select class="fund-source">
        <option value="">—</option>
        ${fundingSources.map(s => `<option ${prefill.source === s ? "selected" : ""}>${s}</option>`).join("")}
      </select>
    </td>
    <td><input type="number" class="fund-amount" min="0" step="100" value="${escapeHtml(prefill.amount ?? "")}"></td>
    <td><input type="text" class="fund-note" value="${escapeHtml(prefill.note || "")}"></td>
    <td><button type="button" class="secondary fund-del">Supprimer</button></td>
  `;

  tr.querySelector(".fund-del").addEventListener("click", () => {
    tr.remove();
    recomputeAll();
  });

  // Recompute when editing
  tr.querySelectorAll("input,select").forEach(el => el.addEventListener("input", () => recomputeAll()));

  fundingTableBody.appendChild(tr);
}

function readFundingTable() {
  const rows = [];
  fundingTableBody.querySelectorAll("tr").forEach(tr => {
    const date = tr.querySelector(".fund-date")?.value || "";
    const source = tr.querySelector(".fund-source")?.value || "";
    const amount = toNumber(tr.querySelector(".fund-amount")?.value || 0);
    const note = tr.querySelector(".fund-note")?.value || "";
    if (date || source || amount || note) {
      rows.push({ date, source, amount, note });
    }
  });
  return rows;
}

function fillFundingTable(rows) {
  fundingTableBody.innerHTML = "";
  (rows || []).forEach(r => addFundingRow(r));
}

/* ---------------------------
   Save / load / reset
---------------------------- */
function saveLocal() {
  const d = getFormData();
  const payload = {
    savedAt: new Date().toISOString(),
    data: d
  };
  localStorage.setItem(LS_KEY, JSON.stringify(payload));
  showLastSaved(payload.savedAt);
  alert("Sauvegarde locale effectuée ✅");
}

function loadLocal() {
  const raw = localStorage.getItem(LS_KEY);
  if (!raw) {
    alert("Aucune sauvegarde trouvée.");
    return;
  }
  const payload = JSON.parse(raw);
  applyDataToForm(payload.data || {});
  showLastSaved(payload.savedAt);
  handleConditionalFields();
  recomputeAll();
  alert("Sauvegarde chargée ✅");
}

function resetAll() {
  if (!confirm("Réinitialiser la fiche ?")) return;
  form.reset();
  fundingTableBody.innerHTML = "";
  localStorage.removeItem(LS_KEY);
  lastSavedEl.textContent = "";
  handleConditionalFields();
  recomputeAll();
}

function applyDataToForm(d) {
  // simple fields (text/select/number/date)
  for (const [k, v] of Object.entries(d)) {
    if (k === "fundingHistory") continue;
    const el = form.elements[k];
    if (!el) continue;

    if (el.type === "checkbox") {
      el.checked = !!v;
    } else {
      el.value = v ?? "";
    }
  }

  // funding rows
  fillFundingTable(d.fundingHistory || []);
}

/* ---------------------------
   PDF generation
---------------------------- */
async function generatePdf() {
  // basic required check
  const d = getFormData();
  const missing = [];
  if (!String(d.titre || "").trim()) missing.push("Titre du projet");
  if (!String(d.porteur || "").trim()) missing.push("Porteur");
  if (!String(d.email || "").trim()) missing.push("Email");
  if (!String(d.resume || "").trim()) missing.push("Résumé");

  if (missing.length) {
    alert("Champs requis manquants :\n- " + missing.join("\n- "));
    return;
  }

  const q1 = computeQ1(d);
  const q2 = computeQ2(d);
  const q3 = computeQ3(d);
  const q4 = computeQ4(d);
  const q5 = computeQ5(d);
  const alerts = computeAlerts(d, q1, q2, q3, q4, q5);

  // jsPDF UMD access
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });

  // Header
  const now = new Date();
  const dateStr = now.toLocaleString("fr-FR");

  let y = 12;

  // Try embed logo if present & same-origin (optional)
  await tryAddLogoToPdf(doc);

  doc.setFontSize(16);
  doc.text("Fiche Projet — Centre Léon Bérard", 14, y);
  y += 6;

  doc.setFontSize(10);
  doc.text(`Générée le : ${dateStr}`, 14, y);
  y += 6;

  // Summary box (scores)
  doc.setFontSize(11);
  doc.text(`Scores :  Q1 ${q1.total}/12  |  Q2 ${q2.score}/5 (${q2.label})  |  Q3 ${q3.score} (${q3.label})  |  Q4 ${q4.score}/6  |  Q5 ${q5.score} (${q5.label})`, 14, y);
  y += 6;

  // Alerts
  const alertLines = alerts.length ? alerts.map(a => `- ${a.text}`) : ["- Aucune alerte particulière."];
  doc.setFontSize(11);
  doc.text("Alertes générées :", 14, y);
  y += 5;
  doc.setFontSize(10);
  y = addWrappedLines(doc, alertLines.join("\n"), 14, y, 180);
  y += 3;

  // Q0 table
  doc.autoTable({
    startY: y,
    head: [["Q0 — Informations générales", ""]],
    body: [
      ["Titre du projet", safe(d.titre)],
      ["Porteur", safe(d.porteur)],
      ["Service / unité", safe(d.unite)],
      ["Email", safe(d.email)],
      ["Date début", safe(d.date_debut)],
      ["Date fin", safe(d.date_fin)],
      ["Pièce jointe", d.pj ? "Oui" : "Non"],
      ["Résumé / descriptif", safe(d.resume)],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245] },
    columnStyles: { 0: { cellWidth: 45 }, 1: { cellWidth: 140 } },
    didParseCell: (data) => {
      data.cell.styles.valign = "top";
    }
  });

  // Q1 table
  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 6,
    head: [["Q1 — Partenaires / International", ""]],
    body: [
      ["Type partenaire", [
        d.p_acad ? "Académique" : null,
        d.p_indus ? "Industriel" : null,
        d.p_multi ? "Multicentrique" : null
      ].filter(Boolean).join(", ") || "—"],
      ["Nb partenaires extérieurs", safe(d.nbPart)],
      ["Coordonnateur", safe(d.coord)],
      ["Zones cochées", zonesSelectedText(d)],
      ["Score Q1", `${q1.total}/12 (Type ${q1.type} + Nb ${q1.nb} + Pays ${q1.pays})`],
      ["Vigilance", q1.vigilance],
      ["Autres (préciser)", d.z_autres ? safe(d.z_autres_txt) : "—"],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245] },
    columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  // Q2 table
  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 6,
    head: [["Q2 — Impact médico-scientifique", ""]],
    body: [
      ["Publication", d.pub ? "Oui" : "Non"],
      ["Impact éducatif", d.educ ? "Oui" : "Non"],
      ["Application clinique CT", d.clin ? "Oui" : "Non"],
      ["1er/dernier auteur", d.auteur ? "Oui" : "Non"],
      ["Journal", safe(d.journal)],
      ["Score Q2", `${q2.score}/5 — ${q2.label}`],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245] },
    columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  // Q3 table
  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 6,
    head: [["Q3 — Ressources / activités mobilisées", ""]],
    body: [
      ["Activités cochées", activitiesSelectedText()],
      ["Plateformes cochées", platformsSelectedText(d)],
      ["Score Q3", `${q3.score} — ${q3.label} (activités=${q3.actCount}, plateformes=${q3.pfCount})`],
      ["Autre plateforme (préciser)", d.pf_autre ? safe(d.pf_autre_txt) : "—"],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245] },
    columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  // Q4 table
  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 6,
    head: [["Q4 — Financement", ""]],
    body: [
      ["Financement existant", d.fin_exist ? "Oui" : "Non"],
      ["Financement extérieur", d.fin_ext ? "Oui" : "Non"],
      ["Financement validé", d.fin_valide ? "Oui" : "Non"],
      ["Couvre ressources CLB", d.fin_couvre ? "Oui" : "Non"],
      ["Chef de projet identifié", d.fin_cp ? "Oui" : "Non"],
      ["Nom chef de projet", d.fin_cp ? safe(d.fin_cp_nom) : "—"],
      ["Montant total (€)", formatEuros(d.fin_montant)],
      ["Frais de gestion (€)", formatEuros(d.fin_frais)],
      ["Score Q4", `${q4.score}/6`],
      ["Appel à projet associé", d.aap ? "Oui" : "Non"],
      ["AAP — Nom", d.aap ? safe(d.aap_nom) : "—"],
      ["AAP — Organisme", d.aap ? safe(d.aap_org) : "—"],
      ["AAP — Date soumission", d.aap ? safe(d.aap_date) : "—"],
      ["AAP — Statut", d.aap ? safe(d.aap_statut) : "—"],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245] },
    columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  // Funding history table
  const fundingRows = (d.fundingHistory || []).map(r => [
    r.date || "—",
    r.source || "—",
    formatEuros(r.amount || 0),
    r.note || ""
  ]);

  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 4,
    head: [["Historique des financements (optionnel)"]],
    body: fundingRows.length ? fundingRows : [["—", "—", "—", ""]],
    theme: "grid",
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245] }
  });

  // Q5 table
  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 6,
    head: [["Q5 — Impact valorisation", ""]],
    body: [
      ["Brevet (MAX)", `Signature/cible=${d.bv_sig ? "Oui" : "Non"} | Molécule=${d.bv_mol ? "Oui" : "Non"} | Autre=${d.bv_aut ? "Oui" : "Non"} (valeur retenue=${q5.brevet})`],
      ["Contrat licence", d.licence ? `Oui (+${q5.brevet})` : "Non"],
      ["Score Q5", `${q5.score} — ${q5.label}`],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245] },
    columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  // Q6 table
  if (d.act_data || d.act_ech) {
    doc.autoTable({
      startY: doc.lastAutoTable.finalY + 6,
      head: [["Q6 — Données / échantillons (RGPD / MR004)", ""]],
      body: [
        ["Transfert échantillons", d.trf_ech ? "Oui" : "Non"],
        ["Transfert données", d.trf_data ? "Oui" : "Non"],
        ["Multicentrique", d.trf_multi ? "Oui" : "Non"],
        ["Zone destinataire", safe(d.trf_zone)],
        ["Destinataire", safe(d.trf_dest)],
        ["Description", safe(d.trf_desc)],
        ["Commentaires", safe(d.trf_com)],
      ],
      styles: { fontSize: 9, cellPadding: 2 },
      headStyles: { fillColor: [245,245,245] },
      columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
      didParseCell: (data) => { data.cell.styles.valign = "top"; }
    });
  }

  // Signature lines at the end
  const finalY = doc.lastAutoTable.finalY + 12;
  doc.setFontSize(10);
  doc.text("Nom / Signature :", 14, finalY);
  doc.line(45, finalY, 120, finalY);
  doc.text("Date :", 130, finalY);
  doc.line(142, finalY, 195, finalY);

  const filename = `FicheProjet_CLB_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`;
  doc.save(filename);
}

/* ---------------------------
   PDF helpers
---------------------------- */
function safe(v) {
  return (v === undefined || v === null || v === "") ? "—" : String(v);
}

function formatEuros(n) {
  const val = toNumber(n);
  if (!val) return "—";
  return val.toLocaleString("fr-FR") + " €";
}

function slugify(s) {
  return String(s).toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 60);
}

function addWrappedLines(doc, text, x, y, maxWidth) {
  const lines = doc.splitTextToSize(text, maxWidth);
  doc.text(lines, x, y);
  return y + lines.length * 4.2;
}

async function tryAddLogoToPdf(doc) {
  // Optional: try to embed displayed logo if accessible
  try {
    const imgEl = document.querySelector("img.logo");
    if (!imgEl) return;
    // Wait if not loaded
    if (!imgEl.complete) await new Promise(res => imgEl.addEventListener("load", res, { once: true }));
    const dataUrl = imageToDataUrl(imgEl);
    if (!dataUrl) return;
    doc.addImage(dataUrl, "PNG", 170, 10, 25, 25);
  } catch {
    // ignore if fails
  }
}

function imageToDataUrl(imgEl) {
  try {
    const canvas = document.createElement("canvas");
    canvas.width = imgEl.naturalWidth || 200;
    canvas.height = imgEl.naturalHeight || 200;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(imgEl, 0, 0);
    return canvas.toDataURL("image/png");
  } catch {
    return null;
  }
}

/* ---------------------------
   Text for selections
---------------------------- */
function zonesSelectedText(d) {
  const labels = [];
  if (d.z_fr) labels.push("France");
  if (d.z_ueuk) labels.push("UE/UK");
  if (d.z_canada) labels.push("Canada");
  if (d.z_usau) labels.push("Australie/USA");
  if (d.z_asie) labels.push("Japon/Corée/Taiwan/Singapour");
  if (d.z_chineinde) labels.push("Chine/Inde");
  if (d.z_autres) labels.push("Autres");
  return labels.join(", ") || "—";
}

function activitiesSelectedText() {
  const map = [
    ["act_essai","Essai clinique"],
    ["act_prom","CLB promoteur"],
    ["act_data","Recherche données/échantillons"],
    ["act_ech","Échanges données/échantillons"],
    ["act_drci","Besoin expertise DRCI"],
    ["act_clin","Expertise clinicien"],
    ["act_cher","Expertise chercheur"],
    ["act_st","Sous-traitance prévue"],
    ["act_achat","Achat équipement prévu"],
    ["act_eq","Plusieurs équipes impliquées"],
    ["act_heberg","Personnel externe hébergé"],
    ["act_pf","Utilisation plateformes internes"],
  ];
  const out = map.filter(([k]) => form.elements[k]?.checked).map(([,lbl]) => lbl);
  return out.join(", ") || "—";
}

function platformsSelectedText(d) {
  const map = [
    ["pf_pic","PIC"],
    ["pf_par","PAR"],
    ["pf_cyto","Cytométrie"],
    ["pf_geno","Génomique"],
    ["pf_bioinfo","Bioinformatique"],
    ["pf_preclin","Modèles précliniques"],
    ["pf_autre","Autre"],
  ];
  const out = map.filter(([k]) => d[k]).map(([,lbl]) => lbl);
  return out.join(", ") || "—";
}

/* ---------------------------
   Small utilities
---------------------------- */
function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function showLastSaved(iso) {
  if (!iso) { lastSavedEl.textContent = ""; return; }
  const d = new Date(iso);
  lastSavedEl.textContent = `Dernière sauvegarde : ${d.toLocaleString("fr-FR")}`;
}

/* ---------------------------
   Init
---------------------------- */
(function init() {
  handleConditionalFields();
  recomputeAll();

  // Optionnel: charge auto si existe
  const raw = localStorage.getItem(LS_KEY);
  if (raw) {
    try {
      const payload = JSON.parse(raw);
      showLastSaved(payload.savedAt);
    } catch {}
  }
})();
