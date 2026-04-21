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
const q6Guidance = $("#q6Guidance");
const mr004Block = $("#mr004Block");
const cmtBlock = $("#cmtBlock");
const mr004StateEl = $("#mr004State");
const cmtStateEl = $("#cmtState");
const q6SiteOther = $("#q6_site_other");
const q6SiteOtherWrap = $("#q6SiteOtherWrap");
const mrPopulationOther = $("#mr_population_other");
const mrPopulationOtherWrap = $("#mrPopulationOtherWrap");
const mrRetentionWrap = $("#mrRetentionWrap");
const mrSensitiveWrap = $("#mrSensitiveWrap");
const cmtPfOther = $("#cmt_pf_other");
const cmtPfOtherWrap = $("#cmtPfOtherWrap");

const fundingTableBody = $("#fundingTable tbody");

$("#btnAddFunding").addEventListener("click", addFundingRow);
$("#btnSave").addEventListener("click", saveLocal);
$("#btnLoad").addEventListener("click", loadLocal);
$("#btnReset").addEventListener("click", resetAll);
$("#btnPdf").addEventListener("click", generatePdf);
$("#btnQ6Docs").addEventListener("click", generateQ6Documents);

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
  if (q6SiteOtherWrap) q6SiteOtherWrap.style.display = q6SiteOther?.checked ? "block" : "none";
  if (mrPopulationOtherWrap) mrPopulationOtherWrap.style.display = mrPopulationOther?.checked ? "block" : "none";
  if (mrRetentionWrap) mrRetentionWrap.style.display = form.elements.mr_retention?.value === "Durée spécifique" ? "block" : "none";
  if (mrSensitiveWrap) mrSensitiveWrap.style.display = form.elements.mr_sensitive?.value === "Oui" ? "block" : "none";
  if (cmtPfOtherWrap) cmtPfOtherWrap.style.display = cmtPfOther?.checked ? "block" : "none";

  const needMr004 = !!form.elements.trf_data?.checked;
  const needCmt = !!form.elements.trf_ech?.checked;

  if (mr004Block) mr004Block.style.display = needMr004 ? "block" : "none";
  if (cmtBlock) cmtBlock.style.display = needCmt ? "block" : "none";
  if (mr004StateEl) mr004StateEl.textContent = needMr004 ? "requise" : "non requise";
  if (cmtStateEl) cmtStateEl.textContent = needCmt ? "requise" : "non requise";
  if (q6Guidance) q6Guidance.textContent = getQ6GuidanceMessage(showQ6, needMr004, needCmt);

  if (!showQ6) {
    [
      "trf_ech", "trf_data", "trf_multi",
      "q6_site_clb", "q6_site_ihope", "q6_site_other",
      "mr_population_patients", "mr_population_aidants", "mr_population_pros", "mr_population_other",
      "mr_flow_ecrf", "mr_flow_owncloud", "mr_flow_mss",
      "mr_data_pathology", "mr_data_treatments", "mr_data_genetics", "mr_data_imaging", "mr_data_slides", "mr_data_samples",
      "cmt_matching",
      "cmt_pf_biopath", "cmt_pf_par", "cmt_pf_pgeb", "cmt_pf_onco3d", "cmt_pf_pgc", "cmt_pf_pgt", "cmt_pf_licl", "cmt_pf_pathec", "cmt_pf_other"
    ].forEach((name) => {
      if (form.elements[name]) form.elements[name].checked = false;
    });
    [
      "trf_zone", "trf_dest", "trf_desc", "trf_com",
      "q6_resp_scientifique", "q6_contact", "q6_site_other_txt", "q6_period", "q6_contract", "q6_objective",
      "mr_population_other_txt", "mr_population_desc", "mr_population_counts", "mr_pathology", "mr_case", "mr_flow_other",
      "mr_data_other", "mr_legal_basis", "mr_retention", "mr_retention_detail", "mr_sensitive", "mr_information_mode",
      "mr_sensitive_detail", "mr_pseudonymisation", "mr_contract_status", "mr_ethics_need",
      "cmt_clinician", "cmt_partnerships", "cmt_funding", "cmt_selection_note", "cmt_sample_type", "cmt_sample_site",
      "cmt_sample_pathology", "cmt_sample_count", "cmt_criteria_clinical", "cmt_criteria_quality", "cmt_summary",
      "cmt_data_list", "cmt_data_objective", "cmt_ethics_impact", "cmt_pf_other_txt"
    ].forEach((name) => {
      if (form.elements[name]) form.elements[name].value = "";
    });
  }

  if (showQ6) applyQ6Prefill();
}

function getQ6GuidanceMessage(showQ6, needMr004, needCmt) {
  if (!showQ6) {
    return "Activez un transfert de données et/ou un transfert d’échantillons pour afficher les exigences associées.";
  }
  if (needMr004 && needCmt) {
    return "Transfert de données et d’échantillons détecté : les volets MR004 et CMT sont requis. Seuls les champs projet partagés sont repris automatiquement.";
  }
  if (needMr004) {
    return "Transfert de données détecté : la fiche MR004 est requise. Seuls les champs projet partagés sont repris automatiquement.";
  }
  if (needCmt) {
    return "Transfert d’échantillons détecté : la fiche CMT est requise. Seuls les champs projet partagés sont repris automatiquement.";
  }
  return "La Q6 est ouverte car la Q3 mentionne des données ou échantillons. Cochez le ou les transferts réellement nécessaires pour afficher la fiche MR004 et/ou la fiche CMT.";
}

function applyQ6Prefill() {
  const d = getFormData();
  const canonical = buildCanonicalQ6Data(d);

  setIfEmpty("q6_resp_scientifique", canonical.scientificLead);
  setIfEmpty("q6_contact", canonical.operationalContact);
  setIfEmpty("trf_dest", canonical.destination);
  setIfEmpty("trf_zone", canonical.zone);
  setIfEmpty("q6_period", canonical.period);
  setIfEmpty("q6_objective", canonical.objective);
  setIfEmpty("cmt_partnerships", canonical.partnerships);
  setIfEmpty("cmt_funding", canonical.funding);

  if (canonical.siteClb && form.elements.q6_site_clb && !form.elements.q6_site_clb.checked) form.elements.q6_site_clb.checked = true;
  if (canonical.siteIhope && form.elements.q6_site_ihope && !form.elements.q6_site_ihope.checked) form.elements.q6_site_ihope.checked = true;
  if (canonical.siteOther && form.elements.q6_site_other && !form.elements.q6_site_other.checked) form.elements.q6_site_other.checked = true;
  setIfEmpty("q6_site_other_txt", canonical.siteOtherText);

  if (q6SiteOtherWrap) q6SiteOtherWrap.style.display = q6SiteOther?.checked ? "block" : "none";
  if (mrPopulationOtherWrap) mrPopulationOtherWrap.style.display = mrPopulationOther?.checked ? "block" : "none";
}

function setIfEmpty(name, value) {
  if (value === undefined || value === null || value === "") return;
  const el = form.elements[name];
  if (!el || el.type === "checkbox") return;
  if (!String(el.value || "").trim()) el.value = value;
}

function buildCanonicalQ6Data(d) {
  const zone = deriveZoneFromMainForm(d);
  const partnerTypes = [
    d.p_acad ? "académique" : null,
    d.p_indus ? "industriel" : null,
    d.p_multi ? "multicentrique" : null,
  ].filter(Boolean).join(", ");

  const fundingBits = [
    d.fin_exist ? "financement existant" : null,
    d.fin_ext ? "financement extérieur" : null,
    d.fin_valide ? "financement validé" : null,
    d.fin_montant ? `${formatEuros(d.fin_montant)}` : null,
  ].filter(Boolean);

  const objective = [
    d.resume ? d.resume.trim() : "",
    d.clin ? "Application clinique à court terme." : "",
    d.pub ? "Valorisation scientifique attendue." : "",
  ].filter(Boolean).join(" ");

  const flowParts = [
    d.act_data ? "Recherche sur données/échantillons" : null,
    d.act_ech ? "Échanges de données/échantillons" : null,
    d.trf_multi || d.p_multi ? "Projet multi-sites / multicentrique" : null,
    d.coord ? `Coordination: ${d.coord}` : null,
  ].filter(Boolean);

  const comments = [
    d.aap ? `AAP: ${[d.aap_nom, d.aap_org, d.aap_statut].filter(Boolean).join(" / ")}` : null,
    d.fin_cp_nom ? `Chef de projet: ${d.fin_cp_nom}` : null,
  ].filter(Boolean).join(" | ");

  return {
    scientificLead: d.porteur || "",
    operationalContact: [d.porteur, d.email].filter(Boolean).join(" — "),
    destination: d.coord || "",
    zone,
    period: derivePeriod(d),
    contract: deriveContract(d),
    objective,
    flowDescription: flowParts.join(" ; "),
    comments,
    siteClb: true,
    siteIhope: false,
    siteOther: d.z_autres || d.nbPart !== "0",
    siteOtherText: d.z_autres_txt || "",
    populationDescription: derivePopulationDescription(d),
    populationCounts: "",
    pathology: "",
    mrCase: deriveMrCase(d),
    collectionTool: deriveCollectionTool(d),
    dataOther: d.trf_desc || "",
    informationMode: d.pub ? "Notice d’information / information des personnes à prévoir" : "",
    contractStatus: deriveContract(d),
    ethicsNeed: d.trf_ech ? "Coordination avec le CMT si transfert d’échantillons" : "",
    pseudonymisation: "A confirmer : pseudonymisation des données, limitation des accès et stockage sur des outils sécurisés.",
    populationPatients: true,
    populationProfessionals: false,
    dataSamples: !!(d.trf_ech || d.act_ech),
    partnerships: [partnerTypes, d.coord].filter(Boolean).join(" ; "),
    funding: fundingBits.join(" ; "),
    cmtSummary: d.resume || "",
    cmtDataList: deriveCmtDataList(d),
  };
}

function deriveZoneFromMainForm(d) {
  if (d.trf_zone) return d.trf_zone;
  if (d.z_autres) return "Autres";
  if (d.z_chineinde) return "Chine/Inde";
  if (d.z_asie) return "Japon/Corée/Taiwan/Singapour";
  if (d.z_usau) return "Australie/USA";
  if (d.z_canada) return "Canada";
  if (d.z_ueuk) return "UE/UK";
  if (d.z_fr) return "France";
  return "";
}

function derivePeriod(d) {
  if (d.date_debut || d.date_fin) return [d.date_debut || "?", d.date_fin || "?"].join(" → ");
  return "";
}

function deriveContract(d) {
  if (d.aap || d.p_indus || d.act_st) {
    return [d.p_indus ? "partenariat industriel" : null, d.act_st ? "sous-traitance" : null, d.aap ? `AAP ${d.aap_nom || ""}`.trim() : null]
      .filter(Boolean)
      .join(" ; ");
  }
  return "";
}

function derivePopulationDescription(d) {
  if (d.clin || d.act_essai) return "Patients inclus dans le projet / l’étude";
  return "Population à confirmer";
}

function deriveMrCase(d) {
  if (d.p_multi || d.trf_multi) return "Cas multicentrique à préciser";
  if (d.p_indus) return "Cas avec partenaire industriel";
  if (d.p_acad) return "Cas avec partenaire académique";
  return "Cas interne CLB à confirmer";
}

function deriveCollectionTool(d) {
  if (d.act_data) return "Outil de recueil / hébergeur à confirmer (ex. Excel sécurisé, REDCap CLB, e-CRF)";
  return "";
}

function deriveCmtDataList(d) {
  const items = [];
  if (d.act_data || d.trf_data) items.push("données cliniques associées");
  if (d.act_cher) items.push("données de recherche / analyses prévues");
  if (d.p_multi) items.push("données issues de plusieurs sites");
  return items.join(", ");
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
    "trf_ech","trf_data","trf_multi",
    "q6_site_clb","q6_site_ihope","q6_site_other",
    "mr_population_patients","mr_population_aidants","mr_population_pros","mr_population_other",
    "mr_flow_ecrf","mr_flow_owncloud","mr_flow_mss",
    "mr_data_pathology","mr_data_treatments","mr_data_genetics","mr_data_imaging","mr_data_slides","mr_data_samples",
    "cmt_matching",
    "cmt_pf_biopath","cmt_pf_par","cmt_pf_pgeb","cmt_pf_onco3d","cmt_pf_pgc","cmt_pf_pgt","cmt_pf_licl","cmt_pf_pathec","cmt_pf_other"
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
  const d = requireCoreProjectData();
  if (!d) return;

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
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
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
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
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
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
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
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
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
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
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
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] }
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
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
    columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  // Q6 table
  if (d.act_data || d.act_ech) {
    doc.autoTable({
      startY: doc.lastAutoTable.finalY + 6,
      head: [["Q6 — Données / échantillons (socle commun)", ""]],
      body: [
        ["Transfert échantillons", d.trf_ech ? "Oui" : "Non"],
        ["Transfert données", d.trf_data ? "Oui" : "Non"],
        ["Multicentrique", d.trf_multi ? "Oui" : "Non"],
        ["Responsable scientifique CLB", safe(d.q6_resp_scientifique)],
        ["Acteur opérationnel / contact", safe(d.q6_contact)],
        ["Sites source", q6SitesSelectedText(d)],
        ["Zone destinataire", safe(d.trf_zone)],
        ["Destinataire", safe(d.trf_dest)],
        ["Période concernée", safe(d.q6_period)],
        ["Contrat / support juridique", safe(d.q6_contract)],
        ["Objectif du traitement", safe(d.q6_objective)],
        ["Description des flux", safe(d.trf_desc)],
        ["Commentaires / sécurité", safe(d.trf_com)],
      ],
      styles: { fontSize: 9, cellPadding: 2 },
      headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
      columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
      didParseCell: (data) => { data.cell.styles.valign = "top"; }
    });

    if (d.trf_data) {
      doc.autoTable({
        startY: doc.lastAutoTable.finalY + 4,
        head: [["MR004 — Champs spécifiques", ""]],
        body: [
          ["Population", mrPopulationSelectedText(d)],
          ["Description population", safe(d.mr_population_desc)],
          ["Nombre concerné / CLB", safe(d.mr_population_counts)],
          ["Pathologie / tumeur", safe(d.mr_pathology)],
          ["Cas MR004", safe(d.mr_case)],
          ["Circulation / recueil", mrFlowSelectedText(d, d.mr_flow_other)],
          ["Catégories de données", mrDataSelectedText(d, d.mr_data_other)],
          ["Fondement juridique", safe(d.mr_legal_basis)],
          ["Conservation", [safe(d.mr_retention), safe(d.mr_retention_detail)].filter(v => v !== "—").join(" — ") || "—"],
          ["Données sensibles", [safe(d.mr_sensitive), safe(d.mr_sensitive_detail)].filter(v => v !== "—").join(" — ") || "—"],
          ["Information des personnes", safe(d.mr_information_mode)],
          ["Pseudonymisation", safe(d.mr_pseudonymisation)],
          ["Contrat / juridique", safe(d.mr_contract_status)],
          ["Certificat / avis éthique", safe(d.mr_ethics_need)],
        ],
        styles: { fontSize: 9, cellPadding: 2 },
        headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
        columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
        didParseCell: (data) => { data.cell.styles.valign = "top"; }
      });
    }

    if (d.trf_ech) {
      doc.autoTable({
        startY: doc.lastAutoTable.finalY + 4,
        head: [["CMT — Champs spécifiques", ""]],
        body: [
          ["Clinicien impliqué", safe(d.cmt_clinician)],
          ["Collaborations / partenariats", safe(d.cmt_partnerships)],
          ["Financement", safe(d.cmt_funding)],
          ["Précision sélection", safe(d.cmt_selection_note)],
          ["Type d’échantillon", safe(d.cmt_sample_type)],
          ["Organe / localisation", safe(d.cmt_sample_site)],
          ["Pathologie", safe(d.cmt_sample_pathology)],
          ["Nombre demandé", safe(d.cmt_sample_count)],
          ["Critères clinico-biologiques", safe(d.cmt_criteria_clinical)],
          ["Critères quantitatifs / qualitatifs", safe(d.cmt_criteria_quality)],
          ["Appariement", d.cmt_matching ? "Oui" : "Non"],
          ["Résumé CMT", safe(d.cmt_summary)],
          ["Données associées", safe(d.cmt_data_list)],
          ["Objectif traitement associé", safe(d.cmt_data_objective)],
          ["Impact éthique", safe(d.cmt_ethics_impact)],
          ["Plateformes / expertises", cmtPlatformsSelectedText(d)],
        ],
        styles: { fontSize: 9, cellPadding: 2 },
        headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
        columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 130 } },
        didParseCell: (data) => { data.cell.styles.valign = "top"; }
      });
    }
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

async function generateQ6Documents() {
  const d = requireCoreProjectData();
  if (!d) return;

  if (!d.trf_data && !d.trf_ech) {
    alert("Cochez d'abord 'Transfert de données' et/ou 'Transfert d'échantillons biologiques' dans la Q6.");
    return;
  }

  if (d.trf_data) await generateMr004Pdf(d);
  if (d.trf_ech) await generateCmtPdf(d);
}

function requireCoreProjectData() {
  const d = getFormData();
  const missing = [];
  if (!String(d.titre || "").trim()) missing.push("Titre du projet");
  if (!String(d.porteur || "").trim()) missing.push("Porteur");
  if (!String(d.email || "").trim()) missing.push("Email");
  if (!String(d.resume || "").trim()) missing.push("Résumé");

  if (missing.length) {
    alert("Champs requis manquants :\n- " + missing.join("\n- "));
    return null;
  }
  return d;
}

async function generateMr004Pdf(d = getFormData()) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  await tryAddLogoToPdf(doc);

  let y = 14;
  doc.setFontSize(15);
  doc.text("Declaration d'un traitement de donnees - MR004", 14, y);
  y += 6;
  doc.setFontSize(10);
  doc.text("Version PDF issue de la fiche projet CLB", 14, y);

  doc.autoTable({
    startY: y + 6,
    head: [["Informations generales", ""]],
    body: [
      ["Date de la fiche", new Date().toLocaleDateString("fr-FR")],
      ["Titre du projet", safe(d.titre)],
      ["Porteur / responsable scientifique", safe(d.porteur)],
      ["Email", safe(d.email)],
      ["Unite / etablissement", safe(d.unite)],
      ["Partenaires / coordination", [partnerTypesText(d), safe(d.coord)].filter(v => v !== "—").join(" ; ") || "—"],
      ["Sites concernes", q6SitesSelectedText(d)],
      ["Periode concernee", safe(d.q6_period || derivePeriod(d))],
      ["Zone destinataire / hebergement", safe(d.trf_zone || deriveZoneFromMainForm(d))],
      ["Resume / finalite du projet", safe(d.resume)],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
    columnStyles: { 0: { cellWidth: 58 }, 1: { cellWidth: 127 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 5,
    head: [["Description scientifique du projet", ""]],
    body: [
      ["Objectif du traitement", safe(d.q6_objective || d.resume)],
      ["Population concernee", mrPopulationSelectedText(d)],
      ["Description de la population", safe(d.mr_population_desc)],
      ["Nombre concerne / nombre CLB", safe(d.mr_population_counts)],
      ["Type de tumeur / pathologie", safe(d.mr_pathology)],
      ["Description generale des flux", safe(d.trf_desc)],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
    columnStyles: { 0: { cellWidth: 58 }, 1: { cellWidth: 127 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 5,
    head: [["Description RGPD et securite", ""]],
    body: [
      ["Cas MR004 / categorisation", safe(d.mr_case)],
      ["Mode de circulation des donnees", mrFlowSelectedText(d, d.mr_flow_other)],
      ["Categories de donnees traitees", mrDataSelectedText(d, d.mr_data_other)],
      ["Fondement juridique", safe(d.mr_legal_basis)],
      ["Duree de conservation", [safe(d.mr_retention), safe(d.mr_retention_detail)].filter(v => v !== "—").join(" — ") || "—"],
      ["Donnees sensibles", [safe(d.mr_sensitive), safe(d.mr_sensitive_detail)].filter(v => v !== "—").join(" — ") || "—"],
      ["Pseudonymisation", safe(d.mr_pseudonymisation)],
      ["Mesures de securite / commentaires", safe(d.trf_com)],
      ["Contrat redige / validation juridique", safe(d.mr_contract_status)],
      ["Besoin de certification RGPD / avis ethique", safe(d.mr_ethics_need)],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
    columnStyles: { 0: { cellWidth: 58 }, 1: { cellWidth: 127 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  const finalY = doc.lastAutoTable.finalY + 14;
  doc.text("Signature porteur :", 14, finalY);
  doc.line(45, finalY, 120, finalY);
  doc.text("Date :", 135, finalY);
  doc.line(147, finalY, 195, finalY);

  doc.save(`MR004_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`);
}

async function generateCmtPdf(d = getFormData()) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  await tryAddLogoToPdf(doc);

  let y = 14;
  doc.setFontSize(15);
  doc.text("Fiche CMT - Demande d'echantillons", 14, y);
  y += 6;
  doc.setFontSize(10);
  doc.text("Version PDF issue de la fiche projet CLB", 14, y);

  doc.autoTable({
    startY: y + 6,
    head: [["En-tete projet", ""]],
    body: [
      ["Date de la demande", new Date().toLocaleDateString("fr-FR")],
      ["Titre du projet", safe(d.titre)],
      ["Porteur(s)", safe(d.porteur)],
      ["Laboratoire / Etablissement", safe(d.unite)],
      ["Coordonnees", safe(d.email)],
      ["Collaborations / partenariats", safe(d.cmt_partnerships) !== "—" ? safe(d.cmt_partnerships) : ([partnerTypesText(d), safe(d.coord)].filter(v => v !== "—").join(" ; ") || "—")],
      ["Clinicien implique", safe(d.cmt_clinician)],
      ["Financement du projet", safe(d.cmt_funding)],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
    columnStyles: { 0: { cellWidth: 58 }, 1: { cellWidth: 127 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 5,
    head: [["Description precise des echantillons demandes", ""]],
    body: [
      ["Type d'echantillon", safe(d.cmt_sample_type)],
      ["Organe / localisation", safe(d.cmt_sample_site)],
      ["Pathologie", safe(d.cmt_sample_pathology)],
      ["Nombre", safe(d.cmt_sample_count)],
      ["Criteres clinico-biologiques", safe(d.cmt_criteria_clinical)],
      ["Criteres quantitatifs / qualitatifs", safe(d.cmt_criteria_quality)],
      ["A apparier", d.cmt_matching ? "Oui" : "Non"],
      ["Autre precision de selection", safe(d.cmt_selection_note)],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
    columnStyles: { 0: { cellWidth: 58 }, 1: { cellWidth: 127 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 5,
    head: [["Projet, donnees associees et ethique", ""]],
    body: [
      ["Resume CMT / rationnel / analyses prevues", safe(d.cmt_summary) !== "—" ? safe(d.cmt_summary) : safe(d.resume)],
      ["Donnees associees a collecter", safe(d.cmt_data_list)],
      ["Objectif du traitement de donnees", safe(d.cmt_data_objective)],
      ["Impact ethique", safe(d.cmt_ethics_impact)],
      ["Plateformes / expertises sollicitees", cmtPlatformsSelectedText(d)],
      ["Commentaires complementaires", safe(d.trf_com)],
    ],
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [245,245,245], textColor: [15,23,42] },
    columnStyles: { 0: { cellWidth: 58 }, 1: { cellWidth: 127 } },
    didParseCell: (data) => { data.cell.styles.valign = "top"; }
  });

  const finalY = doc.lastAutoTable.finalY + 14;
  doc.text("Avis CMT / signature :", 14, finalY);
  doc.line(49, finalY, 125, finalY);
  doc.text("Date :", 138, finalY);
  doc.line(150, finalY, 195, finalY);

  doc.save(`CMT_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`);
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

function partnerTypesText(d) {
  const out = [
    d.p_acad ? "Academique" : null,
    d.p_indus ? "Industriel" : null,
    d.p_multi ? "Multicentrique" : null,
  ].filter(Boolean);
  return out.join(", ") || "—";
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

function q6SitesSelectedText(d) {
  const out = [
    d.q6_site_clb ? "CLB" : null,
    d.q6_site_ihope ? "IHOPe" : null,
    d.q6_site_other ? (d.q6_site_other_txt || "Autre site") : null,
  ].filter(Boolean);
  return out.join(", ") || "—";
}

function mrPopulationSelectedText(d) {
  const out = [
    d.mr_population_patients ? "Patients" : null,
    d.mr_population_aidants ? "Aidants" : null,
    d.mr_population_pros ? "Professionnels de santé" : null,
    d.mr_population_other ? (d.mr_population_other_txt || "Autre") : null,
  ].filter(Boolean);
  return out.join(", ") || "—";
}

function mrFlowSelectedText(d, other) {
  const out = [
    d.mr_flow_ecrf ? "e-CRF" : null,
    d.mr_flow_owncloud ? "Owncloud CLB" : null,
    d.mr_flow_mss ? "MSS / monSISRA" : null,
    other || null,
  ].filter(Boolean);
  return out.join(", ") || "—";
}

function mrDataSelectedText(d, other) {
  const out = [
    d.mr_data_pathology ? "Pathologie" : null,
    d.mr_data_treatments ? "Traitements" : null,
    d.mr_data_genetics ? "Génétiques" : null,
    d.mr_data_imaging ? "Imagerie" : null,
    d.mr_data_slides ? "Lames anapath" : null,
    d.mr_data_samples ? "Échantillons biologiques" : null,
    other || null,
  ].filter(Boolean);
  return out.join(", ") || "—";
}

function cmtPlatformsSelectedText(d) {
  const out = [
    d.cmt_pf_biopath ? "BIOPATH" : null,
    d.cmt_pf_par ? "PAR" : null,
    d.cmt_pf_pgeb ? "PGEB" : null,
    d.cmt_pf_onco3d ? "Onco-3D" : null,
    d.cmt_pf_pgc ? "PGC" : null,
    d.cmt_pf_pgt ? "PGT" : null,
    d.cmt_pf_licl ? "LICL" : null,
    d.cmt_pf_pathec ? "PATHEC" : null,
    d.cmt_pf_other ? (d.cmt_pf_other_txt || "Autre expertise") : null,
  ].filter(Boolean);
  return out.join(", ") || "—";
}

function initFormDoc(doc) {
  doc.setDrawColor(120, 128, 140);
  doc.setTextColor(20, 24, 32);
  doc.setLineWidth(0.25);
}

function ensurePageSpace(doc, y, needed = 18) {
  if (y + needed <= 280) return y;
  doc.addPage();
  return 14;
}

function drawDocTitle(doc, y, title, subtitle = "") {
  y = ensurePageSpace(doc, y, 22);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.text(title, 14, y);
  if (subtitle) {
    y += 6;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(subtitle, 14, y);
  }
  doc.setLineWidth(0.5);
  doc.line(14, y + 4, 196, y + 4);
  doc.setLineWidth(0.25);
  return y + 8;
}

function drawIntroBox(doc, y, text) {
  y = ensurePageSpace(doc, y, 18);
  const lines = doc.splitTextToSize(text, 176);
  const h = 6 + lines.length * 4.4;
  doc.roundedRect(14, y, 182, h, 2, 2);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  doc.text(lines, 18, y + 5);
  return y + h + 5;
}

function drawSectionHeader(doc, y, title) {
  y = ensurePageSpace(doc, y, 12);
  doc.setFillColor(236, 239, 244);
  doc.rect(14, y, 182, 8, "F");
  doc.rect(14, y, 182, 8);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text(title, 17, y + 5.5);
  return y + 11;
}

function drawFieldBox(doc, y, label, value = "", height = 12) {
  y = ensurePageSpace(doc, y, height + 8);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9.2);
  doc.text(label, 14, y);
  const top = y + 2;
  doc.rect(14, top, 182, height);
  if (String(value || "").trim() && value !== "—") {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    const lines = doc.splitTextToSize(String(value), 174);
    doc.text(lines, 17, top + 5);
  }
  return top + height + 5;
}

function drawTwoColumnFields(doc, y, fields) {
  y = ensurePageSpace(doc, y, 20);
  const colW = 88;
  const gap = 6;
  const x1 = 14;
  const x2 = x1 + colW + gap;
  fields.slice(0, 2).forEach(([label, value], idx) => {
    const x = idx === 0 ? x1 : x2;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9.2);
    doc.text(label, x, y);
    doc.rect(x, y + 2, colW, 11);
    if (String(value || "").trim() && value !== "—") {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      const lines = doc.splitTextToSize(String(value), colW - 6);
      doc.text(lines, x + 3, y + 7);
    }
  });
  return y + 18;
}

function drawGridTable(doc, y, headers, rows, widths) {
  doc.autoTable({
    startY: y,
    head: [headers],
    body: rows,
    styles: { fontSize: 8.2, cellPadding: 2, valign: "top", lineColor: [120,128,140], lineWidth: 0.15 },
    headStyles: { fillColor: [236,239,244], textColor: [15,23,42], fontStyle: "bold", lineColor: [120,128,140] },
    columnStyles: Object.fromEntries(widths.map((w, i) => [i, { cellWidth: w }])),
    theme: "grid",
  });
  return doc.lastAutoTable.finalY;
}

function drawSignatureArea(doc, y, title) {
  y = ensurePageSpace(doc, y, 20);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text(title, 14, y);
  doc.setFont("helvetica", "normal");
  doc.text("Date :", 138, y);
  doc.line(45, y, 122, y);
  doc.line(150, y, 195, y);
  return y + 10;
}

async function generateMr004Pdf(d = getFormData()) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  await tryAddLogoToPdf(doc);
  initFormDoc(doc);

  let y = 14;
  y = drawDocTitle(doc, y, "Déclaration d'un traitement de données", "MR004 / analyse simplifiée des risques");
  y = drawIntroBox(doc, y, "A renseigner et à transmettre au DPD du CLB. Cette version PDF cherche à retrouver la logique et la mise en page des fiches d'origine avec de grands blocs de réponse.");

  y = drawSectionHeader(doc, y, "Fiche de traitement de données");
  y = drawTwoColumnFields(doc, y, [
    ["Date d'envoi de la fiche au DPD", new Date().toLocaleDateString("fr-FR")],
    ["Traitement en lien avec des patients", mrPopulationSelectedText(d)],
  ]);
  y = drawFieldBox(doc, y, "Entité(s) responsable(s) du projet / porteur(s) du projet", safe(d.porteur), 11);
  y = drawFieldBox(doc, y, "Responsable scientifique sur le site du CLB", safe(d.porteur), 11);
  y = drawFieldBox(doc, y, "Acteur opérationnel (interne, chercheur...) sur le site du CLB", safe(d.email) !== "—" ? `${safe(d.porteur)} — ${safe(d.email)}` : safe(d.porteur), 11);
  y = drawFieldBox(doc, y, "Nom du projet - titre et acronyme", safe(d.titre), 12);
  y = drawFieldBox(doc, y, "Numéro EDS / équivalent si applicable", "", 10);

  y = drawSectionHeader(doc, y, "I. Description scientifique du projet");
  y = drawFieldBox(doc, y, "Objectif du projet (finalité du traitement de données)", safe(d.q6_objective || d.resume), 16);
  y = drawFieldBox(doc, y, "Description grand public", "", 18);
  y = drawFieldBox(doc, y, "Description détaillée", safe(d.resume), 26);
  y = drawFieldBox(doc, y, "Catégorie de population concernée", mrPopulationSelectedText(d), 10);
  y = drawFieldBox(doc, y, "Description de la population", safe(d.mr_population_desc), 18);
  y = drawTwoColumnFields(doc, y, [
    ["Nombre concerné", safe(d.mr_population_counts)],
    ["Type de tumeur / pathologie", safe(d.mr_pathology)],
  ]);
  y = drawFieldBox(doc, y, "Période concernée", safe(d.q6_period || derivePeriod(d)), 10);

  y = drawSectionHeader(doc, y, "II. Description RGPD et sécurité");
  y = drawFieldBox(doc, y, "Catégorisation du projet (cas 1 à 6 ou autre)", safe(d.mr_case), 12);
  y = drawFieldBox(doc, y, "Centres / sites / partenaires concernés", [q6SitesSelectedText(d), partnerTypesText(d), safe(d.coord)].filter(v => v !== "—").join(" ; ") || "—", 12);
  y = drawFieldBox(doc, y, "Circulation des données / outil de recueil / hébergement", mrFlowSelectedText(d, d.mr_flow_other), 16);
  y = drawFieldBox(doc, y, "Catégories de données traitées", mrDataSelectedText(d, d.mr_data_other), 18);
  y = drawFieldBox(doc, y, "Fondement juridique", safe(d.mr_legal_basis), 10);
  y = drawFieldBox(doc, y, "Durée de conservation des données", [safe(d.mr_retention), safe(d.mr_retention_detail)].filter(v => v !== "—").join(" — ") || "—", 12);
  y = drawFieldBox(doc, y, "Données sensibles traitées", [safe(d.mr_sensitive), safe(d.mr_sensitive_detail)].filter(v => v !== "—").join(" — ") || "—", 14);
  y = drawFieldBox(doc, y, "Pseudonymisation", safe(d.mr_pseudonymisation), 18);
  y = drawFieldBox(doc, y, "Mesures de sécurité / commentaires", safe(d.trf_com), 18);
  y = drawFieldBox(doc, y, "Contrat rédigé / validation juridique", safe(d.mr_contract_status), 12);

  y = drawSectionHeader(doc, y, "III. Information des personnes concernées");
  y = drawFieldBox(doc, y, "Mode d'information prévu", safe(d.mr_information_mode), 16);

  y = drawSectionHeader(doc, y, "IV. Demande de certificat d'instruction et/ou d'avis éthique");
  y = drawFieldBox(doc, y, "Besoin de certification RGPD / avis éthique", safe(d.mr_ethics_need), 16);

  y = drawSectionHeader(doc, y, "V. Avis du délégué à la protection des données");
  y = drawFieldBox(doc, y, "Partie réservée au DPD", "", 28);

  y = drawSectionHeader(doc, y, "VI. Description pour le Comité data");
  y = drawFieldBox(doc, y, "Axe valorisation de la donnée / partage des données à l'issue du projet", "", 24);

  drawSignatureArea(doc, y, "Signature porteur");
  doc.save(`MR004_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`);
}

async function generateCmtPdf(d = getFormData()) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  await tryAddLogoToPdf(doc);
  initFormDoc(doc);

  let y = 14;
  y = drawDocTitle(doc, y, "Fiche CMT", "Demande d'échantillons");
  y = drawIntroBox(doc, y, "Mise en forme inspirée de la fiche d'origine avec grands blocs de saisie, zones structurées et espaces de validation.");

  y = drawTwoColumnFields(doc, y, [
    ["Date de la demande", new Date().toLocaleDateString("fr-FR")],
    ["Clinicien impliqué", safe(d.cmt_clinician)],
  ]);
  y = drawFieldBox(doc, y, "TITRE du projet", safe(d.titre), 12);
  y = drawFieldBox(doc, y, "Porteur(s)", safe(d.porteur), 10);
  y = drawFieldBox(doc, y, "Laboratoire / Etablissement", safe(d.unite), 10);
  y = drawFieldBox(doc, y, "Coordonnées", safe(d.email), 10);
  y = drawFieldBox(doc, y, "Collaboration(s) / Partenariat(s)", safe(d.cmt_partnerships) !== "—" ? safe(d.cmt_partnerships) : ([partnerTypesText(d), safe(d.coord)].filter(v => v !== "—").join(" ; ") || "—"), 12);
  y = drawFieldBox(doc, y, "Financement du projet", safe(d.cmt_funding), 10);

  y = drawSectionHeader(doc, y, "Description précise des échantillons demandés");
  y = drawGridTable(doc, y, [
    "Type d'échantillon",
    "Organe / localisation",
    "Pathologie",
    "Nombre",
    "Critères clinico-biologiques",
    "Critères quantitatifs / qualitatifs",
    "A apparier",
  ], [[
    safe(d.cmt_sample_type),
    safe(d.cmt_sample_site),
    safe(d.cmt_sample_pathology),
    safe(d.cmt_sample_count),
    safe(d.cmt_criteria_clinical),
    safe(d.cmt_criteria_quality),
    d.cmt_matching ? "Oui" : "Non",
  ]], [24, 24, 22, 14, 40, 40, 18]);
  y = doc.lastAutoTable.finalY + 3;
  y = drawFieldBox(doc, y, "Autre précision utile pour la sélection des échantillons", safe(d.cmt_selection_note), 14);

  y = drawSectionHeader(doc, y, "Résumé du projet");
  y = drawFieldBox(doc, y, "Rationnel, finalité de l'utilisation des échantillons, analyses prévues", safe(d.cmt_summary) !== "—" ? safe(d.cmt_summary) : safe(d.resume), 28);

  y = drawSectionHeader(doc, y, "Traitement des données");
  y = drawFieldBox(doc, y, "Liste des données associées à collecter", safe(d.cmt_data_list), 18);
  y = drawFieldBox(doc, y, "Objectif du traitement de données", safe(d.cmt_data_objective), 16);

  y = drawSectionHeader(doc, y, "Ethique");
  y = drawFieldBox(doc, y, "Impact(s) éthique(s) potentiel(s) pour le patient ou la société", safe(d.cmt_ethics_impact), 22);

  y = drawSectionHeader(doc, y, "Plateformes technologiques ou expertises sollicitées");
  y = drawFieldBox(doc, y, "Plateformes / expertises", cmtPlatformsSelectedText(d), 14);

  y = drawSectionHeader(doc, y, "Suivi et validation");
  y = drawFieldBox(doc, y, "Commentaires / réserves / suivi", safe(d.trf_com), 22);

  drawSignatureArea(doc, y, "Pour le Comité Médico-Technique");
  doc.save(`CMT_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`);
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
