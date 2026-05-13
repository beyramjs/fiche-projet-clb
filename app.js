/* ===========================
   Fiche Projet CLB — app.js
   - Scores Q1→Q5 selon règles fournies
   - Alertes dynamiques
   - Historique financements (table)
   - Sauvegarde/chargement localStorage
   - Génération PDF via jsPDF + autotable
   =========================== */

const LS_KEY = "ficheProjet_v1";
const REGISTRY_KEY = "ficheProjet_registry_v1";

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
const MAIL_RECIPIENT = "beyram.frigui@lyon.unicancer.fr";
const MAIL_SERVICE_URL = "http://127.0.0.1:8765/send-mail";
const MAIL_SERVICE_DOCX_URL = "http://127.0.0.1:8765/send-docx";

const $ = (sel) => document.querySelector(sel);

const form = $("#ficheForm");
const alertsEl = $("#alerts");
const lastSavedEl = $("#lastSaved");
const projectCodeBadge = $("#projectCodeBadge");
const registryModal = $("#registryModal");
const btnRegistryClose = $("#btnRegistryClose");
const registryTbody = $("#registryTbody");
const btnExportRegistry2 = $("#btnExportRegistry2");
const btnExportProjectJson = $("#btnExportProjectJson");
const projectJsonInput = $("#projectJsonInput");
const localModeHint = $("#localModeHint");

// Some strings were introduced via copy/paste from Word/PDF and can end up mojibake-encoded.
// We normalize at render/export time so the UI/PDF stays clean even if a few literals slip in.
function fixText(v) {
  if (typeof v !== "string") return v;
  let s = v;
  const rep = [
    ["Ã¢â‚¬â€", "—"],
    ["â€”", "—"],
    ["â€“", "–"],
    ["â€™", "’"],
    ["Â ", ""],

    ["ÃƒÂ©", "é"],
    ["ÃƒÂ¨", "è"],
    ["ÃƒÂª", "ê"],
    ["ÃƒÂ«", "ë"],
    ["ÃƒÂ ", "à"],
    ["ÃƒÂ¢", "â"],
    ["ÃƒÂ´", "ô"],
    ["ÃƒÂ®", "î"],
    ["ÃƒÂ¯", "ï"],
    ["ÃƒÂ§", "ç"],
    ["Ãƒâ€°", "É"],
    ["ÃƒÅ“", "Œ"],
    ["ÃƒÅ“", "Œ"],
    ["Å“", "œ"],

    ["Ã©", "é"],
    ["Ã¨", "è"],
    ["Ãª", "ê"],
    ["Ã«", "ë"],
    ["Ã ", "à"],
    ["Ã¢", "â"],
    ["Ã´", "ô"],
    ["Ã®", "î"],
    ["Ã¯", "ï"],
    ["Ã§", "ç"],
    ["Ã‰", "É"],
  ];
  for (const [from, to] of rep) s = s.split(from).join(to);
  return s;
}

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
const q6Section = $("#q6");
const q6Guidance = $("#q6Guidance");
const mr004Block = $("#mr004Block");
const mr004RegulatoryBlock = $("#mr004RegulatoryBlock");
const cmtBlock = $("#cmtBlock");
const mr004StateEl = $("#mr004State");
const cmtStateEl = $("#cmtState");
const q6SiteOther = $("#q6_site_other");
const q6SiteOtherWrap = $("#q6SiteOtherWrap");
const mrPopulationOther = $("#mr_population_other");
const mrPopulationOtherWrap = $("#mrPopulationOtherWrap");
const mrTumorType = $("#mr_tumor_type");
const mrTumorOtherWrap = $("#mrTumorOtherWrap");
const mrCaseOtherWrap = $("#mrCaseOtherWrap");
const mrCase1Wrap = $("#mrCase1Wrap");
const mrCase2Wrap = $("#mrCase2Wrap");
const mrCase2to6Wrap = $("#mrCase2to6Wrap");
const mrCase3to6Wrap = $("#mrCase3to6Wrap");
const mrRetentionWrap = $("#mrRetentionWrap");
const mrSensitiveWrap = $("#mrSensitiveWrap");
const mrDataHealthOther = $("#mr_data_health_other");
const mrDataHealthOtherWrap = $("#mrDataHealthOtherWrap");
const mrDataNonHealthOther = $("#mr_data_nonhealth_other");
const mrDataNonHealthOtherWrap = $("#mrDataNonHealthOtherWrap");
const mrLegalOtherWrap = $("#mrLegalOtherWrap");
const mrSensitiveWrap2 = $("#mrSensitiveWrap2");
const mrRetentionWrap2 = $("#mrRetentionWrap2");
const mrInfoPatientOther = $("#mr_info_patient_other");
const mrPatientInfoOtherWrap = $("#mrPatientInfoOtherWrap");
const mrInfoNonPatientOther = $("#mr_info_nonpatient_other");
const mrNonPatientInfoOtherWrap = $("#mrNonPatientInfoOtherWrap");
const mrEthicsOpinion = $("#mr_ethics_opinion");
const mrEthicsOpinionWrap = $("#mrEthicsOpinionWrap");
const cmtPfOther = $("#cmt_pf_other");
const cmtPfOtherWrap = $("#cmtPfOtherWrap");

const fundingTableBody = $("#fundingTable tbody");

deconflictLegacyMr004Fields();
normalizeFormPresentation();

$("#btnAddFunding")?.addEventListener("click", addFundingRow);
$("#btnSave")?.addEventListener("click", () => {
  // Save to local registry and open the registry modal to allow re-opening/editing later.
  const d = getFormData();
  upsertRegistryRow(d, "save");
  openRegistryModal();
});

$("#btnGenerateProject")?.addEventListener("click", generatePdf);
$("#btnGenerateCmt")?.addEventListener("click", () => downloadDocx("cmt"));
$("#btnGenerateMr004")?.addEventListener("click", () => downloadDocx("mr004"));
$("#btnSendCmtDoc")?.addEventListener("click", () => sendDocumentEmailAutomaticallyLegacy("cmt"));
$("#btnSendMr004Doc")?.addEventListener("click", () => sendDocumentEmailAutomaticallyLegacy("mr004"));

btnExportRegistry2?.addEventListener("click", exportRegistryCsv);
btnExportProjectJson?.addEventListener("click", exportCurrentProjectJson);
projectJsonInput?.addEventListener("change", importProjectJsonFromFile);
btnRegistryClose?.addEventListener("click", closeRegistryModal);

form.addEventListener("input", () => {
  handleConditionalFields();
  recomputeAll();
});

form.addEventListener("change", () => {
  handleConditionalFields();
  recomputeAll();
});

ensureProjectCode();
renderProjectCodeBadge();
refreshRegistryPicker();
initLocalModeHintsAndBridge();

let bridgeMode = { available: false, baseUrl: "" };

async function initLocalModeHintsAndBridge() {
  // If page is served via https (GitHub Pages), browsers block fetch to http://127.0.0.1 (mixed content).
  // So Word sending + disk registry require opening the app in local mode (served over http by the python bridge).
  const isHttps = location.protocol === "https:";
  const localUrl = "http://127.0.0.1:8765/";

  if (localModeHint) {
    localModeHint.textContent =
      isHttps
        ? `Mode local requis pour Outlook/registre disque : lancez "python .\\outlook_mail_bridge.py" puis ouvrez ${localUrl}`
        : "Mode local actif si le service Python tourne sur ce PC.";
  }

  // Detect bridge availability (works when the page is served from the same origin http://127.0.0.1:8765)
  try {
    const res = await fetch("/health", { method: "GET" });
    if (res.ok) {
      bridgeMode = { available: true, baseUrl: location.origin };
      // Prefer server-side registry when available
      await refreshRegistryPickerFromBridge();
    }
  } catch {
    bridgeMode = { available: false, baseUrl: "" };
  }
}

async function downloadDocx(docType) {
  const d = requireCoreProjectData();
  if (!d) return;

  if (location.protocol === "https:") {
    alert(
      "Generation Word indisponible depuis GitHub Pages (HTTPS).\n\n" +
      "Solution : lancez le service local puis ouvrez http://127.0.0.1:8765/"
    );
    return;
  }

  try {
    const payload = { docType, data: getFormData() };
    const res = await fetch("/generate-docx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `HTTP ${res.status}`);
    }
    const blob = await res.blob();
    const filename = docType === "mr004"
      ? `MR004_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.docx`
      : `CMT_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.docx`;
    downloadBlob(blob, filename);
    upsertRegistryRow(d, docType === "mr004" ? "generate_mr004_docx" : "generate_cmt_docx");
  } catch (e) {
    alert(`Generation Word impossible: ${e.message}`);
  }
}

function handleConditionalFields() {
  if (autresWrap) autresWrap.style.display = zAutres?.checked ? "block" : "none";
  if (pfAutreWrap) pfAutreWrap.style.display = pfAutre?.checked ? "block" : "none";
  if (cpWrap) cpWrap.style.display = finCp?.checked ? "block" : "none";
  if (aapWrap) aapWrap.style.display = aap?.checked ? "grid" : "none";

  const showQ6 = !!(form.elements.act_data?.checked || form.elements.act_ech?.checked);
  if (q6Section) q6Section.style.display = showQ6 ? "block" : "none";
  if (q6SiteOtherWrap) q6SiteOtherWrap.style.display = q6SiteOther?.checked ? "block" : "none";
  if (mrPopulationOtherWrap) mrPopulationOtherWrap.style.display = mrPopulationOther?.checked ? "block" : "none";
  if (mrTumorOtherWrap) mrTumorOtherWrap.style.display = mrTumorType?.value === "other" ? "block" : "none";
  const mrCaseValue = form.elements.mr_case?.value || "";
  if (mrCaseOtherWrap) mrCaseOtherWrap.style.display = mrCaseValue === "other" ? "block" : "none";
  if (mrCase1Wrap) mrCase1Wrap.style.display = mrCaseValue === "1" ? "block" : "none";
  if (mrCase2Wrap) mrCase2Wrap.style.display = mrCaseValue === "2" ? "block" : "none";
  if (mrCase2to6Wrap) mrCase2to6Wrap.style.display = ["2", "3", "4", "5", "6"].includes(mrCaseValue) ? "block" : "none";
  if (mrCase3to6Wrap) mrCase3to6Wrap.style.display = ["3", "4", "5", "6"].includes(mrCaseValue) ? "block" : "none";
  if (mrRetentionWrap) mrRetentionWrap.style.display = form.elements.mr_retention?.value === "Durée spécifique" ? "block" : "none";
  if (mrSensitiveWrap) mrSensitiveWrap.style.display = form.elements.mr_sensitive?.value === "Oui" ? "block" : "none";
  if (mrDataHealthOtherWrap) mrDataHealthOtherWrap.style.display = mrDataHealthOther?.checked ? "block" : "none";
  if (mrDataNonHealthOtherWrap) mrDataNonHealthOtherWrap.style.display = mrDataNonHealthOther?.checked ? "block" : "none";
  if (mrLegalOtherWrap) mrLegalOtherWrap.style.display = form.elements.mr_legal_basis?.value === "other" ? "block" : "none";
  if (mrSensitiveWrap2) mrSensitiveWrap2.style.display = form.elements.mr_sensitive?.value === "yes" ? "block" : "none";
  if (mrRetentionWrap2) mrRetentionWrap2.style.display = form.elements.mr_retention?.value === "other" ? "block" : "none";
  if (mrPatientInfoOtherWrap) mrPatientInfoOtherWrap.style.display = mrInfoPatientOther?.checked ? "block" : "none";
  if (mrNonPatientInfoOtherWrap) mrNonPatientInfoOtherWrap.style.display = mrInfoNonPatientOther?.checked ? "block" : "none";
  if (mrEthicsOpinionWrap) mrEthicsOpinionWrap.style.display = mrEthicsOpinion?.checked ? "block" : "none";
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
      "mr_case_doc_risk", "mr_case_doc_notice", "mr_case_doc_ecrf",
      "mr_flow_ecrf", "mr_flow_owncloud", "mr_flow_mss",
      "mr_data_pathology", "mr_data_treatments", "mr_data_genetics_somatic", "mr_data_genetics_germline", "mr_data_imaging", "mr_data_slides", "mr_data_samples", "mr_data_pgeb_contacted", "mr_data_health_other", "mr_data_social", "mr_data_nonhealth_other",
      "mr_info_patient_level5", "mr_info_patient_notice", "mr_info_patient_other",
      "mr_info_nonpatient_oral", "mr_info_nonpatient_written", "mr_info_nonpatient_other",
      "mr_ethics_certificate", "mr_ethics_opinion",
      "cmt_matching",
      "cmt_pf_biopath", "cmt_pf_par", "cmt_pf_pgeb", "cmt_pf_onco3d", "cmt_pf_pgc", "cmt_pf_pgt", "cmt_pf_licl", "cmt_pf_pathec", "cmt_pf_other"
    ].forEach((name) => {
      if (form.elements[name]) form.elements[name].checked = false;
    });
    [
      "trf_zone", "trf_dest", "trf_desc", "trf_com",
      "q6_resp_scientifique", "q6_contact", "q6_site_other_txt", "q6_period", "q6_contract", "q6_objective",
      "mr_population_other_txt", "mr_population_desc", "mr_population_counts", "mr_tumor_type", "mr_tumor_other_txt", "mr_case", "mr_case_other_detail",
      "mr_case1_internal_teams", "mr_case1_subcontractors", "mr_case1_stats_team",
      "mr_case2_centers", "mr_case2_internal_teams", "mr_case2_subcontractors", "mr_case2_stats_team",
      "mr_case36_internal_teams", "mr_case_host_country", "mr_case_collection_responsible", "mr_case_collection_tool", "mr_flow_other",
      "mr_data_health_other_txt", "mr_data_nonhealth_other_txt", "mr_data_schema", "mr_legal_basis", "mr_legal_basis_other", "mr_retention", "mr_retention_detail", "mr_sensitive",
      "mr_sensitive_types", "mr_sensitive_justification", "mr_questionnaire_tool", "mr_pseudonymisation", "mr_identity_removal", "mr_contract_drafted", "mr_contract_validated_by",
      "mr_info_patient_other_txt", "mr_info_nonpatient_other_txt", "mr_ethics_topic",
      "cmt_clinician", "cmt_partnerships", "cmt_funding", "cmt_selection_note", "cmt_sample_type", "cmt_sample_site",
      "cmt_sample_pathology", "cmt_sample_count", "cmt_criteria_clinical", "cmt_criteria_quality", "cmt_criteria_quantitative", "cmt_criteria_quality_detail", "cmt_matching_detail", "cmt_summary",
      "cmt_data_list", "cmt_data_objective", "cmt_ethics_impact", "cmt_pf_other_txt", "cmt_platform_details"
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

function deconflictLegacyMr004Fields() {
  if (!mr004Block || !mr004RegulatoryBlock) return;
  const legacyNames = [
    "mr_data_pathology", "mr_data_treatments", "mr_data_genetics", "mr_data_imaging", "mr_data_slides", "mr_data_samples",
    "mr_data_other", "mr_legal_basis", "mr_retention", "mr_retention_detail", "mr_sensitive", "mr_information_mode",
    "mr_sensitive_detail", "mr_pseudonymisation", "mr_contract_status", "mr_ethics_need"
  ];
  const hidden = new Set();
  legacyNames.forEach((name) => {
    mr004Block.querySelectorAll(`[name="${name}"]`).forEach((el) => {
      if (mr004RegulatoryBlock.contains(el)) return;
      el.name = `legacy_${name}`;
      const container = el.closest("label, div, fieldset");
      if (container && !hidden.has(container)) {
        container.style.display = "none";
        hidden.add(container);
      }
    });
  });
}

function normalizeFormPresentation() {
  const straySpan = Array.from(document.querySelectorAll("span")).find((el) => {
    const text = (el.textContent || "").toLowerCase();
    return text.includes("plateformes ou expertises") && !el.closest("#cmtBlock");
  });
  if (straySpan) straySpan.remove();

  const cmtNote = document.querySelector("#cmtBlock .small");
  if (cmtNote) {
    cmtNote.textContent = "Une participation forfaitaire aux frais du CRB est demandée. Cette section suit la logique de la fiche CMT d'origine.";
  }

  const cmtCriteriaLabel = document.querySelector('textarea[name="cmt_criteria_quality"]')?.closest("label")?.querySelector("span");
  if (cmtCriteriaLabel) {
    cmtCriteriaLabel.textContent = "Critères de sélection complémentaires";
  }

  const cmtMatchingLabel = document.querySelector('input[name="cmt_matching"]')?.closest("label")?.querySelector("span");
  if (cmtMatchingLabel) {
    cmtMatchingLabel.textContent = "Échantillons à apparier";
  }
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
    "mr_tumor_all","mr_tumor_brain","mr_tumor_colorectal","mr_tumor_stomach","mr_tumor_liver","mr_tumor_small_intestine","mr_tumor_eye","mr_tumor_orl","mr_tumor_bone","mr_tumor_ovary","mr_tumor_pancreas","mr_tumor_skin","mr_tumor_pleura","mr_tumor_lung","mr_tumor_prostate","mr_tumor_kidney","mr_tumor_hematology","mr_tumor_breast","mr_tumor_testicle","mr_tumor_thyroid","mr_tumor_uterus","mr_tumor_bladder","mr_tumor_soft_tissue","mr_tumor_other_solid","mr_tumor_unknown_primary","mr_tumor_other",
    "mr_flow_ecrf","mr_flow_owncloud","mr_flow_mss",
    "mr_data_pathology","mr_data_treatments","mr_data_genetics_somatic","mr_data_genetics_germline","mr_data_imaging","mr_data_slides","mr_data_samples","mr_data_pgeb_contacted","mr_data_health_other","mr_data_social","mr_data_nonhealth_other",
    "mr_info_patient_level5","mr_info_patient_notice","mr_info_patient_other",
    "mr_info_nonpatient_oral","mr_info_nonpatient_written","mr_info_nonpatient_other",
    "mr_ethics_certificate","mr_ethics_opinion",
    "cmt_matching",
    "cmt_pf_biopath","cmt_pf_par","cmt_pf_pgeb","cmt_pf_onco3d","cmt_pf_pgc","cmt_pf_pgt","cmt_pf_licl","cmt_pf_pathec","cmt_pf_other"
  ];
  checkNames.forEach((n) => (obj[n] = !!form.elements[n]?.checked));

  // numbers
  obj.fin_montant = toNumber(obj.fin_montant);
  obj.fin_frais = toNumber(obj.fin_frais);

  // funding table rows
  obj.fundingHistory = readFundingTable();

  obj.mr_legal_basis = fixText(mrLegalBasisText(obj));
  obj.mr_retention = fixText(mrRetentionText(obj));
  obj.mr_sensitive_detail = fixText(mrSensitiveText(obj));
  obj.mr_information_mode = fixText(mrInformationText(obj));
  obj.mr_contract_status = fixText(mrContractText(obj));
  obj.mr_ethics_need = fixText(mrEthicsNeedText(obj));
  obj.mr_data_other = obj.mr_data_schema || "";
  obj.cmt_criteria_quality = [obj.cmt_criteria_quantitative, obj.cmt_criteria_quality_detail || obj.cmt_criteria_quality].filter(Boolean).join(" | ");
  if (obj.cmt_matching_detail) {
    obj.cmt_selection_note = [obj.cmt_selection_note, `Appariement: ${obj.cmt_matching_detail}`].filter(Boolean).join(" | ");
  }
  if (obj.cmt_platform_details) {
    obj.cmt_pf_other_txt = [obj.cmt_pf_other_txt, obj.cmt_platform_details].filter(Boolean).join(" | ");
  }

  return obj;
}

function ensureProjectCode() {
  if (!form) return "";
  let code = String(localStorage.getItem("ficheProjet_current_code") || "").trim();
  if (!code) {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const d = String(now.getDate()).padStart(2, "0");
    const r = Math.random().toString(36).slice(2, 6).toUpperCase();
    code = `FP-${y}${m}${d}-${r}`;
    localStorage.setItem("ficheProjet_current_code", code);
  }
  return code;
}

function renderProjectCodeBadge() {
  if (!projectCodeBadge) return;
  projectCodeBadge.textContent = ensureProjectCode() || "—";
}

function getRegistry() {
  try {
    const raw = localStorage.getItem(REGISTRY_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function upsertRegistryRow(d, source = "manual") {
  const code = ensureProjectCode();
  const nowIso = new Date().toISOString();
  const registry = getRegistry();
  const row = {
    code,
    created_at: registry.find((r) => r.code === code)?.created_at || nowIso,
    updated_at: nowIso,
    source,
    titre: d.titre || "",
    porteur: d.porteur || "",
    email: d.email || "",
    unite: d.unite || "",
    mr004: !!d.trf_data,
    cmt: !!d.trf_ech,
    zone: d.trf_zone || "",
    snapshot: d, // full form state for re-opening/editing later
  };
  const idx = registry.findIndex((r) => r.code === code);
  if (idx >= 0) registry[idx] = row;
  else registry.unshift(row);
  localStorage.setItem(REGISTRY_KEY, JSON.stringify(registry));
  renderProjectCodeBadge();
  refreshRegistryPicker();
  renderRegistryTableIfOpen();

  // Also persist to disk registry if the local bridge is available
  persistRegistryRowToBridge(row).catch(() => {});
}

function exportRegistryCsv() {
  // If local bridge is available, export the disk-based registry (shared across browser sessions on this PC).
  if (bridgeMode.available) {
    window.open("/registry/export.csv", "_blank");
    return;
  }

  const registry = getRegistry();
  const headers = ["code", "created_at", "updated_at", "source", "titre", "porteur", "email", "unite", "mr004", "cmt", "zone"];
  const lines = [headers.join(",")];
  registry.forEach((r) => {
    const vals = headers.map((h) => {
      const v = r[h];
      const s = v === undefined || v === null ? "" : String(v);
      // CSV escaping
      return `"${s.replaceAll('"', '""')}"`;
    });
    lines.push(vals.join(","));
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  downloadBlob(blob, `registre_projets_${new Date().toISOString().slice(0, 10)}.csv`);
}

function refreshRegistryPicker() {
  // Kept for backward compatibility (older UI had a select). Now we only use the modal table.
}

function openRegistryModal() {
  if (!registryModal) return;
  registryModal.showModal();
  renderRegistryTable();
}

function closeRegistryModal() {
  if (!registryModal) return;
  registryModal.close();
}

function renderRegistryTableIfOpen() {
  if (!registryModal) return;
  if (!registryModal.open) return;
  renderRegistryTable();
}

async function renderRegistryTable() {
  if (!registryTbody) return;
  registryTbody.innerHTML = "";

  let rows = [];
  if (bridgeMode.available) {
    const res = await fetch("/registry/list");
    rows = (await res.json().catch(() => [])) || [];
  } else {
    rows = getRegistry();
  }

  if (!Array.isArray(rows) || rows.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="7" class="small" style="padding:14px;">Aucun projet pour le moment.</td>`;
    registryTbody.appendChild(tr);
    return;
  }

  rows.forEach((r) => {
    const tr = document.createElement("tr");
    const code = String(r.code || "").trim();
    const updated = String(r.updated_at || "").slice(0, 16).replace("T", " ");
    tr.innerHTML = `
      <td>${escapeHtml(code)}</td>
      <td>${escapeHtml(r.titre || "")}</td>
      <td>${escapeHtml(r.porteur || "")}</td>
      <td>${r.mr004 ? "Oui" : "Non"}</td>
      <td>${r.cmt ? "Oui" : "Non"}</td>
      <td>${escapeHtml(updated)}</td>
      <td><button type="button" class="secondary" data-open-code="${escapeAttr(code)}">Ouvrir</button></td>
    `;
    registryTbody.appendChild(tr);
  });
}

function escapeHtml(s) {
  return String(s || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function escapeAttr(s) {
  return escapeHtml(s).replaceAll("'", "&#39;");
}

registryTbody?.addEventListener("click", (e) => {
  const btn = e.target?.closest?.("[data-open-code]");
  if (!btn) return;
  const code = String(btn.getAttribute("data-open-code") || "").trim();
  if (!code) return;
  loadProjectByCode(code);
});

async function loadProjectByCode(code) {
  if (bridgeMode.available) {
    try {
      await loadSnapshotFromBridge(code);
      closeRegistryModal();
      return;
    } catch (err) {
      alert(`Chargement impossible: ${err.message}`);
      return;
    }
  }

  const registry = getRegistry();
  const row = registry.find((r) => r.code === code);
  if (!row || !row.snapshot) {
    alert("Projet introuvable dans le registre (ou snapshot manquant).");
    return;
  }
  localStorage.setItem("ficheProjet_current_code", code);
  applyDataToForm(row.snapshot);
  handleConditionalFields();
  recomputeAll();
  saveLocal();
  upsertRegistryRow(getFormData(), "load_registry");
  closeRegistryModal();
  showAlert(`Projet ${code} chargé.`, "ok");
}

async function persistRegistryRowToBridge(row) {
  if (!bridgeMode.available) return;
  await fetch("/registry/upsert", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(row),
  });
}

async function refreshRegistryPickerFromBridge() {
  // Older UI had a select picker. Now we only use the modal table.
  renderRegistryTableIfOpen();
  return;
  if (!registrySelect) return;
  const res = await fetch("/registry/list");
  if (!res.ok) return;
  const items = await res.json().catch(() => []);
  if (!Array.isArray(items)) return;
  registrySelect.innerHTML = `<option value="">Charger depuis le registre…</option>`;
  items.forEach((r) => {
    const opt = document.createElement("option");
    opt.value = r.code || "";
    opt.textContent = [r.code, r.titre].filter(Boolean).join(" — ") || r.code || "(sans code)";
    registrySelect.appendChild(opt);
  });
}

async function loadSnapshotFromBridge(code) {
  const res = await fetch(`/registry/get?code=${encodeURIComponent(code)}`);
  const payload = await res.json().catch(() => null);
  if (!res.ok || !payload || !payload.snapshot) throw new Error(payload?.error || `HTTP ${res.status}`);
  localStorage.setItem("ficheProjet_current_code", code);
  applyDataToForm(payload.snapshot);
  handleConditionalFields();
  recomputeAll();
  saveLocal();
  upsertRegistryRow(getFormData(), "load_registry");
  showAlert(`Projet ${code} chargé depuis le registre disque.`, "ok");
}

function exportCurrentProjectJson() {
  const d = getFormData();
  const payload = {
    code: ensureProjectCode(),
    exported_at: new Date().toISOString(),
    data: d,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
  downloadBlob(blob, `fiche_projet_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}_${payload.code}.json`);
  upsertRegistryRow(d, "export_json");
}

async function importProjectJsonFromFile(e) {
  const file = e.target?.files?.[0];
  if (!file) return;
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const data = parsed?.data || parsed;
    if (!data || typeof data !== "object") throw new Error("Fichier JSON invalide");
    const code = String(parsed?.code || data?.code || "").trim();
    if (code) localStorage.setItem("ficheProjet_current_code", code);
    applyDataToForm(data);
    handleConditionalFields();
    recomputeAll();
    saveLocal();
    upsertRegistryRow(getFormData(), "import_json");
  } catch (err) {
    alert(`Import impossible : ${err.message}`);
  } finally {
    e.target.value = "";
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
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
  upsertRegistryRow(d, "save_local");
  renderProjectCodeBadge();
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
  if (payload?.data && typeof payload.data === "object") upsertRegistryRow(getFormData(), "load_local");
  renderProjectCodeBadge();
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
  localStorage.removeItem("ficheProjet_current_code");
  renderProjectCodeBadge();
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
  const d = requireCoreProjectData();
  if (!d) return;
  const doc = await buildProjectPdfDoc(d);
  const filename = `FicheProjet_CLB_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`;
  doc.save(filename);
  upsertRegistryRow(d, "generate_project_pdf");
}

async function buildProjectPdfDoc(d) {
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
          ["Pathologie / tumeur", mrTumorSelectedText(d)],
          ["Cas MR004", mrCaseLabelText(d)],
          ["Détails de catégorisation", mrCaseDetailsText(d)],
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

  return doc;
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
  upsertRegistryRow(d, "generate_q6_pdfs");
}

async function sendDocumentEmailAutomaticallyLegacy(docType) {
  const d = requireCoreProjectData();
  if (!d) return;

  if (location.protocol === "https:") {
    alert(
      "Envoi Word automatique indisponible depuis GitHub Pages (HTTPS).\n\n" +
      "Le navigateur bloque l'accès au service local Outlook (HTTP) : c'est normal.\n\n" +
      "Solution :\n" +
      "1) Lancez : python .\\outlook_mail_bridge.py\n" +
      "2) Ouvrez : http://127.0.0.1:8765/\n" +
      "3) Refaite l'envoi depuis ce mode local."
    );
    return;
  }

  if (docType === "mr004" && !d.trf_data) {
    alert("Cochez d'abord 'Transfert de données' dans la Q6 pour envoyer la fiche MR004.");
    return;
  }

  if (docType === "cmt" && !d.trf_ech) {
    alert("Cochez d'abord 'Transfert d'échantillons biologiques' dans la Q6 pour envoyer la fiche CMT.");
    return;
  }

  try {
    const payload = {
      to: MAIL_RECIPIENT,
      subject: `${docType === "mr004" ? "MR004" : "CMT"} - ${d.titre || "Fiche projet CLB"}`,
      body: buildDocumentEmailBody(d, docType),
      docType,
      data: getFormData(),
    };

    const response = await fetch(MAIL_SERVICE_DOCX_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    const result = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(result.error || `Erreur HTTP ${response.status}`);
    }
    upsertRegistryRow(d, docType === "mr004" ? "send_mr004_docx" : "send_cmt_docx");
    upsertRegistryRow(d, docType === "mr004" ? "send_mr004" : "send_cmt");

    alert(`Email envoyé automatiquement à ${MAIL_RECIPIENT}.`);
  } catch (error) {
    alert(
      "Envoi automatique impossible.\n\n" +
      "Vérifiez que le service local Outlook est démarré avec :\n" +
      "python .\\outlook_mail_bridge.py\n\n" +
      "Puis ouvrez l'application en mode local :\n" +
      "http://127.0.0.1:8765/\n\n" +
      `Détail : ${error.message}`
    );
  }
}

async function sendDocumentEmailAutomatically(docType) {
  const d = requireCoreProjectData();
  if (!d) return;

  if (docType === "mr004" && !d.trf_data) {
    alert("Cochez d'abord 'Transfert de données' dans la Q6 pour envoyer la fiche MR004.");
    return;
  }

  if (docType === "cmt" && !d.trf_ech) {
    alert("Cochez d'abord 'Transfert d'échantillons biologiques' dans la Q6 pour envoyer la fiche CMT.");
    return;
  }

  try {
    const attachments = [];

    if (docType === "mr004") {
      const mrDoc = await buildMr004PdfDoc(d);
      attachments.push({
        filename: `MR004_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`,
        blob: mrDoc.output("blob"),
      });
    }

    if (docType === "cmt") {
      const cmtDoc = await buildCmtPdfDoc(d);
      attachments.push({
        filename: `CMT_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`,
        blob: cmtDoc.output("blob"),
      });
    }

    const formData = new FormData();
    formData.append("to", MAIL_RECIPIENT);
    formData.append("subject", `${docType === "mr004" ? "MR004" : "CMT"} - ${d.titre || "Fiche projet CLB"}`);
    formData.append("body", buildDocumentEmailBody(d, docType));
    attachments.forEach((file, index) => {
      formData.append(`attachment_${index}`, file.blob, file.filename);
    });

    const response = await fetch(MAIL_SERVICE_URL, {
      method: "POST",
      body: formData,
    });

    const result = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(result.error || `Erreur HTTP ${response.status}`);
    }

    alert(`Fiche ${docType === "mr004" ? "MR004" : "CMT"} envoyée automatiquement à ${MAIL_RECIPIENT}.`);
  } catch (error) {
    alert(
      "Envoi automatique impossible.\n\n" +
      "Vérifiez que le service local Outlook est démarré avec :\n" +
      "python .\\outlook_mail_bridge.py\n\n" +
      `Détail : ${error.message}`
    );
  }
}

function buildDocumentEmailBody(d, docType) {
  const docLabel = docType === "mr004" ? "fiche MR004" : "fiche CMT";
  return [
    "Bonjour,",
    "",
    `Veuillez trouver ci-joint la ${docLabel} générée pour le projet.`,
    "",
    `Titre du projet : ${d.titre || "—"}`,
    `Porteur : ${d.porteur || "—"}`,
    `Email : ${d.email || "—"}`,
    "",
    "Cordialement,",
  ].join("\n");
}

function buildAutomaticEmailBody(d) {
  return [
    "Bonjour,",
    "",
    "Veuillez trouver ci-joint les documents générés pour la fiche projet.",
    "",
    `Titre du projet : ${d.titre || "—"}`,
    `Porteur : ${d.porteur || "—"}`,
    `Email porteur : ${d.email || "—"}`,
    `Unité : ${d.unite || "—"}`,
    `Partenaires / coordination : ${[partnerTypesText(d), safe(d.coord)].filter(v => v !== "—").join(" ; ") || "—"}`,
    `Transfert de données : ${d.trf_data ? "Oui" : "Non"}`,
    `Transfert d'échantillons : ${d.trf_ech ? "Oui" : "Non"}`,
    "",
    "Cordialement,",
  ].join("\n");
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
  doc.text(`Code projet : ${ensureProjectCode()}`, 14, y);
  y += 5;
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
      ["Type de tumeur / pathologie", mrTumorSelectedText(d)],
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
      ["Cas MR004 / categorisation", mrCaseLabelText(d)],
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
    doc.setFillColor(255, 255, 255);
    doc.roundedRect(168, 4, 24, 24, 2, 2, "F");
    doc.addImage(dataUrl, "PNG", 169, 5, 22, 22);
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

function mrTumorSelectedText(d) {
  if (d.mr_tumor_type === "other") return d.mr_tumor_other_txt || "Autre tumeur";
  return d.mr_tumor_type || "—";
}

function mrCaseLabelText(d) {
  const labels = {
    "1": "Cas 1 : pilotage CLB / IHOPe, monocentrique",
    "2": "Cas 2 : pilotage CLB / IHOPe, multicentrique",
    "3": "Cas 3 : pilotage académique, hébergement UE",
    "4": "Cas 4 : pilotage industriel, hébergement UE",
    "5": "Cas 5 : hébergement hors UE avec décision d'adéquation",
    "6": "Cas 6 : hébergement hors UE sans décision d'adéquation",
    "other": "Autre cas à détailler",
  };
  return labels[d.mr_case] || "—";
}

function mrCaseDetailsText(d) {
  const parts = [];
  if (d.mr_case === "other" && d.mr_case_other_detail) parts.push(`Détail: ${d.mr_case_other_detail}`);
  if (d.mr_case === "1") {
    if (d.mr_case1_internal_teams) parts.push(`Équipes internes: ${d.mr_case1_internal_teams}`);
    if (d.mr_case1_subcontractors) parts.push(`Sous-traitants: ${d.mr_case1_subcontractors}`);
    if (d.mr_case1_stats_team) parts.push(`Statistiques: ${d.mr_case1_stats_team}`);
  }
  if (d.mr_case === "2") {
    if (d.mr_case2_centers) parts.push(`Centres participants: ${d.mr_case2_centers}`);
    if (d.mr_case2_internal_teams) parts.push(`Équipes internes: ${d.mr_case2_internal_teams}`);
    if (d.mr_case2_subcontractors) parts.push(`Sous-traitants: ${d.mr_case2_subcontractors}`);
    if (d.mr_case2_stats_team) parts.push(`Statistiques: ${d.mr_case2_stats_team}`);
  }
  if (["3", "4", "5", "6"].includes(d.mr_case || "")) {
    if (d.mr_case36_internal_teams) parts.push(`Équipes internes: ${d.mr_case36_internal_teams}`);
    const docs = [
      d.mr_case_doc_risk ? "Analyse des risques / AIPD" : null,
      d.mr_case_doc_notice ? "Note d'information" : null,
      d.mr_case_doc_ecrf ? "E-CRF détaillé" : null,
    ].filter(Boolean);
    if (docs.length) parts.push(`Pièces jointes: ${docs.join(", ")}`);
    if (d.mr_case_host_country) parts.push(`Pays d'hébergement: ${d.mr_case_host_country}`);
  }
  if (["2", "3", "4", "5", "6"].includes(d.mr_case || "")) {
    const flow = mrFlowSelectedText(d, d.mr_flow_other);
    if (flow !== "—") parts.push(`Circulation des données: ${flow}`);
  }
  if (d.mr_case_collection_responsible) parts.push(`Recueil des données: ${d.mr_case_collection_responsible}`);
  if (d.mr_case_collection_tool) parts.push(`Outil / sécurité: ${d.mr_case_collection_tool}`);
  return parts.join(" | ") || "—";
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

function mrDataSelectedTextLegacy(d, other) {
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

function mrDataRegulatoryText(d) {
  const out = [
    d.mr_data_pathology ? "Description de la pathologie" : null,
    d.mr_data_treatments ? "Description des traitements" : null,
    d.mr_data_genetics_somatic ? "DonnÃ©es gÃ©nÃ©tiques somatiques" : null,
    d.mr_data_genetics_germline ? "DonnÃ©es gÃ©nÃ©tiques constitutionnelles" : null,
    d.mr_data_imaging ? "DonnÃ©es d'imagerie (DICOM)" : null,
    d.mr_data_slides ? "Lames virtuelles d'anapath" : null,
    d.mr_data_samples ? "Ã‰chantillons biologiques humains" : null,
    d.mr_data_pgeb_contacted ? "PGEB du CLB contactÃ©e" : null,
    d.mr_data_health_other ? (d.mr_data_health_other_txt || "Autres donnÃ©es de santÃ©") : null,
    d.mr_data_social ? "DonnÃ©es sociales ou relatives au mode de vie" : null,
    d.mr_data_nonhealth_other ? (d.mr_data_nonhealth_other_txt || "Autres donnÃ©es hors santÃ©") : null,
  ].filter(Boolean);
  if (d.mr_data_schema) out.push(`SchÃ©ma gÃ©nÃ©ral: ${d.mr_data_schema}`);
  return out.join(", ") || "â€”";
}

function mrLegalBasisText(d) {
  const labels = {
    mission: "ExÃ©cution d'une mission d'intÃ©rÃªt public",
    legitimate: "Poursuite d'un intÃ©rÃªt lÃ©gitime",
    legal_obligation: "Respect d'une obligation lÃ©gale",
    other: d.mr_legal_basis_other || "Autre fondement juridique",
  };
  return labels[d.mr_legal_basis] || "â€”";
}

function mrRetentionText(d) {
  if (d.mr_retention === "cnil") return "Je respecterai la consigne de la CNIL";
  if (d.mr_retention === "other") return d.mr_retention_detail ? `DurÃ©e diffÃ©rente : ${d.mr_retention_detail}` : "DurÃ©e diffÃ©rente Ã  justifier";
  return "â€”";
}

function mrSensitiveText(d) {
  if (d.mr_sensitive === "no") return "Non";
  if (d.mr_sensitive === "yes") {
    return [
      "Oui",
      d.mr_sensitive_types ? `Lesquelles: ${d.mr_sensitive_types}` : null,
      d.mr_sensitive_justification ? `Justification: ${d.mr_sensitive_justification}` : null,
    ].filter(Boolean).join(" | ");
  }
  return "â€”";
}

function mrInformationText(d) {
  const patient = [
    d.mr_info_patient_level5 ? "Patients CLB : note aux niveaux < 5" : null,
    d.mr_info_patient_notice ? "Remise de la note d'information Ã  chaque patient" : null,
    d.mr_info_patient_other ? (d.mr_info_patient_other_txt || "Autres mÃ©thodes patients") : null,
  ].filter(Boolean);
  const nonPatient = [
    d.mr_info_nonpatient_oral ? "Information orale" : null,
    d.mr_info_nonpatient_written ? "Notice Ã©crite individuelle" : null,
    d.mr_info_nonpatient_other ? (d.mr_info_nonpatient_other_txt || "Autres mÃ©thodes hors patients") : null,
  ].filter(Boolean);
  const parts = [];
  if (patient.length) parts.push(`Patients: ${patient.join(", ")}`);
  if (nonPatient.length) parts.push(`Hors patients: ${nonPatient.join(", ")}`);
  return parts.join(" | ") || "â€”";
}

function mrEthicsNeedText(d) {
  const out = [
    d.mr_ethics_certificate ? "Courrier Certification d'instruction RGPD FR/EN" : null,
    d.mr_ethics_opinion ? `Avis Ã©thique CMT${d.mr_ethics_topic ? ` : ${d.mr_ethics_topic}` : ""}` : null,
  ].filter(Boolean);
  return out.join(" | ") || "â€”";
}

function mrContractText(d) {
  const out = [
    d.mr_contract_drafted ? `Contrat rÃ©digÃ© : ${d.mr_contract_drafted}` : null,
    d.mr_contract_validated_by ? `Validation juridique CLB : ${d.mr_contract_validated_by}` : null,
  ].filter(Boolean);
  return out.join(" | ") || "â€”";
}

function mrDataSelectedText(d, other) {
  return fixText(mrDataRegulatoryText(d));
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
  doc.setTextColor(20, 24, 32);
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
  doc.setTextColor(20, 24, 32);
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
  doc.setTextColor(20, 24, 32);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text(title, 17, y + 5.5);
  return y + 11;
}

function drawFieldBox(doc, y, label, value = "", height = 12) {
  y = ensurePageSpace(doc, y, height + 8);
  doc.setTextColor(20, 24, 32);
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
  doc.setTextColor(20, 24, 32);
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
  doc.setTextColor(20, 24, 32);
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
  doc.setTextColor(20, 24, 32);
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
  const doc = await buildMr004PdfDoc(d);
  doc.save(`MR004_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`);
}

async function buildMr004PdfDoc(d = getFormData()) {
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
    ["Type de tumeur / pathologie", mrTumorSelectedText(d)],
  ]);
  y = drawFieldBox(doc, y, "Période concernée", safe(d.q6_period || derivePeriod(d)), 10);

  y = drawSectionHeader(doc, y, "II. Description RGPD et sécurité");
  y = drawFieldBox(doc, y, "Catégorisation du projet (cas 1 à 6 ou autre)", mrCaseLabelText(d), 12);
  y = drawFieldBox(doc, y, "Questions conditionnelles selon le cas choisi", mrCaseDetailsText(d), 20);
  y = drawFieldBox(doc, y, "Centres / sites / partenaires concernés", [q6SitesSelectedText(d), partnerTypesText(d), safe(d.coord)].filter(v => v !== "—").join(" ; ") || "—", 12);
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
  return doc;
}

async function generateCmtPdf(d = getFormData()) {
  const doc = await buildCmtPdfDoc(d);
  doc.save(`CMT_${slugify(d.porteur || "porteur")}_${slugify(d.titre || "projet")}.pdf`);
}

async function buildCmtPdfDoc(d = getFormData()) {
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
    d.cmt_matching ? (safe(d.cmt_matching_detail) !== "â€”" ? `Oui — ${safe(d.cmt_matching_detail)}` : "Oui") : "Non",
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
  y = drawFieldBox(doc, y, "Plateformes / expertises", [cmtPlatformsSelectedText(d), safe(d.cmt_platform_details)].filter(v => v !== "â€”").join(" | ") || "â€”", 16);

  y = drawSectionHeader(doc, y, "Suivi et validation");
  y = drawFieldBox(doc, y, "Commentaires / réserves / suivi", safe(d.trf_com), 22);

  drawSignatureArea(doc, y, "Pour le Comité Médico-Technique");
  return doc;
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
