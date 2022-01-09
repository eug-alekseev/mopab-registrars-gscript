const SHEETS = ["Sheet1", "Sheet2"];  // названия листов, для которых запускать обработку

const REVIEWED_TEMPLATE_UID = "GDOC_TEMPLATE_UID";
const FINALIZED_TEMPLATE_UID = "GDOC_TEMPLATE_UID";

const COLUMNS_CONFIG = {
  DOCTOR_NAME: "C",
  CLINIC_NAME: "D",
  PATIENT_NAME: "E",
  PATIENT_UID: "F",
  BIRTH_DATE: "G",
  ADDRESS: "H",
  ANAMNESIS: "I",
  DATE_DELIVERED: "K",
  NUMBERS: "L",
  SLIDE_AMOUNT: "M",
  BLOCK_AMOUNT: "N",
  CLINICAL_DIAGNOSIS: "P",
  MOPAB_DIAGNOSIS: "Q",
  MOPAB_DOCTOR: "R",
  MOPAB_DATE: "S",
  DOC_LINK: "T",
  NSPC_DIAGNOSIS: "U",
  NSPC_DOCTOR: "V",
  NSPC_DATE: "W",
  NSPC_NUMBER: "X",
  READY_TO_SEND: "Y",
  SENT_AT: "Z",
  AUTOMATION_STATUS: "AA",  // колонка используется только скриптом, вручную не менять!
};

const DIR_CONFIG = {  // куда будут сохраняться заключения в разных статусах
  FINAL: ["NSPC_GDISK_DIR_UID", "MOPAB_GDISK_DIR_UID"],  // "окончательный"
  READY: ["MOPAB_GDISK_DIR_UID"],  // "готово"
  REVIEW: ["NSPC_GDISK_DIR_UID", "MOPAB_GDISK_DIR_UID"],  // "на пересмотр"
};
