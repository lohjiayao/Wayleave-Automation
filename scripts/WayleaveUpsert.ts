interface ScriptResult {
  action: string;
  matched: boolean;
  engineerNames: string[];
  engineerEmails: string[];
  documentLink: string;
  documentName: string;
  outProjectRef: string;
  outProjectTitle: string;
  outWayleaveType: string;
  outDistrict: string;
  outAgency: string;
  outJkrRequired: string;
  outApprovalRef: string;
  note: string;
}

function main(
  workbook: ExcelScript.Workbook,
  submissionType: string,
  projectRef: string,
  projectTitle: string,
  wayleaveType: string,
  district: string,
  agency: string,
  jkrRequired: string,
  applicationLink: string,
  approvalRef: string,
  approvalLink: string,
  notifyEngineers: string,
  submittedOn: string
): ScriptResult {

  // ============ helpers ============
  const clean = (v: string): string =>
    (v === undefined || v === null) ? "" : ("" + v).trim();

  // "Project Ref / Submission Ref" -> "projectrefsubmissionref"
  const norm = (v: string): string =>
    clean(v).toLowerCase().replace(/[^a-z0-9]/g, "");

  // Microsoft Forms file-upload answers arrive as JSON:
  //   [{"name":"file.jpg","link":"https://..."}]
  // Plain text answers pass straight through.
  function extractLink(raw: string): string {
    const s = clean(raw);
    if (s === "") return "";
    const marker = '"link":"';
    const i = s.indexOf(marker);
    if (i < 0) return s;
    const rest = s.substring(i + marker.length);
    const end = rest.indexOf('"');
    return end >= 0 ? rest.substring(0, end) : rest;
  }

  function extractName(raw: string): string {
    const s = clean(raw);
    if (s === "") return "";
    const marker = '"name":"';
    const i = s.indexOf(marker);
    if (i < 0) return "";
    const rest = s.substring(i + marker.length);
    const end = rest.indexOf('"');
    return end >= 0 ? rest.substring(0, end) : rest;
  }

  // Accepts ["a","b"] / a;b / a, b / a
  function parseNames(raw: string): string[] {
    const s = clean(raw);
    if (s === "") return [];
    let parts: string[] = [];
    if (s.charAt(0) === "[") {
      parts = s.replace(/[\[\]"]/g, "").split(",");
    } else {
      parts = s.split(/[;,]/);
    }
    const out: string[] = [];
    for (let i = 0; i < parts.length; i++) {
      const n = parts[i].trim();
      if (n !== "") out.push(n);
    }
    return out;
  }

  submissionType = clean(submissionType);
  projectRef = clean(projectRef);
  projectTitle = clean(projectTitle);
  wayleaveType = clean(wayleaveType);
  district = clean(district);
  agency = clean(agency);
  jkrRequired = clean(jkrRequired);
  approvalRef = clean(approvalRef);
  submittedOn = clean(submittedOn);

  const rawAppUpload = applicationLink;
  const rawApprUpload = approvalLink;
  applicationLink = extractLink(rawAppUpload);
  approvalLink = extractLink(rawApprUpload);

  const selectedNames: string[] = parseNames(notifyEngineers);
  const isApproval: boolean = submissionType.toLowerCase().indexOf("approval") >= 0;

  // ============ find the right tables automatically ============
  const tables = workbook.getTables();

  let tracker: ExcelScript.Table = null;
  let logTable: ExcelScript.Table = null;
  let engTable: ExcelScript.Table = null;

  function headersOf(t: ExcelScript.Table): string[] {
    const raw = t.getHeaderRowRange().getValues()[0] as string[];
    const out: string[] = [];
    for (let i = 0; i < raw.length; i++) out.push(norm("" + raw[i]));
    return out;
  }

  function hasCol(hs: string[], needle: string): boolean {
    for (let i = 0; i < hs.length; i++) if (hs[i].indexOf(needle) >= 0) return true;
    return false;
  }

  for (let i = 0; i < tables.length; i++) {
    const hs = headersOf(tables[i]);
    if (engTable === null && hasCol(hs, "engineeremail")) { engTable = tables[i]; continue; }
    if (logTable === null && hasCol(hs, "timestamp")) { logTable = tables[i]; continue; }
    if (tracker === null && hasCol(hs, "projectref") && hasCol(hs, "approval")) { tracker = tables[i]; }
  }

  // fall back: any remaining table with a project ref column is the tracker
  if (tracker === null) {
    for (let i = 0; i < tables.length; i++) {
      if (tables[i] === logTable || tables[i] === engTable) continue;
      if (hasCol(headersOf(tables[i]), "projectref")) { tracker = tables[i]; break; }
    }
  }

  if (tracker === null) {
    return {
      action: "ERROR", matched: false, engineerNames: [], engineerEmails: [],
      documentLink: "", documentName: "",
      outProjectRef: "", outProjectTitle: "", outWayleaveType: "",
      outDistrict: "", outAgency: "", outJkrRequired: "", outApprovalRef: "",
      note: "No tracker table found - needs a column containing 'Project Ref'."
    };
  }

  // ============ column lookup on the tracker ============
  const tHeaders = headersOf(tracker);

  // find a column whose normalised name contains ALL the given fragments
  function findCol(fragments: string[]): number {
    for (let i = 0; i < tHeaders.length; i++) {
      let ok = true;
      for (let f = 0; f < fragments.length; f++) {
        if (tHeaders[i].indexOf(fragments[f]) < 0) { ok = false; break; }
      }
      if (ok) return i;
    }
    return -1;
  }

  const cRef        = findCol(["projectref"]);
  const cTitle      = findCol(["projecttitle"]);
  const cType       = findCol(["wayleavetype"]);
  const cDistrict   = findCol(["district"]);
  const cAgency     = findCol(["agency"]);
  const cJkr        = findCol(["jkrrequired"]);
  const cSubDate    = findCol(["submissiondate"]);
  const cAppLink    = findCol(["documentlink", "application"]);
  const cAppvLink   = findCol(["documentlink", "approval"]);
  const cApprRef    = findCol(["approvalref"]);
  const cApprDate   = findCol(["approvaldate"]);
  const cNotify     = findCol(["notifyengineer"]);
  const cNotified   = findCol(["engineernotified"]);
  const cSentDate   = findCol(["notificationsent"]);
  const cStatus     = findCol(["notificationstatus"]);

  const body = tracker.getRangeBetweenHeaderAndTotal();
  const rows = body.getValues();

  let matchRow = -1;
  let firstBlank = -1;
  const key = projectRef.toLowerCase();

  for (let r = 0; r < rows.length; r++) {
    const cell = clean("" + rows[r][cRef]);
    if (cell === "") { if (firstBlank < 0) firstBlank = r; continue; }
    if (key !== "" && cell.toLowerCase() === key) { matchRow = r; break; }
  }

  // ============ build the values to write ============
  const writes: { col: number, value: string }[] = [];
  function put(c: number, v: string) {
    if (c >= 0 && clean(v) !== "") writes.push({ col: c, value: clean(v) });
  }

  if (isApproval) {
    put(cApprRef, approvalRef);
    put(cApprDate, submittedOn);
    put(cAppvLink, approvalLink);
  } else {
    put(cRef, projectRef);
    put(cTitle, projectTitle);
    put(cType, wayleaveType);
    put(cDistrict, district);
    put(cAgency, agency);
    put(cJkr, jkrRequired);
    put(cAppLink, applicationLink);
    put(cSubDate, submittedOn);
  }

  if (selectedNames.length > 0) {
    put(cNotify, selectedNames.join("; "));
    put(cNotified, "Yes");
    put(cSentDate, submittedOn);
    put(cStatus, "Sent");
  } else if (!isApproval) {
    put(cNotified, "No");
    put(cStatus, "Pending");
  }

  // ============ write ============
  let action = "";
  let matched = false;
  let note = "";
  let writtenRow = -1;

  if (matchRow >= 0) {
    for (let i = 0; i < writes.length; i++) {
      body.getCell(matchRow, writes[i].col).setValue(writes[i].value);
    }
    action = isApproval ? "Approval Updated" : "Updated";
    matched = true;
    writtenRow = matchRow;
  } else if (isApproval) {
    action = "Approval - No Matching Project";
    note = "No tracker row found with Project Ref '" + projectRef + "'.";
  } else {
    let target = firstBlank;
    if (target < 0) {
      const blank: string[] = [];
      for (let i = 0; i < tHeaders.length; i++) blank.push("");
      tracker.addRow(-1, blank);
      target = rows.length;
    }
    const freshBody = tracker.getRangeBetweenHeaderAndTotal();
    for (let i = 0; i < writes.length; i++) {
      freshBody.getCell(target, writes[i].col).setValue(writes[i].value);
    }
    action = "Added";
    writtenRow = target;
  }

  // read the finished row back so the email always has full project context,
  // even on an Approval submission where most form fields are blank
  let ctx: string[] = [];
  if (writtenRow >= 0) {
    ctx = tracker.getRangeBetweenHeaderAndTotal().getValues()[writtenRow] as string[];
  }
  function ctxAt(c: number): string {
    if (c < 0 || ctx.length === 0) return "";
    return clean("" + ctx[c]);
  }

  // ============ engineer email lookup ============
  const engineerNames: string[] = [];
  const engineerEmails: string[] = [];

  if (selectedNames.length > 0) {
    if (engTable === null) {
      note = note + " No engineer lookup table found (needs an 'Engineer Email' column).";
    } else {
      const eh = headersOf(engTable);
      let nameCol = -1, mailCol = -1;
      for (let i = 0; i < eh.length; i++) {
        if (mailCol < 0 && eh[i].indexOf("email") >= 0) mailCol = i;
        else if (nameCol < 0 && eh[i].indexOf("name") >= 0) nameCol = i;
      }
      if (nameCol < 0) nameCol = 0;
      if (mailCol < 0) mailCol = 1;

      const eRows = engTable.getRangeBetweenHeaderAndTotal().getValues();
      for (let s = 0; s < selectedNames.length; s++) {
        const want = selectedNames[s].toLowerCase();
        for (let r = 0; r < eRows.length; r++) {
          const n = clean("" + eRows[r][nameCol]).toLowerCase();
          const e = clean("" + eRows[r][mailCol]);
          if (n !== "" && n === want && e !== "") {
            engineerNames.push(clean("" + eRows[r][nameCol]));
            engineerEmails.push(e);
            break;
          }
        }
      }
    }
  }

  // ============ audit log ============
  if (logTable !== null) {
    const lh = headersOf(logTable);
    const logVals: { [key: string]: string } = {
      "timestamp": submittedOn,
      "projectref": projectRef,
      "projecttitle": projectTitle,
      "wayleavetype": wayleaveType,
      "district": district,
      "agency": agency,
      "jkrrequired": jkrRequired,
      "documentlink": isApproval ? approvalLink : applicationLink,
      "notifyengineer": selectedNames.join("; "),
      "notifiedengineer": selectedNames.join("; "),
      "action": (isApproval ? "Approval - " : "Application - ") + action
    };

    const rowOut: string[] = [];
    for (let i = 0; i < lh.length; i++) {
      let v = "";
      for (const k of Object.keys(logVals)) {
        if (lh[i].indexOf(k) >= 0) { v = logVals[k]; break; }
      }
      rowOut.push(v);
    }

    const lBody = logTable.getRangeBetweenHeaderAndTotal();
    const lRows = lBody.getValues();
    let lBlank = -1;
    for (let r = 0; r < lRows.length; r++) {
      if (clean("" + lRows[r][0]) === "") { lBlank = r; break; }
    }
    if (lBlank >= 0) {
      for (let i = 0; i < rowOut.length; i++) {
        if (rowOut[i] !== "") lBody.getCell(lBlank, i).setValue(rowOut[i]);
      }
    } else {
      logTable.addRow(-1, rowOut);
    }
  } else {
    note = note + " No submission log table found (needs a 'Timestamp' column).";
  }

  return {
    action: action,
    matched: matched,
    engineerNames: engineerNames,
    engineerEmails: engineerEmails,
    documentLink: isApproval ? approvalLink : applicationLink,
    documentName: extractName(isApproval ? rawApprUpload : rawAppUpload),
    outProjectRef: projectRef,
    outProjectTitle: ctxAt(cTitle),
    outWayleaveType: ctxAt(cType),
    outDistrict: ctxAt(cDistrict),
    outAgency: ctxAt(cAgency),
    outJkrRequired: ctxAt(cJkr),
    outApprovalRef: ctxAt(cApprRef),
    note: clean(note)
  };
}
