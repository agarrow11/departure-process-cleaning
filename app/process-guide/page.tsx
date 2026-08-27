import type { Metadata } from "next"
import Link from "next/link"
import {
  ArrowLeft,
  ArrowRight,
  Check,
  FileSpreadsheet,
  KeyRound,
  LockKeyhole,
  ShieldAlert,
} from "lucide-react"
import { PrintButton } from "@/components/process-guide/print-button"

export const metadata: Metadata = {
  title: "Process Guide | Departure Data Pipeline",
  description: "Operator and technical guide to the departure-data cleaning, merge, anonymization, and export process.",
}

const sections = [
  ["overview", "1. Purpose & controls"],
  ["inputs", "2. Required inputs"],
  ["operator", "3. Operator workflow"],
  ["flow", "4. Process flow"],
  ["rules", "5. Field-level rules"],
  ["merge", "6. Merge logic"],
  ["example", "7. Worked example"],
  ["outputs", "8. Final outputs"],
  ["checklist", "9. Checklists"],
  ["troubleshooting", "10. Troubleshooting"],
  ["technical", "11. Technical appendix"],
] as const

const inputs = [
  ["Old Survey", ".xlsx", "Departure survey export using row 2 as the header; row 3 is metadata and is discarded.", "Ecode"],
  ["New Survey", ".xlsx", "Newer departure survey export with the same row-2 header convention.", "Ecode"],
  ["Exit Interview", ".xlsx", "Interview export using row 2 as the header; row 3 is metadata and is discarded.", "Cleaned Ecode"],
  ["Beyond Bain Extract", ".xlsx", "Post-departure enrichment extract. Only eight approved fields are retained.", "Ecode"],
  ["Department Hierarchy", ".xlsx", "Department mapping table. The first worksheet whose name contains “dept” is used.", "Department Code + Department Name"],
  ["Geographic Hierarchy", ".xlsx", "Geography mapping table. The first worksheet whose name contains “geo” is used.", "Country + Office + Office Location"],
  ["PD Grade Mapping", ".csv", "Performance-development grade lookup.", "Job Profile"],
] as const

const ruleRows = [
  ["All Excel inputs", "Date-formatted cells", "Detected by spreadsheet cell type, rounded safely, and rendered as MM/DD/YYYY using UTC components.", "Same field name", "Non-date values are unchanged."],
  ["Survey + Exit Interview", "Header row", "Row 2 becomes the header; row 3 metadata is skipped; data begins at row 4.", "Fixed schema per file", "A missing/empty header row is a hard error."],
  ["Survey + Exit Interview", "Duplicate headers", "Repeated names are de-collided in first-seen order with .1, .2, and so on.", "Unique column names", "Prevents one duplicate field from overwriting another."],
  ["All row-based inputs", "Trailing blank columns", "Blank cells at the right edge of the header are removed; interior blank positions remain.", "Stable ordered schema", "Interior columns are never shifted left."],
  ["Survey", "Ecode", "Trim whitespace and uppercase text before mapping and joining.", "Ecode", "Blank stays blank and does not join to another blank."],
  ["Exit Interview", "Cleaned Ecode", "Trim, uppercase, and rename in place.", "Ecode", "The interviewer-entered raw Ecode question is deleted entirely."],
  ["Department mapping", "Department Code + Department Name", "Composite key lookup; mapped fields overwrite any same-named source fields.", "Business Unit, Division, Market, Department Group, Department Name", "Unmatched values become blank and are counted/audited."],
  ["Geography mapping", "Country + Office + Office Location", "Composite key lookup; mapped fields overwrite same-named source fields.", "Region, Country, Office, Office Cluster, Office Status", "Unmatched values become blank and are counted/audited."],
  ["PD Grade mapping", "Job Profile", "Single-key lookup; matching is trimmed and case-insensitive.", "Mapped Employee Level", "Unmatched values become blank and are counted/audited."],
  ["Beyond Bain", "Function (ZID4_117)", "If several functions are separated by semicolons, retain only the first trimmed function.", "BB_Function (ZID4_117) + BB_Mapped Role Function", "Blank/null remains blank."],
  ["Beyond Bain", "Ecode", "Left-enrich the merged population; first Beyond Bain record per standardized Ecode wins.", "Eight BB_ fields only", "No match produces blank BB_ fields."],
  ["Final assembly", "TRUE / FALSE strings", "Normalize SheetJS uppercase boolean strings to Python-style True / False.", "Same field", "Other text is unchanged."],
  ["Final assembly", "Ecode", "Every unique non-blank Ecode receives one random five-digit code (10000–99999); repeated rows reuse that code. A hard reconciliation compares full/no-PII rows and the mapping before files are emitted.", "Ecode + separate mapping file", "Blank stays blank and is excluded from unique counts. The run stops unless output code sets, unique originals, and mapping rows agree exactly."],
  ["PII-redacted exports", "Eight name/email fields", "Remove the exact eight approved PII columns after all joins and mappings are complete.", "Eight headers and their values are absent from no-PII files", "Full XLSX retains these fields; Ecode is anonymized in every output."],
  ["PII-redacted exports", "Seven fixed-schema contact fields", "Retain each header in its original position but clear every value.", "Headers remain in the 175-column no-PII schema; contents are blank", "A hard check stops export if a header is absent, reordered, removed, or populated."],
] as const

const piiColumns = [
  "Recipient Last Name",
  "Recipient First Name",
  "Recipient Email",
  "*Email address (personal):",
  "Business / school email address:",
  "RecipientEmail",
  "Legal name",
  "[For interviewer] Please provide information on the departing employee - Employee name:",
]

const retainedBlankPIIColumns = [
  "LinkedIn URL:",
  "Contact Information - Street Address:",
  "Contact Information - City:",
  "Contact Information - State",
  "Contact Information - Zip Code:",
  "Contact Information - Country:",
  "Email - Home",
]

function SectionHeading({ kicker, title, intro }: { kicker: string; title: string; intro?: string }) {
  return (
    <header className="guide-section-heading">
      <p>{kicker}</p>
      <h2>{title}</h2>
      {intro && <div>{intro}</div>}
    </header>
  )
}

export default function ProcessGuidePage() {
  return (
    <main className="process-guide">
      <header className="guide-topbar print-hide">
        <div className="guide-topbar-inner">
          <Link className="guide-back" href="/">
            <ArrowLeft aria-hidden="true" size={16} /> Back to pipeline
          </Link>
          <PrintButton />
        </div>
      </header>

      <div className="guide-hero">
        <div className="guide-hero-inner">
          <div className="guide-rule" />
          <p className="guide-eyebrow">Bain HR Analytics · Operating manual</p>
          <h1>Departure Data Pipeline<br />Process Guide</h1>
          <p className="guide-deck">
            A transparent, step-by-step account of what enters the process, how every material field is cleaned and linked,
            which controls run, and what leaves the process.
          </p>
          <dl className="guide-meta">
            <div><dt>Scope</dt><dd>Current production behavior</dd></div>
            <div><dt>Audience</dt><dd>Operators, reviewers & technical maintainers</dd></div>
            <div><dt>Reviewed</dt><dd>25 August 2026 · commit d199b9a</dd></div>
          </dl>
        </div>
      </div>

      <div className="guide-shell">
        <aside className="guide-toc print-hide" aria-label="On this page">
          <p>On this page</p>
          <nav>
            {sections.map(([id, label]) => <a key={id} href={`#${id}`}>{label}</a>)}
          </nav>
        </aside>

        <article className="guide-article">
          <section id="overview" className="guide-section">
            <SectionHeading kicker="01 · Orientation" title="Purpose, ownership and non-negotiable controls" />
            <div className="guide-lead-grid">
              <div className="guide-big-statement">
                <strong>Seven inputs become four controlled outputs.</strong>
                <span>The process standardizes identifiers, enriches HR fields, reconciles two departure populations, limits Beyond Bain data, anonymizes Ecode, and produces a PII-redacted distribution copy.</span>
              </div>
              <div className="guide-security-callout">
                <ShieldAlert aria-hidden="true" size={22} />
                <div><strong>The mapping file is the re-identification key.</strong><p>Store it separately from distributed data, restrict access, and never send it with the no-PII files unless re-identification is explicitly authorized.</p></div>
              </div>
            </div>
            <div className="guide-three-registers">
              <div><span>Operator owns</span><strong>Correct files, review of warnings, secure distribution</strong></div>
              <div><span>Pipeline owns</span><strong>Repeatable rules, joins, audits, anonymization, exports</strong></div>
              <div><span>Reviewer owns</span><strong>Reasonableness, unmatched records, retention decisions</strong></div>
            </div>
            <div className="guide-note"><strong>Important:</strong> this is a deterministic transformation tool, not a source-system correction tool. It does not repair the underlying workbooks or infer missing mappings.</div>
          </section>

          <section id="inputs" className="guide-section">
            <SectionHeading kicker="02 · Before the run" title="The seven required input files" intro="Use fresh exports from the agreed sources. Do not rename or delete required key columns inside the files." />
            <div className="guide-table-wrap">
              <table className="guide-table">
                <thead><tr><th>Input</th><th>Format</th><th>Purpose / structure</th><th>Link or lookup key</th></tr></thead>
                <tbody>{inputs.map((row) => <tr key={row[0]}>{row.map((cell, index) => <td key={`${row[0]}-${index}`}>{cell}</td>)}</tr>)}</tbody>
              </table>
            </div>
            <h3>File preparation checks</h3>
            <ul className="guide-check-list">
              <li><Check size={15} />Confirm all seven files open and are not password-protected.</li>
              <li><Check size={15} />Confirm the four population files are for the same reporting cycle.</li>
              <li><Check size={15} />Confirm Survey and Exit Interview exports retain their original row-2 headers and row-3 metadata rows.</li>
              <li><Check size={15} />Confirm Ecode columns have not been converted to formulas or scientific notation.</li>
              <li><Check size={15} />Confirm mapping files contain the expected key columns listed above.</li>
              <li><Check size={15} />Treat every source file as confidential HR data.</li>
            </ul>
          </section>

          <section id="operator" className="guide-section">
            <SectionHeading kicker="03 · Runbook" title="How an operator runs the process" />
            <ol className="guide-steps">
              <li><span>01</span><div><strong>Open the pipeline dashboard.</strong><p>Use the “How to use” panel or this guide before selecting files.</p></div></li>
              <li><span>02</span><div><strong>Upload all seven files into their named slots.</strong><p>The Run Pipeline button remains unavailable until every required file is present.</p></div></li>
              <li><span>03</span><div><strong>Check each displayed filename.</strong><p>If a file is in the wrong slot, remove it and upload the correct source before running.</p></div></li>
              <li><span>04</span><div><strong>Select Run Pipeline once.</strong><p>Keep the browser open while mappings load, surveys stack, records merge, Beyond Bain enriches, and outputs render.</p></div></li>
              <li><span>05</span><div><strong>Read the result summary before downloading.</strong><p>Review row counts, source composition, unmatched lookups, duplicate-key audits, and the Ecode anonymization integrity check. It confirms repeated rows reuse one code and all cleaned outputs reconcile to the mapping.</p></div></li>
              <li><span>06</span><div><strong>Download only the files needed for the recipient.</strong><p>Use no-PII files for ordinary sharing. Use the full XLSX only where names/emails are explicitly required. Restrict the Ecode mapping file.</p></div></li>
              <li><span>07</span><div><strong>Record the run and handoff.</strong><p>Log source dates, operator, result counts, warnings accepted, recipient, and secure storage location.</p></div></li>
            </ol>
            <div className="guide-status-grid">
              <div className="status-success"><strong>Success</strong><p>The run completed and four downloads are available. This does not mean all mappings matched.</p></div>
              <div className="status-warning"><strong>Warning</strong><p>The run completed but requires review—for example, unmapped keys, duplicate Ecodes, or missing enrichment.</p></div>
              <div className="status-error"><strong>Error</strong><p>No valid output should be distributed. Correct the named file/schema issue and rerun all seven files together.</p></div>
            </div>
          </section>

          <section id="flow" className="guide-section">
            <SectionHeading kicker="04 · Lineage" title="Source-to-output process flow" />
            <div className="guide-flow" role="img" aria-label="Seven source files pass through ingestion, normalization, mappings, a survey and exit interview merge, Beyond Bain enrichment, anonymization and PII scrubbing, producing four output files.">
              <div className="flow-sources">
                <strong>7 inputs</strong>
                <span>Old Survey</span><span>New Survey</span><span>Exit Interview</span><span>Beyond Bain</span><span>Department map</span><span>Geography map</span><span>PD Grade map</span>
              </div>
              <ArrowRight className="flow-arrow" aria-hidden="true" />
              <div className="flow-stack">
                <span>1 · Ingest & schema</span><span>2 · Normalize & map</span><span>3 · Survey + EI full outer join</span><span>4 · Beyond Bain left enrichment</span><span>5 · Anonymize & scrub</span>
              </div>
              <ArrowRight className="flow-arrow" aria-hidden="true" />
              <div className="flow-outputs"><strong>4 outputs</strong><span>Full XLSX</span><span>No-PII XLSX</span><span>No-PII CSV</span><span>Ecode map CSV</span></div>
            </div>
            <p className="guide-caption">The sequence is load-bearing: mapping and joins use the original standardized Ecode; anonymization happens only after linkage is complete.</p>
          </section>

          <section id="rules" className="guide-section">
            <SectionHeading kicker="05 · Rule catalog" title="What is cleaned, mapped, removed or created" intro="This table is the practical field-level contract. Rules are applied identically on every run." />
            <div className="guide-table-wrap guide-table-wide">
              <table className="guide-table guide-rule-table">
                <thead><tr><th>Stage/source</th><th>Input field</th><th>Rule</th><th>Output effect</th><th>Missing / unmatched behavior</th></tr></thead>
                <tbody>{ruleRows.map((row) => <tr key={`${row[0]}-${row[1]}`}>{row.map((cell, index) => <td key={`${row[0]}-${row[1]}-${index}`}>{cell}</td>)}</tr>)}</tbody>
              </table>
            </div>
            <h3>The eight columns removed from no-PII files</h3>
            <ol className="guide-column-list">{piiColumns.map((column, index) => <li key={column}><span>{String(index + 1).padStart(2, "0")}</span><code>{column}</code></li>)}</ol>
            <h3>The seven headers retained with blank contents in no-PII files</h3>
            <ol className="guide-column-list">{retainedBlankPIIColumns.map((column, index) => <li key={column}><span>{String(index + 1).padStart(2, "0")}</span><code>{column}</code></li>)}</ol>
            <p className="guide-caption">The no-PII schema remains fixed at 175 columns: the prior eight columns stay removed, while these seven remain in their original positions with empty values. The interviewer-entered Employee Ecode field is different: it is deleted upstream from every output, before the final schema is assembled.</p>
          </section>

          <section id="merge" className="guide-section">
            <SectionHeading kicker="06 · Population logic" title="How records are stacked, joined and enriched" />
            <div className="guide-logic-grid">
              <div><span>A</span><h3>Stack surveys</h3><p>Old Survey rows are followed by New Survey rows. Columns follow first-seen order: all Old Survey columns first, then New-only columns.</p></div>
              <div><span>B</span><h3>Apply lookups</h3><p>Department, Geography and PD Grade mappings are applied to the combined survey before the population join. Mapping output overwrites same-named source fields.</p></div>
              <div><span>C</span><h3>Full outer join</h3><p>Survey and Exit Interview join on standardized Ecode. Shared columns coalesce with Survey precedence; EI fills only when Survey is blank.</p></div>
              <div><span>D</span><h3>Preserve every side</h3><p>Survey-only and EI-only records remain. Blank Ecode rows never match each other. Duplicate keys create a cross product and are surfaced in audit results.</p></div>
              <div><span>E</span><h3>Left-enrich Beyond Bain</h3><p>Every merged population row remains. If several Beyond Bain records share an Ecode, the first record is used; no match leaves eight BB_ fields blank.</p></div>
              <div><span>F</span><h3>Fix the output schema</h3><p>Population fields remain in merged order, followed by exactly eight approved BB_ columns. Internal lineage fields never enter the download.</p></div>
            </div>
            <div className="guide-note"><strong>Cardinality risk:</strong> duplicate Ecodes on both Survey and Exit Interview sides multiply rows. For example, two survey rows × three EI rows for one Ecode produce six joined rows. The audit panel reports duplicate-key risk; it does not silently collapse records.</div>
          </section>

          <section id="example" className="guide-section">
            <SectionHeading kicker="07 · Synthetic example" title="Follow one employee through the curtain" intro="All values below are invented for training." />
            <div className="guide-example">
              <div><span>Source</span><strong>Ecode “ ab123 ”</strong><p>Survey has department D14 / Customer; EI has Cleaned Ecode “AB123”; Beyond Bain has “Operations;Strategy”.</p></div>
              <ArrowRight aria-hidden="true" />
              <div><span>Standardize</span><strong>AB123</strong><p>Whitespace is trimmed, text is uppercased, and EI Cleaned Ecode is renamed Ecode.</p></div>
              <ArrowRight aria-hidden="true" />
              <div><span>Map & merge</span><strong>One linked record</strong><p>Department/geography/grade fields are added. Survey values lead shared fields; EI fills Survey blanks.</p></div>
              <ArrowRight aria-hidden="true" />
              <div><span>Enrich & scrub</span><strong>Operations · 48217</strong><p>Only the first Function is kept. AB123 becomes a unique five-digit code; names/emails disappear from no-PII files.</p></div>
            </div>
            <div className="guide-example-note"><KeyRound size={18} /><p>The mapping CSV holds <code>48217 → AB123</code>. “48217” is illustrative only; actual codes are randomly generated on each run and can change between reruns.</p></div>
          </section>

          <section id="outputs" className="guide-section">
            <SectionHeading kicker="08 · Deliverables" title="The four files produced after each successful run" />
            <div className="guide-output-list">
              <div><FileSpreadsheet /><span>01</span><h3>Full output .xlsx</h3><p>Current full schema (183 columns in verified test data). Includes the eight name/email fields but uses anonymized five-digit Ecode. Unresolved mapped cells are highlighted yellow.</p><strong>Restricted use</strong></div>
              <div><FileSpreadsheet /><span>02</span><h3>_no_pii.xlsx</h3><p>Same rows and processing, with eight listed columns removed and seven additional PII columns retained but blank. The fixed schema remains 175 columns and keeps yellow unresolved-lookup highlighting.</p><strong>Preferred sharing copy</strong></div>
              <div><FileSpreadsheet /><span>03</span><h3>_no_pii.csv</h3><p>CSV twin of the no-PII workbook: the same fixed 175-column structure, including the seven retained headers with blank values. CSV cannot carry cell highlighting.</p><strong>Preferred data-feed copy</strong></div>
              <div><KeyRound /><span>04</span><h3>_ecode_mapping.csv</h3><p>Two columns: Anonymized Code and Original Ecode. This is the only output designed to reverse anonymization.</p><strong>Highest access restriction</strong></div>
            </div>
            <div className="guide-security-callout"><LockKeyhole size={22} /><div><strong>Do not co-locate distributed data and its key.</strong><p>The no-PII file plus the mapping file is re-identifiable. Apply the same controls as the original HR source data.</p></div></div>
          </section>

          <section id="checklist" className="guide-section">
            <SectionHeading kicker="09 · Controls" title="Operator and handoff checklist" />
            <div className="guide-check-columns">
              <div><h3>Before running</h3><label><input type="checkbox" /> Correct seven files and reporting cycle</label><label><input type="checkbox" /> Required sheets/keys retained</label><label><input type="checkbox" /> Files open and are not password-protected</label><label><input type="checkbox" /> Authorized workspace and operator</label></div>
              <div><h3>After running</h3><label><input type="checkbox" /> Row and source counts are plausible</label><label><input type="checkbox" /> Unmatched mappings reviewed</label><label><input type="checkbox" /> Duplicate-key audit reviewed</label><label><input type="checkbox" /> Four downloads available</label></div>
              <div><h3>Before sharing</h3><label><input type="checkbox" /> Recipient needs selected format</label><label><input type="checkbox" /> No-PII copy used by default</label><label><input type="checkbox" /> Mapping key sent separately, if authorized</label><label><input type="checkbox" /> Retention/deletion date recorded</label></div>
            </div>
          </section>

          <section id="troubleshooting" className="guide-section">
            <SectionHeading kicker="10 · Exceptions" title="Troubleshooting and interpretation" />
            <div className="guide-faq">
              <details open><summary>Run button is disabled</summary><p>At least one required file slot is empty. Confirm all four inputs and all three mappings are present.</p></details>
              <details><summary>Header row is empty or columns look shifted</summary><p>Survey and Exit Interview readers expect headers on Excel row 2 and data after the row-3 metadata line. Re-export from the source; do not manually move rows without change approval.</p></details>
              <details><summary>Many mapped fields are blank</summary><p>Review key fields for spelling, whitespace and source-period mismatches. The process intentionally leaves unmatched results blank rather than inventing a mapping.</p></details>
              <details><summary>Population row count is higher than expected</summary><p>Inspect duplicate Ecode audits. Duplicate keys on both sides create a many-to-many cross product; the tool preserves rather than collapses these records.</p></details>
              <details><summary>Two reruns produce different anonymized codes</summary><p>Expected. Five-digit codes are randomly allocated per run. Never compare codes across runs without each run’s own mapping file.</p></details>
              <details><summary>The CSV has no yellow cells</summary><p>Expected. CSV stores values only. Use the XLSX output to see unresolved mapped values highlighted in yellow.</p></details>
            </div>
          </section>

          <section id="technical" className="guide-section">
            <SectionHeading kicker="11 · Behind the curtain" title="Technical appendix and change control" />
            <div className="guide-table-wrap">
              <table className="guide-table">
                <thead><tr><th>Implementation area</th><th>Role in the process</th></tr></thead>
                <tbody>
                  <tr><td><code>app/page.tsx</code></td><td>Seven-file operator interface, execution state, audit/warning presentation, and four download controls.</td></tr>
                  <tr><td><code>app/api/pipeline/route.ts</code></td><td>Validates multipart uploads, calls the pipeline, creates full/no-PII exports, and returns base64 payloads.</td></tr>
                  <tr><td><code>readWorkbook / sheetToMatrix</code></td><td>Spreadsheet parsing, typed date detection, MM/DD/YYYY normalization, and row preservation.</td></tr>
                  <tr><td><code>loadSurvey / loadExitInterview</code></td><td>Fixed header conventions, Ecode standardization, raw interviewer-Ecode deletion, lineage fields.</td></tr>
                  <tr><td><code>applyMappings</code></td><td>Department, geography, and PD Grade lookups plus unmatched audit references.</td></tr>
                  <tr><td><code>fullOuterJoin</code></td><td>Survey/EI reconciliation, source precedence, one-sided retention, duplicate-key cross products.</td></tr>
                  <tr><td><code>applyBeyondBainLogic</code></td><td>First-function cleaning and mapped-role derivation before enrichment.</td></tr>
                  <tr><td><code>BB_ALLOWLIST</code></td><td>Defines the only eight Beyond Bain fields allowed into the output.</td></tr>
                  <tr><td><code>anonymizeEcodes</code></td><td>Random unique five-digit code assignment and reversible mapping-table creation.</td></tr>
                  <tr><td><code>PII_COLUMNS / PII_COLUMNS_TO_CLEAR / stripPIIColumns</code></td><td>Two-part no-PII contract: remove eight approved columns, retain seven fixed-schema headers with blank contents.</td></tr>
                  <tr><td><code>validatePIIRedaction</code></td><td>Hard export gate confirming the 175-column order, required retained headers, and blank values in every retained PII field.</td></tr>
                  <tr><td><code>rowsToXLSX / rowsToCSV</code></td><td>Final serialization; XLSX also highlights unresolved mapping blanks.</td></tr>
                </tbody>
              </table>
            </div>
            <h3>Known constraints</h3>
            <ul className="guide-bullets">
              <li>Input structure is contractual: Survey/EI headers are fixed at row 2; mapping header rows are detected differently.</li>
              <li>Mapping tables use first-key-wins behavior when a key appears more than once.</li>
              <li>Beyond Bain enrichment uses the first record found for a repeated Ecode.</li>
              <li>Anonymized codes are run-specific, random and not stable identifiers across time.</li>
              <li>The full XLSX is not a no-PII artifact; it retains eight explicit name/email columns.</li>
              <li>Excel formulas are read as their cached values. Charts, images and workbook formatting are not part of the transformation.</li>
            </ul>
            <h3>Change-control rule</h3>
            <div className="guide-note">Whenever a source header, mapping key, join rule, allowlist, anonymization rule, PII list, output count or export name changes: update the pipeline and this guide in the same commit; rerun type/build checks and a real-file reconciliation; verify the no-PII column census; then update the “Reviewed” marker at the top of this page.</div>
          </section>

          <footer className="guide-footer">
            <div className="guide-rule" />
            <strong>Departure Data Pipeline · Process Guide</strong>
            <p>Current behavior as reviewed 25 August 2026. Synthetic examples only.</p>
          </footer>
        </article>
      </div>
    </main>
  )
}
