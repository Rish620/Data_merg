import { ChangeEvent, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  Activity,
  BarChart3,
  Bell,
  Check,
  ChevronDown,
  CircleAlert,
  ClipboardCheck,
  Code2,
  Database,
  Download,
  FileSpreadsheet,
  FolderKanban,
  GitBranch,
  LayoutDashboard,
  Menu,
  Play,
  Plus,
  Search,
  Settings,
  ShieldCheck,
  SlidersHorizontal,
  Sparkles,
  UploadCloud,
  Users,
  X,
  Zap,
} from "lucide-react";

type Row = Record<string, string>;
type Field = {
  name: string;
  type: string;
  recommended: string;
  confidence: number;
  nulls: number;
  unique: number;
  examples: string[];
  accepted: boolean;
  required: boolean;
};
type Rule = {
  name: string;
  fields: string[];
  method: string;
  threshold: number;
  priority: number;
};
type LogItem = {
  time: string;
  action: string;
  status: "Running" | "Completed" | "Warning" | "Failed";
  details: string;
};
type Config = {
  project: {
    name: string;
    tenant: string;
    environment: string;
    entityType: string;
    customEntityType?: string;
    source: string;
  };
  fileName: string;
  fileSize: number;
  rows: Row[];
  fields: Field[];
  rules: Rule[];
  survivorship: Record<string, string>;
  validations: string[];
  completed: boolean[];
  generated: string;
  logs: LogItem[];
};

const steps = [
  "Project Setup",
  "Data Upload",
  "Data Profiling",
  "Entity Configuration",
  "Match Analysis",
  "Match Rules",
  "Survivorship",
  "Data Validation",
  "Configuration Review",
  "JSON Generation",
  "Deployment",
];
const nav = [
  [
    "WORKSPACE",
    [
      ["Dashboard", LayoutDashboard],
      ["Projects", FolderKanban],
    ],
  ],
  [
    "BUILD",
    [
      ["Data Upload", UploadCloud],
      ["Data Profiling", BarChart3],
      ["Entity Configuration", SlidersHorizontal],
      ["Match Analysis", GitBranch],
      ["Match Rules", GitBranch],
      ["Survivorship", Users],
      ["Data Validation", ShieldCheck],
    ],
  ],
  [
    "DELIVER",
    [
      ["Configuration Review", ClipboardCheck],
      ["JSON Generator", Code2],
      ["Deployment", Zap],
    ],
  ],
] as const;
const blank: Config = {
  project: {
    name: "",
    tenant: "",
    environment: "Development",
    entityType: "Individual",
    source: "",
  },
  fileName: "",
  fileSize: 0,
  rows: [],
  fields: [],
  rules: [],
  survivorship: {},
  validations: [],
  completed: Array(11).fill(false),
  generated: "",
  logs: [],
};

export default function App() {
  const [config, setConfig] = useState<Config>(blank);
  const [active, setActive] = useState("Dashboard");
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(0);
  const [projectOpen, setProjectOpen] = useState(true);
  const [error, setError] = useState("");
  const [pendingApply, setPendingApply] = useState<
    "Match Rules" | "Survivorship" | null
  >(null);
  const input = useRef<HTMLInputElement>(null);
  const log = (
    action: string,
    details: string,
    status: LogItem["status"] = "Completed",
  ) => {
    const time = new Date().toLocaleTimeString([], {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    });
    setConfig((c) => ({
      ...c,
      logs: [{ time, action, status, details }, ...c.logs],
    }));
  };
  const mark = (index: number) =>
    setConfig((c) => ({
      ...c,
      completed: c.completed.map((done, i) => (i === index ? true : done)),
    }));
  const parseFile = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setError("");
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json<Row>(sheet, { defval: "" });
      if (!raw.length || !Object.keys(raw[0]).length)
        throw new Error(
          "The uploaded file has no readable records or headers.",
        );
      const fields = profile(raw);
      const recommendedFields = fields
        .filter((field) => ["Email", "Phone"].includes(field.type))
        .slice(0, 2)
        .map((field) => field.name);
      const automaticRules: Rule[] = recommendedFields.length
        ? [{
            name: "Recommended identity match",
            fields: recommendedFields,
            method: "Exact",
            threshold: 80,
            priority: 1,
          }]
        : [];
      setConfig((c) => ({
        ...c,
        fileName: file.name,
        fileSize: file.size,
        rows: raw,
        fields,
        rules: automaticRules,
        completed: c.completed.map((done, i) =>
          i === 0 || i === 1 || i === 2 ? true : done,
        ),
      }));
      setActive("Data Profiling");
      log(
        "Dataset uploaded",
        `${raw.length.toLocaleString()} records and ${fields.length} attributes detected`,
      );
      log("Data profiling completed", "Actual uploaded records analyzed");
      if (automaticRules.length) {
        log(
          "Match rule recommended",
          "Draft rule created from high-confidence uploaded attributes",
          "Warning",
        );
      }
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Unable to analyze dataset.",
      );
      log(
        "Dataset upload failed",
        "Please upload a valid CSV, XLSX, or XLS file",
        "Failed",
      );
    }
  };
  const profile = (rows: Row[]): Field[] =>
    Object.keys(rows[0]).map((name) => {
      const values = rows
        .map((r) => String(r[name] ?? "").trim())
        .filter(Boolean);
      const lower = values.map((v) => v.toLowerCase());
      const unique = new Set(lower).size;
      const email =
        values.filter((v) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)).length /
        Math.max(values.length, 1);
      const phone =
        values.filter((v) => /^\+?[\d\s().-]{7,}$/.test(v)).length /
        Math.max(values.length, 1);
      const date =
        values.filter((v) => !Number.isNaN(Date.parse(v))).length /
        Math.max(values.length, 1);
      const numeric =
        values.filter((v) => !Number.isNaN(Number(v))).length /
        Math.max(values.length, 1);
      const recommended =
        email > 0.8
          ? "Email"
          : phone > 0.8
            ? "Phone"
            : date > 0.8
              ? "Date"
              : numeric > 0.8
                ? "Decimal"
                : "String";
      return {
        name,
        type: recommended,
        recommended,
        confidence: Math.round(
          Math.max(email, phone, date, numeric, 0.55) * 100,
        ),
        nulls: rows.length - values.length,
        unique,
        examples: values.slice(0, 3),
        accepted: true,
        required: false,
      };
    });
  const updateField = (name: string, patch: Partial<Field>) =>
    setConfig((c) => ({
      ...c,
      fields: c.fields.map((f) => (f.name === name ? { ...f, ...patch } : f)),
    }));
  const generate = () => {
    const entityName =
      config.project.entityType === "Custom Entity Type"
        ? config.project.customEntityType?.trim() || "CustomEntity"
        : config.project.entityType;
    const json = {
      schemaVersion: "1.0",
      project: {
        name: config.project.name,
        tenant: config.project.tenant,
        environment: config.project.environment,
        sourceSystem: config.project.source,
        dataset: {
          fileName: config.fileName,
          recordCount: config.rows.length,
          attributeCount: config.fields.length,
        },
      },
      entityTypes: [
        {
          name: entityName,
          attributes: config.fields
            .filter((f) => f.accepted)
            .map((f) => ({ name: f.name, type: f.type, required: f.required })),
        },
      ],
      matchGroups: config.rules.map((rule) => ({
        name: rule.name,
        matchCriteria: rule.fields.map((attribute) => ({
          attribute,
          matchType: rule.method,
        })),
        threshold: rule.threshold,
        priority: rule.priority,
      })),
      survivorshipGroups: Object.entries(config.survivorship).map(
        ([attribute, strategy]) => ({ attribute, strategy }),
      ),
      dataValidationFunctions: config.validations.map((name) => ({
        name,
        enabled: true,
      })),
    };
    setConfig((c) => ({ ...c, generated: JSON.stringify(json, null, 2) }));
    mark(9);
    log("JSON generated", "Generated from current user-approved configuration");
  };
  const filtered = useMemo(
    () =>
      config.rows.filter((row) =>
        Object.values(row).some((value) =>
          String(value).toLowerCase().includes(search.toLowerCase()),
        ),
      ),
    [config.rows, search],
  );
  const current = steps.indexOf(active);
  const done = config.completed.filter(Boolean).length;
  const canAccess = (i: number) => i === 0 || config.completed[i - 1];
  const go = (label: string) => {
    const i = steps.indexOf(label);
    if (canAccess(i)) setActive(label);
  };
  const applyPending = () => {
    if (pendingApply === "Match Rules") {
      mark(5);
      log(
        "Match rules applied",
        `${config.rules.length} rule(s) will be included in JSON`,
      );
      setActive("Survivorship");
    }
    if (pendingApply === "Survivorship") {
      mark(6);
      log(
        "Survivorship applied",
        `${Object.keys(config.survivorship).length} strategy(ies) will be included in JSON`,
      );
      setActive("Data Validation");
    }
    setPendingApply(null);
  };
  return (
    <div className="shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">
            <Sparkles size={16} />
          </div>
          <strong>
            config<span>flow</span>
          </strong>
          <small>for Reltio</small>
        </div>
        <div className="workspace">
          <span>WORKSPACE</span>
          <strong>Configuration Builder</strong>
          <ChevronDown size={14} />
        </div>
        <nav>
          {nav.map(([section, items]) => (
            <div className="nav-group" key={section}>
              <label>{section}</label>
              {items.map(([label, Icon]) => (
                <button
                  className={
                    active === label ? "nav-link selected" : "nav-link"
                  }
                  onClick={() => go(label)}
                  key={label}
                >
                  <Icon size={16} />
                  <span>{label}</span>
                  {steps.indexOf(label) > -1 &&
                    !canAccess(steps.indexOf(label)) && (
                      <span className="lock">Locked</span>
                    )}
                </button>
              ))}
            </div>
          ))}
        </nav>
        <div className="side-bottom">
          <button className="nav-link">
            <Settings size={16} />
            <span>Settings</span>
          </button>
          <div className="secure">
            <ShieldCheck size={15} />
            <span>
              Credentials stay
              <br />
              on the backend
            </span>
          </div>
        </div>
      </aside>
      <main className="main">
        <header className="topbar">
          <button className="mobile-menu">
            <Menu size={19} />
          </button>
          <div className="crumb">
            Projects <span>/</span>{" "}
            <strong>{config.project.name || "Untitled project"}</strong>
          </div>
          <div className="top-right">
            <span className="not-connected">
              <i /> Reltio not connected
            </span>
            <Bell size={18} />
            <div className="avatar">RS</div>
          </div>
        </header>
        <div className="page">
          <div className="title-row">
            <div>
              <span className="eyebrow">GUIDED CONFIGURATION</span>
              <h1>
                {active === "Dashboard" ? "Build your configuration" : active}
              </h1>
              <p>
                {config.fileName
                  ? `${config.fileName} is the source of truth for this project.`
                  : "Start with a project and upload your real dataset."}
              </p>
            </div>
            <button className="primary" onClick={() => setProjectOpen(true)}>
              <Plus size={16} /> New project
            </button>
          </div>
          {error && (
            <div className="error">
              <CircleAlert size={16} /> {error}
              <button onClick={() => setError("")}>
                <X size={14} />
              </button>
            </div>
          )}
          <div className="project-card">
            <div className="project-icon">
              <Users size={20} />
            </div>
            <div className="project-info">
              <span>ACTIVE PROJECT</span>
              <strong>{config.project.name || "No project selected"}</strong>
              <small>
                {config.project.entityType} · {config.project.environment} ·{" "}
                {config.project.tenant || "Tenant not configured"}
              </small>
            </div>
            <div className="project-stat">
              <span>RECORDS</span>
              <strong>{config.rows.length.toLocaleString()}</strong>
            </div>
            <div className="project-stat">
              <span>ATTRIBUTES</span>
              <strong>{config.fields.length}</strong>
            </div>
          </div>
          <section className="workflow">
            <div className="section-head">
              <div>
                <span className="eyebrow">WORKFLOW PROGRESS</span>
                <h2>Complete actions in order</h2>
              </div>
              <strong>
                {done} / {steps.length} complete
              </strong>
            </div>
            <div className="steps">
              {steps.map((step, i) => (
                <button
                  key={step}
                  className={`${config.completed[i] ? "complete" : i === current ? "current" : ""} ${!canAccess(i) ? "disabled" : ""}`}
                  onClick={() => go(step)}
                >
                  <b>{config.completed[i] ? <Check size={12} /> : i + 1}</b>
                  <span>{step}</span>
                </button>
              ))}
            </div>
          </section>
          {active === "Dashboard" && (
            <Dashboard
              config={config}
              setActive={setActive}
              upload={() => input.current?.click()}
            />
          )}{" "}
          {active === "Data Upload" && (
            <UploadView config={config} upload={() => input.current?.click()} />
          )}{" "}
          {active === "Data Profiling" && (
            <Profiling
              config={config}
              setActive={setActive}
              updateField={updateField}
            />
          )}{" "}
          {active === "Entity Configuration" && (
            <EntityView
              config={config}
              setConfig={setConfig}
              updateField={updateField}
              complete={() => {
                mark(3);
                log(
                  "Entity configuration completed",
                  `${config.fields.length} attributes reviewed`,
                );
                setActive("Match Analysis");
              }}
            />
          )}{" "}
          {active === "Match Analysis" && (
            <MatchView
              config={config}
              setConfig={setConfig}
              complete={() => {
                mark(4);
                log(
                  "Match analysis completed",
                  "Recommendations reviewed from actual dataset",
                );
                setActive("Match Rules");
              }}
            />
          )}{" "}
          {active === "Match Rules" && (
            <RulesView
              config={config}
              setConfig={setConfig}
              complete={() => setPendingApply("Match Rules")}
            />
          )}{" "}
          {active === "Survivorship" && (
            <Survivorship
              config={config}
              setConfig={setConfig}
              complete={() => setPendingApply("Survivorship")}
            />
          )}{" "}
          {active === "Data Validation" && (
            <Validation
              config={config}
              setConfig={setConfig}
              complete={() => {
                mark(7);
                log(
                  "Validation configured",
                  `${config.validations.length} validations saved`,
                );
                setActive("Configuration Review");
              }}
            />
          )}{" "}
          {active === "Configuration Review" && (
            <Review
              config={config}
              complete={() => {
                mark(8);
                log(
                  "Configuration review completed",
                  "No blocking errors found",
                );
                setActive("JSON Generation");
              }}
            />
          )}{" "}
          {active === "JSON Generation" && (
            <JsonView config={config} generate={generate} />
          )}{" "}
          {active === "Deployment" && <Deployment config={config} />}
          <input
            ref={input}
            hidden
            type="file"
            accept=".csv,.xlsx,.xls"
            onChange={parseFile}
          />
        </div>
      </main>
      {projectOpen && (
        <ProjectModal
          config={config}
          setConfig={setConfig}
          close={() => setProjectOpen(false)}
        />
      )}
      {pendingApply && (
        <ApplyModal
          kind={pendingApply}
          count={pendingApply === "Match Rules" ? config.rules.length : Object.keys(config.survivorship).length}
          apply={applyPending}
          close={() => setPendingApply(null)}
        />
      )}
    </div>
  );
}

function Dashboard({
  config,
  setActive,
  upload,
}: {
  config: Config;
  setActive: (s: string) => void;
  upload: () => void;
}) {
  return (
    <>
      <div className="metric-grid">
        <Metric
          title="Records uploaded"
          value={config.rows.length.toLocaleString()}
          icon={Database}
        />
        <Metric
          title="Detected attributes"
          value={String(config.fields.length)}
          icon={SlidersHorizontal}
        />
        <Metric
          title="Actions complete"
          value={`${config.completed.filter(Boolean).length}/11`}
          icon={Activity}
        />
      </div>
      <div className="two-col">
        <section className="panel">
          <div className="panel-title">
            <div>
              <span className="eyebrow">GET STARTED</span>
              <h2>
                {config.fileName
                  ? "Your data is ready"
                  : "Upload your real dataset"}
              </h2>
              <p>
                {config.fileName
                  ? "Continue reviewing actual records and data-driven recommendations."
                  : "CSV, XLSX, and XLS files are parsed in your browser. No fake records are added."}
              </p>
            </div>
          </div>
          {config.fileName ? (
            <div className="ready">
              <Check size={20} />
              <div>
                <strong>{config.fileName}</strong>
                <span>
                  {config.rows.length.toLocaleString()} records ·{" "}
                  {config.fields.length} columns ·{" "}
                  {(config.fileSize / 1024).toFixed(1)} KB
                </span>
              </div>
              <button
                className="secondary"
                onClick={() => setActive("Data Upload")}
              >
                View data
              </button>
            </div>
          ) : (
            <button className="dropzone" onClick={upload}>
              <UploadCloud size={25} />
              <strong>Choose a dataset to begin</strong>
              <span>Supported: CSV, XLSX, XLS</span>
            </button>
          )}
          <div className="data-truth">
            <ShieldCheck size={16} />
            <span>
              <strong>Data integrity guardrail</strong> Actual records, system
              recommendations, and your decisions are tracked separately.
            </span>
          </div>
        </section>
        <ActivityPanel logs={config.logs} />
      </div>
    </>
  );
}
function UploadView({
  config,
  upload,
}: {
  config: Config;
  upload: () => void;
}) {
  const [query, setQuery] = useState("");
  const [page, setPage] = useState(0);
  const filtered = config.rows.filter((row) =>
    Object.values(row).some((value) =>
      String(value).toLowerCase().includes(query.toLowerCase()),
    ),
  );
  const visible = filtered.slice(page * 10, page * 10 + 10);
  const columns = config.fields.map((field) => field.name);
  return (
    <div className="two-col">
      <section className="panel">
        <div className="panel-title">
          <span className="eyebrow">DATA UPLOAD</span>
          <h2>Bring your source data</h2>
          <p>
            Upload the dataset you want to analyze. The first worksheet or CSV
            table will be used.
          </p>
        </div>
        <button className="dropzone large" onClick={upload}>
          <FileSpreadsheet size={28} />
          <strong>{config.fileName || "Drag and drop or browse"}</strong>
          <span>CSV, XLSX, XLS · No placeholder records</span>
        </button>
        {config.fileName && (
          <>
            <div className="upload-summary">
              <strong>Dataset uploaded successfully</strong>
              <span>
                File size {(config.fileSize / 1024).toFixed(1)} KB ·{" "}
                {config.rows.length.toLocaleString()} records ·{" "}
                {config.fields.length} columns
              </span>
            </div>
            <div className="preview-head">
              <div>
                <span className="eyebrow">REAL DATA PREVIEW</span>
                <strong>
                  {filtered.length.toLocaleString()} matching records
                </strong>
              </div>
              <label className="search">
                <Search size={14} />
                <input
                  value={query}
                  onChange={(e) => {
                    setQuery(e.target.value);
                    setPage(0);
                  }}
                  placeholder="Search records"
                />
              </label>
            </div>
            {visible.length ? (
              <div className="data-table">
                <div className="data-row data-header">
                  {columns.map((column) => (
                    <span key={column}>{column}</span>
                  ))}
                </div>
                {visible.map((row, index) => (
                  <div className="data-row" key={index}>
                    {columns.map((column) => (
                      <span key={column}>
                        {String(row[column] ?? "") || "—"}
                      </span>
                    ))}
                  </div>
                ))}
              </div>
            ) : (
              <NoRecords text="No records available." />
            )}
            <div className="pagination">
              <span>
                Showing {visible.length} of {filtered.length.toLocaleString()}{" "}
                visible records
              </span>
              <div>
                <button
                  className="secondary"
                  disabled={page === 0}
                  onClick={() => setPage(page - 1)}
                >
                  Previous
                </button>
                <button
                  className="secondary"
                  disabled={(page + 1) * 10 >= filtered.length}
                  onClick={() => setPage(page + 1)}
                >
                  Next
                </button>
              </div>
            </div>
          </>
        )}
      </section>
      <ActivityPanel logs={config.logs} />
    </div>
  );
}
function Profiling({
  config,
  setActive,
  updateField,
}: {
  config: Config;
  setActive: (s: string) => void;
  updateField: (name: string, patch: Partial<Field>) => void;
}) {
  return (
    <section className="panel full">
      <div className="panel-title">
        <span className="eyebrow">ACTUAL DATA ANALYSIS</span>
        <h2>Data profiling</h2>
        <p>
          Types and recommendations are inferred from values in your uploaded
          dataset. Accept or override each recommendation.
        </p>
      </div>
      {config.fields.length ? (
        <div className="profile-table">
          <div className="profile-header">
            <span>Source attribute</span>
            <span>Detected type</span>
            <span>Recommendation</span>
            <span>Confidence</span>
            <span>Examples from data</span>
            <span>Decision</span>
          </div>
          {config.fields.map((f) => (
            <div className="profile-row" key={f.name}>
              <strong>{f.name}</strong>
              <select
                value={f.type}
                onChange={(e) =>
                  updateField(f.name, { type: e.target.value, accepted: false })
                }
              >
                <option>String</option>
                <option>Email</option>
                <option>Phone</option>
                <option>Date</option>
                <option>DateTime</option>
                <option>Integer</option>
                <option>Decimal</option>
                <option>Boolean</option>
              </select>
              <span>
                <b className="recommend">{f.recommended}</b>
                <small>
                  {f.nulls} null · {f.unique} unique
                </small>
              </span>
              <span>{f.confidence}%</span>
              <span className="examples">
                {f.examples.join(" · ") || "No values available"}
              </span>
              <button
                className={f.accepted ? "decision accepted" : "decision"}
                onClick={() => updateField(f.name, { accepted: !f.accepted })}
              >
                {f.accepted ? "Accepted" : "Override saved"}
              </button>
            </div>
          ))}
        </div>
      ) : (
        <NoRecords text="No records available. Upload a dataset to start profiling." />
      )}
      <div className="footer-actions">
        <span>
          Recommendations never change your final configuration without your
          decision.
        </span>
        <button
          className="primary"
          disabled={!config.fields.length}
          onClick={() => setActive("Entity Configuration")}
        >
          Review attributes <span>→</span>
        </button>
      </div>
    </section>
  );
}
function EntityView({
  config,
  setConfig,
  updateField,
  complete,
}: {
  config: Config;
  setConfig: (config: Config) => void;
  updateField: (name: string, patch: Partial<Field>) => void;
  complete: () => void;
}) {
  return (
    <section className="panel full">
      <div className="panel-title">
        <span className="eyebrow">ENTITY CONFIGURATION</span>
        <h2>Which entity are you configuring?</h2>
        <p>
          Map accepted source attributes to the selected Reltio entity type.
        </p>
      </div>
      <label className="field-label">
        Entity Type
        <select
          value={config.project.entityType}
          onChange={(e) => {
            setConfig({
              ...config,
              project: { ...config.project, entityType: e.target.value },
            });
          }}
        >
          <option>Individual</option>
          <option>Organization</option>
          <option>Product</option>
          <option>Location</option>
          <option>Custom Entity Type</option>
        </select>
      </label>
      {config.project.entityType === "Custom Entity Type" && (
        <label className="field-label">
          Custom Entity Type Name
          <input
            value={config.project.customEntityType || ""}
            onChange={(e) =>
              setConfig({
                ...config,
                project: {
                  ...config.project,
                  customEntityType: e.target.value,
                },
              })
            }
            placeholder="e.g. Customer"
          />
        </label>
      )}
      <div className="entity-list">
        {config.fields.map((f) => (
          <div className="entity-row" key={f.name}>
            <strong>{f.name}</strong>
            <span>→</span>
            <select
              value={f.type}
              onChange={(e) => updateField(f.name, { type: e.target.value })}
            >
              <option>String</option>
              <option>Email</option>
              <option>Phone</option>
              <option>Date</option>
              <option>Integer</option>
              <option>Decimal</option>
            </select>
            <label>
              <input
                type="checkbox"
                checked={f.required}
                onChange={(e) =>
                  updateField(f.name, { required: e.target.checked })
                }
              />{" "}
              Required
            </label>
          </div>
        ))}
      </div>
      <div className="footer-actions">
        <span>
          Selected: <strong>{config.project.entityType}</strong>
        </span>
        <button className="primary" onClick={complete}>
          Save entity configuration →
        </button>
      </div>
    </section>
  );
}
function MatchView({
  config,
  setConfig,
  complete,
}: {
  config: Config;
  setConfig: (c: Config) => void;
  complete: () => void;
}) {
  const candidates = config.fields.filter(
    (f) =>
      ["Email", "Phone"].includes(f.type) ||
      f.unique > config.rows.length * 0.7,
  );
  const acceptRecommendation = (fieldName: string) => {
    const field = config.fields.find((f) => f.name === fieldName);
    if (!field) return;
    
    // Create a new rule from the accepted recommendation
    const newRule: Rule = {
      name: `${fieldName} Match`,
      fields: [fieldName],
      method: field.type === "Email" || field.type === "Phone" ? "Exact" : "Normalized",
      threshold: 80,
      priority: 1,
    };
    
    setConfig({
      ...config,
      rules: [...config.rules, newRule],
    });
    
    alert(`✅ "${fieldName}" added as a matching rule!`);
  };
  return (
    <section className="panel full">
      <div className="panel-title">
        <span className="eyebrow">MATCH ANALYSIS</span>
        <h2>Review match recommendations</h2>
        <p>
          Recommendations are calculated from uniqueness, null rate, and
          detected value patterns in your actual dataset.
        </p>
      </div>
      {candidates.length ? (
        <div className="recommend-grid">
          {candidates.map((f) => (
            <div className="recommend-card" key={f.name}>
              <div>
                <strong>{f.name}</strong>
                <b>HIGH VALUE</b>
              </div>
              <span>
                Uniqueness:{" "}
                {config.rows.length
                  ? Math.round((f.unique / config.rows.length) * 100)
                  : 0}
                %
              </span>
              <small>
                Recommended:{" "}
                {f.type === "Email" || f.type === "Phone"
                  ? "Exact match"
                  : "Normalized match"}
              </small>
              <button className="secondary" onClick={() => acceptRecommendation(f.name)}>Accept recommendation</button>
            </div>
          ))}
        </div>
      ) : (
        <NoRecords text="No high-confidence match candidates found in the uploaded records." />
      )}
      <div className="footer-actions">
        <span>{candidates.length} candidates analyzed</span>
        <button className="primary" onClick={complete}>
          Save match analysis →
        </button>
      </div>
    </section>
  );
}
function RulesView({
  config,
  setConfig,
  complete,
}: {
  config: Config;
  setConfig: (c: Config) => void;
  complete: () => void;
}) {
  const add = () =>
    setConfig({
      ...config,
      rules: [
        ...config.rules,
        {
          name: `Rule ${config.rules.length + 1}`,
          fields: config.fields.slice(0, 2).map((f) => f.name),
          method: "Exact",
          threshold: 80,
          priority: config.rules.length + 1,
        },
      ],
    });
  return (
    <section className="panel full">
      <div className="panel-title">
        <span className="eyebrow">MATCH RULE BUILDER</span>
        <h2>Configure matching rules</h2>
        <p>
          Build rules from the attributes detected in your uploaded dataset.
        </p>
      </div>
      <button className="secondary" onClick={add}>
        <Plus size={15} /> Add rule
      </button>
      {config.rules.map((r, i) => (
        <div className="rule-card" key={r.name}>
          <input
            value={r.name}
            onChange={(e) => {
              const rules = [...config.rules];
              rules[i] = { ...r, name: e.target.value };
              setConfig({ ...config, rules });
            }}
          />
          <select
            value={r.method}
            onChange={(e) => {
              const rules = [...config.rules];
              rules[i] = { ...r, method: e.target.value };
              setConfig({ ...config, rules });
            }}
          >
            <option>Exact</option>
            <option>Normalized</option>
            <option>Fuzzy</option>
          </select>
          <strong>{r.fields.join(" AND ") || "Choose attributes"}</strong>
          <label>
            Threshold{" "}
            <input
              type="number"
              value={r.threshold}
              onChange={(e) => {
                const rules = [...config.rules];
                rules[i] = { ...r, threshold: Number(e.target.value) };
                setConfig({ ...config, rules });
              }}
            />
          </label>
        </div>
      ))}
      {!config.rules.length && (
        <NoRecords text="No rules configured yet. Add a rule using actual attributes." />
      )}
      <div className="footer-actions">
        <span>These rules will be written to the final JSON.</span>
        <button className="primary" onClick={complete}>
          Save match rules →
        </button>
      </div>
    </section>
  );
}
function Survivorship({
  config,
  setConfig,
  complete,
}: {
  config: Config;
  setConfig: (c: Config) => void;
  complete: () => void;
}) {
  const strategies = [
    {
      name: "LUD",
      label: "Last Update Date",
      description: "Attributes with the most recent update dates have highest priority. This is the default strategy."
    },
    {
      name: "SRC_SYS",
      label: "Source System",
      description: "User provides a priority list of sources. Highest priority source values become winners."
    },
    {
      name: "Frequency",
      label: "Frequency",
      description: "Values with the most crosswalks become winners. Uses LUD for final selection if tied."
    },
    {
      name: "Aggregation",
      label: "Aggregation",
      description: "All presented values become winners with OV (Operational Value) set to True."
    },
    {
      name: "RecordValuesSourcePriority",
      label: "Record-values Based Source Priority",
      description: "All values from the winner source type become winners."
    },
    {
      name: "OldestValue",
      label: "Oldest Value",
      description: "Finds the crosswalk with the oldest create date. All values from this crosswalk become winners."
    },
    {
      name: "MinValue",
      label: "Min Value",
      description: "The smallest value becomes the winner (numeric, date, or lexicographical order)."
    },
    {
      name: "MaxValue",
      label: "Max Value",
      description: "The largest value becomes the winner (numeric, date, or lexicographical order)."
    },
    {
      name: "OtherAttributeWinnerCrosswalk",
      label: "Other Attribute Winner Crosswalk",
      description: "Gets winner crosswalks from a primary attribute. Depends on another attribute's OV calculation."
    },
    {
      name: "WinnerEntityCrosswalk",
      label: "Winner Entity Crosswalk",
      description: "The entity value of the current winning entity becomes the winner."
    },
    {
      name: "ValueBasedPriority",
      label: "Value-based Priority",
      description: "User-defined priority list of attribute values determines the OV."
    }
  ];

  const generateFromMatchRules = () => {
    if (config.rules.length === 0) {
      alert("❌ No match rules found. Please create match rules first!");
      return;
    }

    const newSurvivorship: Record<string, string> = {
      ...config.survivorship,
    };

    // Generate survivorship rules for all fields used in match rules
    config.rules.forEach((rule) => {
      rule.fields.forEach((fieldName) => {
        if (!newSurvivorship[fieldName]) {
          newSurvivorship[fieldName] = "LUD"; // Default LUD strategy
        }
      });
    });

    setConfig({
      ...config,
      survivorship: newSurvivorship,
    });

    alert(`✅ Generated survivorship strategies for ${config.rules.length} match rule(s) using LUD (Last Update Date)!`);
  };

  return (
    <section className="panel full">
      <div className="panel-title">
        <span className="eyebrow">SURVIVORSHIP</span>
        <h2>Choose the winning value strategy</h2>
        <p>
          Select how Reltio should resolve conflicting values for each
          configured attribute based on Reltio survivorship strategies.
        </p>
      </div>
      
      {config.rules.length > 0 && (
        <div style={{ marginBottom: "24px" }}>
          <div style={{ 
            background: "#edf3ff", 
            border: "1px solid #d7e4fb", 
            borderRadius: "5px", 
            padding: "12px",
            marginBottom: "12px"
          }}>
            <strong style={{ fontSize: "12px", color: "#34526b" }}>
              📋 Auto-generate from Match Rules
            </strong>
            <p style={{ fontSize: "11px", color: "#718793", margin: "4px 0 8px" }}>
              {config.rules.length} match rule(s) found. Click to automatically create survivorship strategies using LUD (Last Update Date) - Reltio's default.
            </p>
            <button 
              className="secondary"
              onClick={generateFromMatchRules}
              style={{ fontSize: "11px", padding: "6px 12px" }}
            >
              Generate from Match Rules
            </button>
          </div>
        </div>
      )}

      <div className="strategy-list">
        {config.fields.map((f) => (
          <div key={f.name} style={{ 
            marginBottom: "16px",
            padding: "12px",
            border: "1px solid #e3e9ec",
            borderRadius: "5px",
            background: "#fafbfc"
          }}>
            <strong style={{ display: "block", marginBottom: "8px" }}>{f.name}</strong>
            <select
              value={config.survivorship[f.name] || "LUD"}
              onChange={(e) =>
                setConfig({
                  ...config,
                  survivorship: {
                    ...config.survivorship,
                    [f.name]: e.target.value,
                  },
                })
              }
              style={{ width: "100%", padding: "8px", marginBottom: "6px" }}
            >
              <option value="">-- Select Strategy --</option>
              {strategies.map((strategy) => (
                <option key={strategy.name} value={strategy.name}>
                  {strategy.label}
                </option>
              ))}
            </select>
            
            {config.survivorship[f.name] && (
              <div style={{ 
                fontSize: "10px", 
                color: "#667681",
                marginTop: "6px",
                padding: "6px",
                background: "#f9fbfc",
                borderRadius: "3px",
                borderLeft: "2px solid #1a9a6e"
              }}>
                <strong style={{ color: "#1a9a6e" }}>✓ Selected:</strong> {" "}
                {strategies.find(s => s.name === config.survivorship[f.name])?.description || config.survivorship[f.name]}
              </div>
            )}
          </div>
        ))}
      </div>

      {Object.keys(config.survivorship).length > 0 && (
        <div style={{ 
          marginTop: "20px", 
          padding: "14px", 
          background: "#f0f7f4", 
          borderRadius: "5px",
          border: "1px solid #d0e8e1"
        }}>
          <strong style={{ fontSize: "13px", color: "#1a5f52" }}>📊 Created Survivorship Rules:</strong>
          <ul style={{ fontSize: "11px", margin: "10px 0", color: "#34526b", paddingLeft: "20px" }}>
            {Object.entries(config.survivorship).map(([field, strategy]) => {
              const strategyInfo = strategies.find(s => s.name === strategy);
              return (
                <li key={field} style={{ marginBottom: "6px" }}>
                  <strong>{field}</strong> → <em>{strategyInfo?.label || strategy}</em>
                  <div style={{ fontSize: "10px", color: "#718793", marginTop: "2px" }}>
                    {strategyInfo?.description}
                  </div>
                </li>
              );
            })}
          </ul>
        </div>
      )}

      <div className="footer-actions">
        <span>
          {Object.keys(config.survivorship).length} strategy(ies) configured
        </span>
        <button className="primary" onClick={complete}>
          Save survivorship →
        </button>
      </div>
    </section>
  );
}
function Validation({
  config,
  setConfig,
  complete,
}: {
  config: Config;
  setConfig: (c: Config) => void;
  complete: () => void;
}) {
  const options = [
    "Required fields",
    "Email format",
    "Phone format",
    "Date format",
    "Maximum length",
  ];
  return (
    <section className="panel full">
      <div className="panel-title">
        <span className="eyebrow">DATA VALIDATION</span>
        <h2>Define quality checks</h2>
        <p>
          Choose validations to run against the actual uploaded records before
          configuration review.
        </p>
      </div>
      <div className="validation-options">
        {options.map((option) => (
          <label key={option}>
            <input
              type="checkbox"
              checked={config.validations.includes(option)}
              onChange={(e) =>
                setConfig({
                  ...config,
                  validations: e.target.checked
                    ? [...config.validations, option]
                    : config.validations.filter((v) => v !== option),
                })
              }
            />{" "}
            {option}
          </label>
        ))}
      </div>
      <div className="validation-note">
        <ShieldCheck size={18} />
        <span>
          Validation results are tied to this dataset and will be included as
          configuration metadata, not entity records.
        </span>
      </div>
      <div className="footer-actions">
        <span>{config.validations.length} validation rules selected</span>
        <button className="primary" onClick={complete}>
          Save validations →
        </button>
      </div>
    </section>
  );
}
function Review({
  config,
  complete,
}: {
  config: Config;
  complete: () => void;
}) {
  const errors = !config.project.entityType || !config.fields.length;
  return (
    <section className="panel full">
      <div className="panel-title">
        <span className="eyebrow">CONFIGURATION REVIEW</span>
        <h2>Review before generating JSON</h2>
        <p>
          This is the last checkpoint. JSON generation remains locked until this
          review is complete.
        </p>
      </div>
      <div className={`review-status ${errors ? "bad" : "good"}`}>
        <ShieldCheck size={19} />
        <strong>
          {errors ? "Blocking issues found" : "No blocking errors"}
        </strong>
        <span>
          {errors
            ? "Upload data and configure an entity first."
            : "Ready for JSON generation after you confirm this review."}
        </span>
      </div>
      <div className="review-grid">
        <ReviewItem label="Dataset" value={config.fileName || "No dataset"} />
        <ReviewItem label="Entity type" value={config.project.entityType} />
        <ReviewItem label="Attributes" value={String(config.fields.length)} />
        <ReviewItem label="Match rules" value={String(config.rules.length)} />
        <ReviewItem
          label="Survivorship"
          value={String(Object.keys(config.survivorship).length)}
        />
        <ReviewItem
          label="Validations"
          value={String(config.validations.length)}
        />
      </div>
      <div className="footer-actions">
        <span>Actual records: {config.rows.length.toLocaleString()}</span>
        <button className="primary" disabled={errors} onClick={complete}>
          Confirm review →
        </button>
      </div>
    </section>
  );
}
function ReviewItem({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}
function JsonView({
  config,
  generate,
}: {
  config: Config;
  generate: () => void;
}) {
  const valid = Boolean(config.generated);
  const download = () => {
    const blob = new Blob([config.generated], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${(config.project.name || "configuration").toLowerCase().replace(/[^a-z0-9]+/g, "-")}-configuration.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  };
  return (
    <section className="panel full">
      <div className="panel-title json-head">
        <div>
          <span className="eyebrow">FINAL OUTPUT</span>
          <h2>JSON Generator</h2>
          <p>
            Generated from the current configuration state, never from a static
            template.
          </p>
        </div>
        <div className="json-actions">
          <button className="secondary" onClick={generate}>
            <Play size={14} /> {valid ? "Regenerate" : "Generate JSON"}
          </button>
          <button className="secondary" disabled={!valid} onClick={download}>
            <Download size={14} /> Download
          </button>
        </div>
      </div>
      {valid ? (
        <>
          <div className="json-valid">
            <Check size={15} /> Valid JSON · configuration fields present ·
            ready to download
          </div>
          <pre className="code">
            <code>{config.generated}</code>
          </pre>
        </>
      ) : (
        <NoRecords text="JSON generation is waiting for completed configuration review." />
      )}
    </section>
  );
}
function Deployment({ config }: { config: Config }) {
  return (
    <section className="panel full">
      <div className="panel-title">
        <span className="eyebrow">OPTIONAL DEPLOYMENT</span>
        <h2>Deploy to Reltio</h2>
        <p>
          Deployment is intentionally unavailable until a real backend
          connection is verified.
        </p>
      </div>
      <div className="not-ready">
        <CircleAlert size={19} />
        <strong>Reltio not connected</strong>
        <span>
          Configure and verify a backend connection before deployment. No
          deployment was attempted.
        </span>
      </div>
      <button className="secondary" disabled>
        <Zap size={14} /> Deploy configuration
      </button>
      <div className="review-grid">
        <ReviewItem label="Environment" value={config.project.environment} />
        <ReviewItem
          label="Configuration"
          value={config.generated ? "Generated" : "Not generated"}
        />
      </div>
    </section>
  );
}
function ActivityPanel({ logs }: { logs: LogItem[] }) {
  return (
    <aside className="activity panel">
      <div className="panel-title">
        <span className="eyebrow">LIVE ACTIVITY</span>
        <h2>What happened</h2>
      </div>
      {logs.length ? (
        logs.slice(0, 8).map((item, i) => (
          <div className="log" key={`${item.time}-${i}`}>
            <i className={item.status.toLowerCase()} />
            <div>
              <strong>{item.action}</strong>
              <span>{item.details}</span>
              <small>
                {item.time} · {item.status}
              </small>
            </div>
          </div>
        ))
      ) : (
        <NoRecords text="No activity yet." />
      )}
    </aside>
  );
}
function Metric({
  title,
  value,
  icon: Icon,
}: {
  title: string;
  value: string;
  icon: typeof Database;
}) {
  return (
    <div className="metric panel">
      <Icon size={18} />
      <span>{title}</span>
      <strong>{value}</strong>
    </div>
  );
}
function NoRecords({ text }: { text: string }) {
  return (
    <div className="no-records">
      <Database size={20} />
      <span>{text}</span>
    </div>
  );
}
function ProjectModal({
  config,
  setConfig,
  close,
}: {
  config: Config;
  setConfig: (config: Config) => void;
  close: () => void;
}) {
  const [project, setProject] = useState(config.project);
  const save = () => {
    setConfig({
      ...config,
      project,
      completed: config.completed.map((done, index) =>
        index === 0 ? Boolean(project.name && project.entityType) : done,
      ),
    });
    close();
  };
  return (
    <div className="modal-backdrop">
      <div className="modal">
        <button className="modal-close" onClick={close}>
          <X size={17} />
        </button>
        <div className="modal-icon">
          <FolderKanban size={19} />
        </div>
        <span className="eyebrow">PROJECT SETUP</span>
        <h2>Create or edit project</h2>
        <p>
          Project metadata is configuration metadata and will be kept separate
          from uploaded records.
        </p>
        <label>
          Project name
          <input
            autoFocus
            value={project.name}
            onChange={(e) => setProject({ ...project, name: e.target.value })}
            placeholder="Customer 360 Foundation"
          />
        </label>
        <label>
          Reltio tenant name
          <input
            value={project.tenant}
            onChange={(e) => setProject({ ...project, tenant: e.target.value })}
            placeholder="Tenant name"
          />
        </label>
        <div className="form-grid">
          <label>
            Environment
            <select
              value={project.environment}
              onChange={(e) =>
                setProject({ ...project, environment: e.target.value })
              }
            >
              <option>Development</option>
              <option>QA</option>
              <option>Production</option>
            </select>
          </label>
          <label>
            Entity type
            <select
              value={project.entityType}
              onChange={(e) =>
                setProject({ ...project, entityType: e.target.value })
              }
            >
              <option>Individual</option>
              <option>Organization</option>
              <option>Product</option>
              <option>Location</option>
              <option>Custom Entity Type</option>
            </select>
          </label>
        </div>
        <label>
          Source system
          <input
            value={project.source}
            onChange={(e) => setProject({ ...project, source: e.target.value })}
            placeholder="Salesforce CRM"
          />
        </label>
        <div className="modal-actions">
          <button className="secondary" onClick={close}>
            Cancel
          </button>
          <button
            className="primary"
            disabled={!project.name.trim()}
            onClick={save}
          >
            <Check size={15} /> Save project
          </button>
        </div>
      </div>
    </div>
  );
}

function ApplyModal({
  kind,
  count,
  apply,
  close,
}: {
  kind: "Match Rules" | "Survivorship";
  count: number;
  apply: () => void;
  close: () => void;
}) {
  return (
    <div className="modal-backdrop">
      <div className="modal apply-modal">
        <button className="modal-close" onClick={close} aria-label="Close">
          <X size={17} />
        </button>
        <div className="modal-icon"><ClipboardCheck size={19} /></div>
        <span className="eyebrow">USER CONFIRMATION</span>
        <h2>Apply {kind}?</h2>
        <p>
          Your {kind.toLowerCase()} configuration is ready. Applying it will mark
          this action complete and include it in the final JSON.
        </p>
        <div className="apply-summary">
          <strong>{count} {kind === "Match Rules" ? "match rule(s)" : "survivorship strategy(ies)"}</strong>
          <span>Only your current saved choices will be used.</span>
        </div>
        <div className="modal-actions">
          <button className="secondary" onClick={close}>No, keep editing</button>
          <button className="primary" onClick={apply}><Check size={15} /> Yes, apply</button>
        </div>
      </div>
    </div>
  );
}
