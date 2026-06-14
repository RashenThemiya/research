import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  ArrowLeft,
  ArrowRight,
  BarChart3,
  Check,
  Download,
  LogOut,
  Save,
  Search,
  UserPlus,
} from "lucide-react";
import * as XLSX from "xlsx";
import "./styles.css";

const TEMPLATE_URL = "/complte-one-template.xlsx";
const STORAGE_KEY = "response-sheet-app-v1";
const ANSWER_FIELDS = [
  { key: "correctnessClaude", label: "Correctness", cell: "AH", model: "Claude", type: "yesno" },
  { key: "correctnessGpt", label: "Correctness", cell: "AI", model: "GPT5", type: "yesno" },
  { key: "correctnessGemini", label: "Correctness", cell: "AJ", model: "Gemini", type: "yesno" },
  { key: "correctnessDeepSeek", label: "Correctness", cell: "AK", model: "DeepSeek", type: "yesno" },
  { key: "relevanceClaude", label: "Requirement Relevance", cell: "AL", model: "Claude", type: "yesno" },
  { key: "relevanceGpt", label: "Requirement Relevance", cell: "AM", model: "GPT5", type: "yesno" },
  { key: "relevanceGemini", label: "Requirement Relevance", cell: "AN", model: "Gemini", type: "yesno" },
  { key: "relevanceDeepSeek", label: "Requirement Relevance", cell: "AO", model: "DeepSeek", type: "yesno" },
  { key: "readabilityClaude", label: "Readability", cell: "AP", model: "Claude", type: "yesno" },
  { key: "readabilityGpt", label: "Readability", cell: "AQ", model: "GPT5", type: "yesno" },
  { key: "readabilityGemini", label: "Readability", cell: "AR", model: "Gemini", type: "yesno" },
  { key: "readabilityDeepSeek", label: "Readability", cell: "AS", model: "DeepSeek", type: "yesno" },
  { key: "semanticClaude", label: "Semantic Similarity", cell: "AT", model: "Claude", type: "scale" },
  { key: "semanticGpt", label: "Semantic Similarity", cell: "AU", model: "GPT5", type: "scale" },
  { key: "semanticGemini", label: "Semantic Similarity", cell: "AV", model: "Gemini", type: "scale" },
  { key: "semanticDeepSeek", label: "Semantic Similarity", cell: "AW", model: "DeepSeek", type: "scale" },
];
const MODEL_NAMES = ["Claude", "GPT5", "Gemini", "DeepSeek"];
const VALID_COUNT_CELLS = {
  Claude: { fr: "AB", nfr: "AC" },
  DeepSeek: { fr: "AD", nfr: "AE" },
  Gemini: { fr: "AF", nfr: "AG" },
  GPT5: { fr: "Z", nfr: "AA" },
};

function readStore() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || { users: {}, currentPhone: "" };
  } catch {
    return { users: {}, currentPhone: "" };
  }
}

function saveStore(store) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
}

function slug(value) {
  return value.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "user";
}

function normalizePhone(value) {
  return value.replace(/[^\d+]/g, "");
}

function cellText(sheet, address) {
  return sheet[address]?.v == null ? "" : String(sheet[address].v);
}

function parseRows(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const rows = [];
  for (let r = 3; r <= range.e.r; r += 1) {
    const excelRow = r + 1;
    rows.push({
      excelRow,
      sampleNo: cellText(sheet, `A${excelRow}`),
      originalNo: cellText(sheet, `B${excelRow}`),
      source: cellText(sheet, `C${excelRow}`),
      storyNo: cellText(sheet, `D${excelRow}`),
      story: cellText(sheet, `E${excelRow}`),
      claude: cellText(sheet, `F${excelRow}`),
      gpt: cellText(sheet, `G${excelRow}`),
      gemini: cellText(sheet, `H${excelRow}`),
      deepseek: cellText(sheet, `I${excelRow}`),
    });
  }
  return rows;
}

function cleanRequirementText(text) {
  return String(text || "")
    .replace(/\*\*/g, "")
    .replace(/^#+\s*/gm, "")
    .trim();
}

function parseRequirementItems(body) {
  const cleaned = cleanRequirementText(body);
  const lines = cleaned
    .split(/\r?\n/)
    .map((line) => line.replace(/^[-*]\s*/, "").trim())
    .filter(Boolean);
  const requirements = lines.filter((line) => /^(FR|NFR)[-\s]?\d+/i.test(line));
  return requirements.length ? requirements : lines.filter((line) => !/requirements|none/i.test(line));
}

function splitRequirementText(text) {
  const cleaned = cleanRequirementText(text);
  if (/error:/i.test(cleaned)) return [{ title: "Output", type: "output", items: [cleaned || "No output available."] }];
  const functionalMatch = cleaned.match(/(?:functional requirements|## functional requirements|\bfr\b)/i);
  const nonFunctionalMatch = cleaned.match(/(?:non-functional requirements|non functional requirements|## non-functional requirements|\bnfr\b)/i);
  if (!functionalMatch && !nonFunctionalMatch) {
    return [{ title: "Output", type: "output", items: parseRequirementItems(cleaned) }];
  }

  const sections = [];
  const functionalStart = functionalMatch?.index ?? 0;
  const nonFunctionalStart = nonFunctionalMatch?.index;
  if (functionalMatch) {
    sections.push({
      title: "Functional Requirements",
      type: "fr",
      items: parseRequirementItems(cleaned.slice(functionalStart, nonFunctionalStart ?? cleaned.length)),
    });
  }
  if (nonFunctionalMatch) {
    sections.push({
      title: "Non-Functional Requirements",
      type: "nfr",
      items: parseRequirementItems(cleaned.slice(nonFunctionalStart)),
    });
  }
  return sections;
}

function requirementKey(model, type, index) {
  return `${model}:${type}:${index}`;
}

function countValidRequirements(row, answer, model) {
  const textByModel = { Claude: row.claude, GPT5: row.gpt, Gemini: row.gemini, DeepSeek: row.deepseek };
  const validity = answer?.requirementValidity || {};
  return splitRequirementText(textByModel[model]).reduce(
    (counts, section) => {
      if (section.type !== "fr" && section.type !== "nfr") return counts;
      section.items.forEach((_, index) => {
        if (validity[requirementKey(model, section.type, index)] === "valid") counts[section.type] += 1;
      });
      return counts;
    },
    { fr: 0, nfr: 0 },
  );
}

function ModelOutput({ name, text, answer, onValidityChange }) {
  const validity = answer?.requirementValidity || {};
  return (
    <article>
      <h3>{name}</h3>
      <div className="requirement-sections">
        {splitRequirementText(text).map((section) => (
          <section key={section.title} className="requirement-block">
            <h4>{section.title}</h4>
            <ul className="requirement-list">
              {section.items.map((item, index) => {
                const key = requirementKey(name, section.type, index);
                return (
                  <li key={key}>
                    <p>{item}</p>
                    {(section.type === "fr" || section.type === "nfr") && (
                      <ChoiceButtons
                        value={validity[key] || ""}
                        options={["valid", "not valid"]}
                        onChange={(value) => onValidityChange(key, value)}
                      />
                    )}
                  </li>
                );
              })}
            </ul>
          </section>
        ))}
      </div>
    </article>
  );
}

function ChoiceButtons({ value, options, onChange }) {
  return (
    <div className="choice-buttons">
      {options.map((option) => (
        <button
          type="button"
          key={option}
          className={String(value) === String(option) ? "selected" : ""}
          onClick={() => onChange(option)}
        >
          {option}
        </button>
      ))}
    </div>
  );
}

function App() {
  const [store, setStore] = useState(readStore);
  const [workbook, setWorkbook] = useState(null);
  const [rows, setRows] = useState([]);
  const [mode, setMode] = useState(store.currentPhone ? "dashboard" : "auth");
  const [auth, setAuth] = useState({ name: "", phone: "" });
  const [activeIndex, setActiveIndex] = useState(0);
  const [query, setQuery] = useState("");

  useEffect(() => {
    fetch(TEMPLATE_URL)
      .then((res) => res.arrayBuffer())
      .then((buffer) => {
        const wb = XLSX.read(buffer, { type: "array" });
        setWorkbook(wb);
        setRows(parseRows(wb));
      });
  }, []);

  useEffect(() => saveStore(store), [store]);

  const currentUser = store.users[store.currentPhone];
  const answers = currentUser?.answers || {};
  const authPhone = normalizePhone(auth.phone);
  const existingAuthUser = authPhone ? store.users[authPhone] : null;
  const answeredCount = rows.filter((row) => answers[row.excelRow]?.complete).length;
  const progress = rows.length ? Math.round((answeredCount / rows.length) * 100) : 0;
  const activeRow = rows[activeIndex];

  const filteredRows = useMemo(() => {
    const term = query.toLowerCase();
    return rows.filter((row) =>
      [row.sampleNo, row.storyNo, row.source, row.story].some((value) => value.toLowerCase().includes(term)),
    );
  }, [rows, query]);

  function updateStore(next) {
    setStore((prev) => {
      const value = typeof next === "function" ? next(prev) : next;
      return value;
    });
  }

  function signIn(event) {
    event.preventDefault();
    const name = auth.name.trim();
    const phone = normalizePhone(auth.phone);
    if (!phone || (!store.users[phone] && !name)) return;
    updateStore((prev) => ({
      ...prev,
      currentPhone: phone,
      users: {
        ...prev.users,
        [phone]: prev.users[phone] || {
          name,
          phone,
          createdAt: new Date().toISOString(),
          answers: {},
        },
      },
    }));
    setMode("dashboard");
  }

  function saveAnswer(rowNumber, patch) {
    updateStore((prev) => {
      const user = prev.users[prev.currentPhone];
      const existing = user.answers[rowNumber] || {};
      const nextAnswer = {
        ...existing,
        ...patch,
        requirementValidity: {
          ...(existing.requirementValidity || {}),
          ...(patch.requirementValidity || {}),
        },
        updatedAt: new Date().toISOString(),
      };
      const complete =
        ANSWER_FIELDS.some((field) => String(nextAnswer[field.key] || "").trim()) ||
        Object.keys(nextAnswer.requirementValidity || {}).length > 0;
      nextAnswer.complete = complete;
      return {
        ...prev,
        users: {
          ...prev.users,
          [prev.currentPhone]: {
            ...user,
            answers: { ...user.answers, [rowNumber]: nextAnswer },
          },
        },
      };
    });
  }

  function exportWorkbook() {
    if (!workbook || !currentUser) return;
    const copy = XLSX.read(XLSX.write(workbook, { type: "array", bookType: "xlsx" }), { type: "array" });
    const sheet = copy.Sheets[copy.SheetNames[0]];
    Object.entries(currentUser.answers).forEach(([rowNumber, answer]) => {
      const row = rows.find((item) => String(item.excelRow) === String(rowNumber));
      ANSWER_FIELDS.forEach((field) => {
        const value = answer[field.key];
        if (value !== undefined && value !== "") {
          sheet[`${field.cell}${rowNumber}`] = { t: "s", v: String(value) };
        }
      });
      if (row) {
        MODEL_NAMES.forEach((model) => {
          const counts = countValidRequirements(row, answer, model);
          Object.entries(VALID_COUNT_CELLS[model]).forEach(([type, column]) => {
            sheet[`${column}${rowNumber}`] = { t: "n", v: counts[type] };
          });
        });
      }
    });
    XLSX.writeFile(copy, `${slug(currentUser.name)}-${currentUser.phone}-responses.xlsx`);
  }

  if (!rows.length) {
    return <main className="loading">Loading workbook...</main>;
  }

  if (!currentUser || mode === "auth") {
    return (
      <main className="auth-screen">
        <form className="auth-panel" onSubmit={signIn}>
          <div className="brand"><UserPlus size={28} /> Response Sheet</div>
          <h1>{existingAuthUser ? "Welcome back" : "Create account or sign in"}</h1>
          <label>Phone number<input value={auth.phone} onChange={(e) => setAuth({ ...auth, phone: e.target.value })} /></label>
          {existingAuthUser ? (
            <div className="restore-note">
              <strong>{existingAuthUser.name}</strong>
              <span>Your saved sheet is ready. Continue with this phone number.</span>
            </div>
          ) : (
            <label>Name<input value={auth.name} onChange={(e) => setAuth({ ...auth, name: e.target.value })} /></label>
          )}
          <button className="primary" type="submit"><Check size={18} /> Continue</button>
        </form>
      </main>
    );
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand"><BarChart3 size={26} /> Response Sheet</div>
        <div className="profile">
          <strong>{currentUser.name}</strong>
          <span>{currentUser.phone}</span>
        </div>
        <button onClick={() => setMode("dashboard")} className={mode === "dashboard" ? "active" : ""}>Dashboard</button>
        <button onClick={() => setMode("fill")} className={mode === "fill" ? "active" : ""}>Fill rows</button>
        <button onClick={exportWorkbook}><Download size={17} /> Export Excel</button>
        <button onClick={() => updateStore((prev) => ({ ...prev, currentPhone: "" }))}><LogOut size={17} /> Sign out</button>
      </aside>

      {mode === "dashboard" ? (
        <section className="workspace">
          <header className="topbar">
            <div><h1>Dashboard</h1><p>{answeredCount} of {rows.length} rows saved</p></div>
            <button className="primary" onClick={() => setMode("fill")}><ArrowRight size={18} /> Continue filling</button>
          </header>
          <div className="save-note">
            <strong>Saved account sheet</strong>
            <span>This browser keeps {currentUser.name}'s work under phone number {currentUser.phone}. Close and reopen, then enter the same phone number to continue.</span>
          </div>
          <div className="stats">
            <div><span>Progress</span><strong>{progress}%</strong></div>
            <div><span>Completed rows</span><strong>{answeredCount}</strong></div>
            <div><span>Remaining</span><strong>{rows.length - answeredCount}</strong></div>
          </div>
          <div className="progress"><span style={{ width: `${progress}%` }} /></div>
          <div className="search"><Search size={18} /><input placeholder="Search story, source, or number" value={query} onChange={(e) => setQuery(e.target.value)} /></div>
          <div className="row-list">
            {filteredRows.map((row) => (
              <button key={row.excelRow} onClick={() => { setActiveIndex(rows.findIndex((item) => item.excelRow === row.excelRow)); setMode("fill"); }}>
                <span>Row {row.excelRow}</span>
                <strong>{row.story}</strong>
                <em>{answers[row.excelRow]?.complete ? "Saved" : "Open"}</em>
              </button>
            ))}
          </div>
        </section>
      ) : (
        <section className="workspace fill-view">
          <header className="topbar">
            <button onClick={() => setActiveIndex(Math.max(0, activeIndex - 1))}><ArrowLeft size={18} /> Previous</button>
            <div><h1>Row {activeRow.excelRow}</h1><p>{activeIndex + 1} of {rows.length}</p></div>
            <button className="primary" onClick={() => setActiveIndex(Math.min(rows.length - 1, activeIndex + 1))}>Next <ArrowRight size={18} /></button>
          </header>
          <div className="story-panel">
            <span>{activeRow.source} / Story {activeRow.storyNo}</span>
            <h2>{activeRow.story}</h2>
          </div>
          <div className="outputs">
            <ModelOutput
              name="Claude"
              text={activeRow.claude}
              answer={answers[activeRow.excelRow]}
              onValidityChange={(key, value) =>
                saveAnswer(activeRow.excelRow, { requirementValidity: { [key]: value } })
              }
            />
            <ModelOutput
              name="GPT5"
              text={activeRow.gpt}
              answer={answers[activeRow.excelRow]}
              onValidityChange={(key, value) =>
                saveAnswer(activeRow.excelRow, { requirementValidity: { [key]: value } })
              }
            />
            <ModelOutput
              name="Gemini"
              text={activeRow.gemini}
              answer={answers[activeRow.excelRow]}
              onValidityChange={(key, value) =>
                saveAnswer(activeRow.excelRow, { requirementValidity: { [key]: value } })
              }
            />
            <ModelOutput
              name="DeepSeek"
              text={activeRow.deepseek}
              answer={answers[activeRow.excelRow]}
              onValidityChange={(key, value) =>
                saveAnswer(activeRow.excelRow, { requirementValidity: { [key]: value } })
              }
            />
          </div>
          <form className="answer-grid">
            {MODEL_NAMES.map((model) => (
              <fieldset key={model} className="model-score-card">
                <legend>{model}</legend>
                {ANSWER_FIELDS.filter((field) => field.model === model).map((field) => {
                  const value = answers[activeRow.excelRow]?.[field.key] || "";
                  return (
                    <div key={field.key} className="score-control">
                      <span>{field.label}</span>
                      {field.type === "yesno" && (
                        <ChoiceButtons
                          value={value}
                          options={["Yes", "No"]}
                          onChange={(nextValue) => saveAnswer(activeRow.excelRow, { [field.key]: nextValue })}
                        />
                      )}
                      {field.type === "scale" && (
                        <ChoiceButtons
                          value={value}
                          options={["1", "2", "3", "4", "5"]}
                          onChange={(nextValue) => saveAnswer(activeRow.excelRow, { [field.key]: nextValue })}
                        />
                      )}
                      {field.type === "number" && (
                        <input
                          type="number"
                          min="0"
                          value={value}
                          onChange={(e) => saveAnswer(activeRow.excelRow, { [field.key]: e.target.value })}
                        />
                      )}
                    </div>
                  );
                })}
              </fieldset>
            ))}
          </form>
          <div className="bottom-actions">
            <button onClick={() => saveAnswer(activeRow.excelRow, { complete: true })}><Save size={18} /> Save row</button>
            <button className="primary" onClick={() => setActiveIndex(Math.min(rows.length - 1, activeIndex + 1))}>Save and next <ArrowRight size={18} /></button>
          </div>
        </section>
      )}
    </main>
  );
}

createRoot(document.getElementById("root")).render(<App />);
