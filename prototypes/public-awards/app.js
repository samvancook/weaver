const state = { records: [], search: "", year: "", result: "", organization: "" };
const $ = (selector) => document.querySelector(selector);
const normalize = (value) => String(value ?? "").normalize("NFKD").toLocaleLowerCase().trim();
const safeUrl = (value) => {
  try { const url = new URL(value); return url.protocol === "https:" ? url.href : null; }
  catch { return null; }
};
const required = ["id","year","book","author","product_url","award","result","organization","source_url"];
function validate(record) {
  if (!record || typeof record !== "object" || required.some((key) => record[key] === undefined || record[key] === null || record[key] === "")) return false;
  if (!Number.isInteger(record.year) || record.year < 1900 || record.year > 2100) return false;
  if (!safeUrl(record.source_url) || !safeUrl(record.product_url)) return false;
  return ["id","book","author","award","result","organization"].every((key) => typeof record[key] === "string");
}
function uniqueBy(key) { return [...new Set(state.records.map((entry) => String(entry[key])))].sort((a,b)=> key === "year" ? Number(b)-Number(a) : a.localeCompare(b)); }
function options(node, items) {
  for (const value of items) { const option = document.createElement("option"); option.value = value; option.textContent = value; node.append(option); }
}
function link(label,url,kind) {
  const a = document.createElement("a");
  a.textContent = label;
  a.href = safeUrl(url) || "#";
  if (kind === "source") { a.target = "_blank"; a.rel = "noopener noreferrer"; a.setAttribute("aria-label", label + " (opens new tab)"); }
  return a;
}
function text(tag, cls, value) {
  const node = document.createElement(tag);
  if (cls) node.className = cls;
  node.textContent = value;
  return node;
}
function renderCard(record) {
  const article = document.createElement("article");
  article.className = "card";
  article.append(text("div", "card-top", String(record.year) + " · " + record.organization));
  const heading = text("h3", "card-title", record.book);
  article.append(heading);
  article.append(text("p", "author", "by " + record.author));
  article.append(text("span", "result-chip", record.result));
  article.append(text("p", "award", record.award));
  if (record.category) article.append(text("p", "category", record.category));
  const footer = document.createElement("div");
  footer.className = "card-links";
  footer.append(link("View the book", record.product_url, "book"), link("Official award source ↗", record.source_url, "source"));
  article.append(footer);
  return article;
}
function filtered() {
  return state.records.filter((record) => {
    if (state.year && String(record.year) !== state.year) return false;
    if (state.result && record.result !== state.result) return false;
    if (state.organization && record.organization !== state.organization) return false;
    const haystack = normalize([record.book, record.author, record.award, record.category, record.organization, record.result].join(" "));
    return !state.search || haystack.includes(normalize(state.search));
  });
}
function render() {
  const matches = filtered();
  const cards = $("#cards"); cards.replaceChildren(...matches.map(renderCard));
  $("#empty").hidden = matches.length > 0;
  $("#result-count").textContent = matches.length + " of " + state.records.length + " recognitions";
}
function syncUrl() {
  const url = new URL(window.location.href);
  for (const key of ["search","year","result","organization"]) {
    if (state[key]) url.searchParams.set(key, state[key]); else url.searchParams.delete(key);
  }
  window.history.replaceState(null, "", url);
}
function setValue(key,value) { state[key] = value; syncUrl(); render(); }
function reset() {
  for (const key of ["search","year","result","organization"]) { state[key] = ""; $("#" + key).value = ""; }
  syncUrl(); render();
}
async function boot() {
  try {
    const response = await fetch("./awards.json", { cache: "no-store" });
    if (!response.ok) throw new Error("Could not load awards data.");
    const data = await response.json();
    if (data.schema_version !== "1.0" || data.status !== "review_only" || !Array.isArray(data.records)) throw new Error("Unexpected data format.");
    if (!data.records.every(validate) || new Set(data.records.map(item => item.id)).size !== data.records.length) throw new Error("Invalid or duplicate awards records.");
    state.records = data.records.slice().sort((a,b)=>b.year-a.year || a.book.localeCompare(b.book));
    for (const key of ["year","result","organization"]) options($("#" + key), uniqueBy(key));
    const url = new URL(window.location.href);
    for (const key of ["search","year","result","organization"]) {
      const value = url.searchParams.get(key) || "";
      if (key === "search" || uniqueBy(key).includes(value)) { state[key] = value; $("#" + key).value = value; }
    }
    for (const key of ["search","year","result","organization"]) {
      $("#" + key).addEventListener(key === "search" ? "input" : "change", (event) => setValue(key,event.target.value));
    }
    $("#reset").addEventListener("click",reset);
    $("#copy-link").addEventListener("click", async () => {
      syncUrl();
      const button = $("#copy-link"); const before=button.textContent;
      try {
        if (!navigator.clipboard) throw new Error("Clipboard not available");
        await navigator.clipboard.writeText(window.location.href); button.textContent="Link copied";
      } catch { button.textContent="Copy address from browser"; }
      window.setTimeout(()=>{button.textContent=before;},2300);
    });
    render();
  } catch (error) {
    $("#error").hidden = false;
    $("#error").textContent = error?.message || "Unable to load awards.";
    $("#result-count").textContent = "Awards unavailable";
  }
}
boot();
