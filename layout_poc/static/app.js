const layoutGrid = document.getElementById("layoutGrid");
const layoutMenu = document.getElementById("layoutMenu");
const layoutDropdownBtn = document.getElementById("layoutDropdownBtn");
const selectedPreview = document.getElementById("selectedPreview");
const dropdownSelectedName = document.getElementById("dropdownSelectedName");
const dropdownSelectedMeta = document.getElementById("dropdownSelectedMeta");
const templatePath = document.getElementById("templatePath");
const previewMode = document.getElementById("previewMode");
const selectedName = document.getElementById("selectedName");
const selectedMeta = document.getElementById("selectedMeta");
const renderBtn = document.getElementById("renderBtn");
const refreshBtn = document.getElementById("refreshBtn");
const statusText = document.getElementById("statusText");

const titleInput = document.getElementById("titleInput");
const bulletsInput = document.getElementById("bulletsInput");
const chartTitleInput = document.getElementById("chartTitleInput");
const categoriesInput = document.getElementById("categoriesInput");
const valuesInput = document.getElementById("valuesInput");
const chartToggle = document.getElementById("chartToggle");
const paletteGrid = document.getElementById("paletteGrid");
const paletteMeta = document.getElementById("paletteMeta");
const selectAllColorsBtn = document.getElementById("selectAllColorsBtn");
const clearColorsBtn = document.getElementById("clearColorsBtn");

let layouts = [];
let selectedLayout = null;
let palette = [];
let selectedColors = new Set();

function setStatus(message) {
  statusText.textContent = message || "";
}

function parseList(value) {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function parseValues(value) {
  return parseList(value).map((item) => Number(item));
}

function selectLayout(layout) {
  selectedLayout = layout;
  selectedName.textContent = `${layout.index}: ${layout.name}`;
  selectedMeta.textContent = `${layout.placeholder_count} placeholders`;
  dropdownSelectedName.textContent = `${layout.index}: ${layout.name}`;
  dropdownSelectedMeta.textContent = `${layout.placeholder_count} placeholders`;
  selectedPreview.src = layout.preview;
  renderBtn.disabled = false;
  layoutMenu.hidden = true;

  document.querySelectorAll(".layout-card").forEach((card) => {
    card.classList.toggle(
      "selected",
      Number(card.dataset.index) === layout.index
    );
  });
}

function renderLayoutCard(layout) {
  const card = document.createElement("button");
  card.type = "button";
  card.className = "layout-card";
  card.dataset.index = String(layout.index);

  const image = document.createElement("img");
  image.src = layout.preview;
  image.alt = `${layout.name} preview`;

  const title = document.createElement("strong");
  title.textContent = `${layout.index}: ${layout.name}`;

  const meta = document.createElement("span");
  meta.textContent = `${layout.placeholder_count} placeholders`;

  card.append(image, title, meta);
  card.addEventListener("click", () => selectLayout(layout));

  return card;
}

function selectedColorList() {
  return palette
    .map((color) => color.hex)
    .filter((color) => selectedColors.has(color));
}

function renderPalette() {
  paletteGrid.replaceChildren();

  if (!palette.length) {
    paletteMeta.textContent = "No explicit RGB colors found.";
    return;
  }

  paletteMeta.textContent = `${selectedColors.size} of ${palette.length} selected`;

  const swatches = palette.map((color) => {
    const swatch = document.createElement("button");
    swatch.type = "button";
    swatch.className = "swatch";
    swatch.style.backgroundColor = color.hex;
    swatch.title = `${color.hex} (${color.count} uses)`;
    swatch.classList.toggle("selected", selectedColors.has(color.hex));

    swatch.addEventListener("click", () => {
      if (selectedColors.has(color.hex)) {
        selectedColors.delete(color.hex);
      } else {
        selectedColors.add(color.hex);
      }

      renderPalette();
    });

    return swatch;
  });

  paletteGrid.replaceChildren(...swatches);
}

async function loadLayouts() {
  setStatus("Loading layouts...");
  renderBtn.disabled = true;
  selectedLayout = null;
  selectedName.textContent = "None";
  selectedMeta.textContent = "Pick a layout to render a sample slide.";
  dropdownSelectedName.textContent = "Select a layout";
  dropdownSelectedMeta.textContent = "Open layout gallery";
  selectedPreview.removeAttribute("src");
  layoutMenu.hidden = true;
  layoutGrid.replaceChildren();

  const response = await fetch("/api/layouts");

  if (!response.ok) {
    throw new Error(await response.text());
  }

  const payload = await response.json();
  layouts = payload.layouts || [];
  palette = payload.palette || [];
  selectedColors = new Set(palette.map((color) => color.hex));
  templatePath.textContent = payload.template || "Template loaded";
  if (payload.preview_mode === "powerpoint") {
    previewMode.textContent = "Rendered previews via Microsoft PowerPoint.";
  } else if (payload.preview_mode === "libreoffice") {
    previewMode.textContent = `Rendered previews via ${payload.renderer}.`;
  } else {
    previewMode.textContent =
      payload.renderer_error
        ? `Schematic previews. ${payload.renderer_error}`
        : "Schematic previews. Install pywin32/Office on Windows or LibreOffice for rendered previews.";
  }

  layoutGrid.replaceChildren(...layouts.map(renderLayoutCard));
  renderPalette();

  if (layouts.length) {
    selectLayout(layouts[0]);
  }

  setStatus(`Loaded ${layouts.length} layouts.`);
}

async function renderSlide() {
  if (!selectedLayout) {
    return;
  }

  const categories = parseList(categoriesInput.value);
  const values = parseValues(valuesInput.value);

  if (categories.length !== values.length || values.some(Number.isNaN)) {
    setStatus("Categories and numeric values must match.");
    return;
  }

  const body = {
    layout_index: selectedLayout.index,
    title: titleInput.value.trim() || "Sample Insights",
    bullets: bulletsInput.value
      .split("\n")
      .map((item) => item.trim())
      .filter(Boolean),
    chart_title: chartTitleInput.value.trim() || "Sample Chart",
    categories,
    values,
    include_chart: chartToggle.checked,
    selected_colors: selectedColorList(),
  };

  setStatus("Rendering PPTX...");
  renderBtn.disabled = true;

  try {
    const response = await fetch("/api/render-slide", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      throw new Error(await response.text());
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "layout_poc_render.pptx";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setStatus("Downloaded layout_poc_render.pptx.");
  } catch (error) {
    setStatus(`Render failed: ${error.message}`);
  } finally {
    renderBtn.disabled = false;
  }
}

refreshBtn.addEventListener("click", () => {
  loadLayouts().catch((error) => setStatus(`Load failed: ${error.message}`));
});

selectAllColorsBtn.addEventListener("click", () => {
  selectedColors = new Set(palette.map((color) => color.hex));
  renderPalette();
});

clearColorsBtn.addEventListener("click", () => {
  selectedColors = new Set();
  renderPalette();
});

layoutDropdownBtn.addEventListener("click", () => {
  layoutMenu.hidden = !layoutMenu.hidden;
});

document.addEventListener("click", (event) => {
  const target = event.target;

  if (
    !layoutMenu.hidden &&
    !layoutMenu.contains(target) &&
    !layoutDropdownBtn.contains(target)
  ) {
    layoutMenu.hidden = true;
  }
});

renderBtn.addEventListener("click", renderSlide);

loadLayouts().catch((error) => setStatus(`Load failed: ${error.message}`));
