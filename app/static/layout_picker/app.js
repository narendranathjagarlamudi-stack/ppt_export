const layoutGrid = document.getElementById("layoutGrid");
const layoutsPanel = document.getElementById("layoutsPanel");
const layoutMenu = document.getElementById("layoutMenu");
const layoutDropdownBtn = document.getElementById("layoutDropdownBtn");
const selectedPreview = document.getElementById("selectedPreview");
const dropdownSelectedName = document.getElementById("dropdownSelectedName");
const dropdownSelectedMeta = document.getElementById("dropdownSelectedMeta");
const templatePath = document.getElementById("templatePath");
const previewMode = document.getElementById("previewMode");
const selectedCard = document.getElementById("selectedCard");
const selectedName = document.getElementById("selectedName");
const selectedMeta = document.getElementById("selectedMeta");
const exportBtn = document.getElementById("exportBtn");
const previewBtn = document.getElementById("previewBtn");
const applyLayoutToSlidesBtn = document.getElementById("applyLayoutToSlidesBtn");
const refreshBtn = document.getElementById("refreshBtn");
const refreshCacheBtn = document.getElementById("refreshCacheBtn");
const statusText = document.getElementById("statusText");

const chartToggle = document.getElementById("chartToggle");
const paletteGrid = document.getElementById("paletteGrid");
const paletteMeta = document.getElementById("paletteMeta");
const defaultColorsBtn = document.getElementById("defaultColorsBtn");
const selectAllColorsBtn = document.getElementById("selectAllColorsBtn");
const clearColorsBtn = document.getElementById("clearColorsBtn");
const dcsIdInput = document.getElementById("dcsIdInput");
const dihIdsInput = document.getElementById("dihIdsInput");
const exportItemsInput = document.getElementById("exportItemsInput");
const filenameInput = document.getElementById("filenameInput");
const pptConfigInput = document.getElementById("pptConfigInput");
const configUploadInput = document.getElementById("configUploadInput");
const templateUploadInput = document.getElementById("templateUploadInput");
const fallbackModeNote = document.getElementById("fallbackModeNote");
const exportSummary = document.getElementById("exportSummary");
const slidePreviewMeta = document.getElementById("slidePreviewMeta");
const slidePreviewList = document.getElementById("slidePreviewList");

const layoutsEndpoint = "/export/layouts";

let layouts = [];
let selectedLayout = null;
let palette = [];
let selectedColors = new Set();
let useTemplateDefaultColors = true;
let previewSlides = [];
let skippedMessages = [];
let selectedSlideIndices = new Set();
let slideSelections = {};
let templateLayoutsAvailable = false;
let uploadedTemplateId = null;

function setStatus(message) {
  statusText.textContent = message || "";
}

function setBusy(isBusy, message) {
  if (message) {
    setStatus(message);
  }

  previewBtn.disabled = isBusy || (templateLayoutsAvailable ? !selectedLayout : false);
  exportBtn.disabled = isBusy || !selectedSlideIndices.size;
}

function selectLayout(layout) {
  selectedLayout = layout;
  selectedName.textContent = `${layout.index}: ${layout.name}`;
  selectedMeta.textContent = layoutSummary(layout);
  dropdownSelectedName.textContent = `${layout.index}: ${layout.name}`;
  dropdownSelectedMeta.textContent = layoutSummary(layout);
  selectedPreview.src = layout.preview;
  previewBtn.disabled = false;
  exportBtn.disabled = !previewSlides.length;
  layoutMenu.hidden = true;

  document.querySelectorAll(".layout-card").forEach((card) => {
    card.classList.toggle(
      "selected",
      Number(card.dataset.index) === layout.index
    );
  });

  renderSlidePreviewList();
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
  meta.textContent = layoutSummary(layout);

  card.append(image, title, meta);
  card.addEventListener("click", () => selectLayout(layout));

  return card;
}

function defaultContentLayout() {
  return (
    layouts.find((layout) => layout.name === "Title & Body")
    || layouts.find((layout) => layout.name.includes("Title & Body"))
    || layouts.find((layout) => layout.name === "Custom Layout")
    || layouts.find((layout) => layout.name === "Blank")
    || layouts[0]
  );
}

function selectedColorList() {
  if (useTemplateDefaultColors) {
    return [];
  }

  return palette
    .map((color) => color.hex)
    .filter((color) => selectedColors.has(color));
}

function layoutByIndex(index) {
  return layouts.find((layout) => layout.index === Number(index));
}

function layoutSummary(layout) {
  if (!layout) {
    return "No layout selected";
  }

  const contentText = layout.has_content_placeholder
    ? "has content area"
    : "no content area";

  return `${layout.placeholder_count} placeholders, ${contentText}`;
}

function resetPreviewState() {
  previewSlides = [];
  skippedMessages = [];
  selectedSlideIndices = new Set();
  slideSelections = {};
  slidePreviewMeta.textContent = "Preview slide count before export.";
  exportSummary.textContent = "Preview before export.";
  slidePreviewList.replaceChildren();
  applyLayoutToSlidesBtn.disabled = true;
  exportBtn.disabled = true;
}

function skippedMessageSummary() {
  if (!skippedMessages.length) {
    return "";
  }

  return skippedMessages
    .map((item) => `DIH ${item.dih_id || "-"}${item.message_id ? ` / ${item.message_id}` : ""}`)
    .join(", ");
}

function enterFallbackConfigMode(message) {
  layouts = [];
  palette = [];
  selectedColors = new Set();
  useTemplateDefaultColors = true;
  templateLayoutsAvailable = false;
  selectedLayout = null;

  selectedName.textContent = "Fallback config mode";
  selectedMeta.textContent = "No template layouts are available. Export will use fallback PPT config JSON.";
  dropdownSelectedName.textContent = "No template loaded";
  dropdownSelectedMeta.textContent = "Using fallback PPT config";
  selectedPreview.removeAttribute("src");
  templatePath.textContent = "Template not available";
  previewMode.textContent = message || "Fallback config mode.";
  layoutsPanel.hidden = true;
  selectedCard.hidden = true;
  layoutDropdownBtn.disabled = true;
  fallbackModeNote.hidden = false;
  layoutGrid.replaceChildren();
  layoutMenu.hidden = true;
  previewBtn.disabled = false;
  applyLayoutToSlidesBtn.hidden = true;
  renderPalette();
}

function renderPalette() {
  paletteGrid.replaceChildren();

  if (!palette.length) {
    paletteMeta.textContent = "No explicit RGB colors found.";
    return;
  }

  if (useTemplateDefaultColors) {
    paletteMeta.textContent = "Using template default chart colors.";
  } else {
    paletteMeta.textContent = `${selectedColors.size} of ${palette.length} selected`;
  }

  const swatches = palette.map((color) => {
    const swatch = document.createElement("button");
    swatch.type = "button";
    swatch.className = "swatch";
    swatch.style.backgroundColor = color.hex;
    swatch.title = `${color.hex} (${color.count} uses)`;
    swatch.classList.toggle(
      "selected",
      !useTemplateDefaultColors && selectedColors.has(color.hex)
    );

    swatch.addEventListener("click", () => {
      useTemplateDefaultColors = false;

      if (selectedColors.has(color.hex)) {
        selectedColors.delete(color.hex);
      } else {
        selectedColors.add(color.hex);
      }

      renderPalette();
      updateExportSummary();
    });

    return swatch;
  });

  paletteGrid.replaceChildren(...swatches);
}

async function loadLayouts() {
  setStatus("Loading layouts...");
  exportBtn.disabled = true;
  previewBtn.disabled = true;
  resetPreviewState();
  selectedLayout = null;
  templateLayoutsAvailable = false;
  layoutsPanel.hidden = false;
  selectedCard.hidden = false;
  layoutDropdownBtn.disabled = false;
  applyLayoutToSlidesBtn.hidden = false;
  fallbackModeNote.hidden = true;
  selectedName.textContent = "None";
  selectedMeta.textContent = "Pick a layout for the exported slides.";
  dropdownSelectedName.textContent = "Select a layout";
  dropdownSelectedMeta.textContent = "Open layout gallery";
  selectedPreview.removeAttribute("src");
  layoutMenu.hidden = true;
  layoutGrid.replaceChildren();

  const params = new URLSearchParams();

  if (uploadedTemplateId) {
    params.set("template_id", uploadedTemplateId);
  }

  const response = await fetch(
    params.toString()
      ? `${layoutsEndpoint}?${params.toString()}`
      : layoutsEndpoint
  );

  if (!response.ok) {
    enterFallbackConfigMode(
      "Template layouts could not be loaded. You can still export with fallback PPT config JSON."
    );
    setStatus("Template layout load failed. Fallback config export is available.");
    return;
  }

  const payload = await response.json();
  layouts = payload.layouts || [];
  templateLayoutsAvailable = layouts.length > 0;
  palette = payload.palette || [];
  selectedColors = new Set(palette.map((color) => color.hex));
  useTemplateDefaultColors = true;
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
    selectLayout(defaultContentLayout());
  }

  setStatus(`Loaded ${layouts.length} layouts.`);
}

function parseIds(value) {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item.length)
    .map((item) => Number(item))
    .filter((item) => Number.isInteger(item));
}

function parsePptConfig() {
  const rawConfig = pptConfigInput.value.trim();

  if (!rawConfig) {
    return null;
  }

  try {
    const parsed = JSON.parse(rawConfig);

    if (!parsed || Array.isArray(parsed) || typeof parsed !== "object") {
      throw new Error("Config must be a JSON object.");
    }

    return parsed;
  } catch (error) {
    throw new Error(`Fallback PPT config JSON is invalid: ${error.message}`);
  }
}

function parseExportItems() {
  const rawItems = exportItemsInput.value.trim();

  if (!rawItems) {
    return null;
  }

  try {
    const parsed = JSON.parse(rawItems);

    if (!Array.isArray(parsed)) {
      throw new Error("Export items must be a JSON array.");
    }

    parsed.forEach((item) => {
      if (
        !Number.isInteger(Number(item.dcs_id))
        || !Array.isArray(item.dih_ids)
        || !item.dih_ids.length
      ) {
        throw new Error("Each item must include dcs_id and dih_ids.");
      }
    });

    return parsed.map((item) => ({
      dcs_id: Number(item.dcs_id),
      dih_ids: item.dih_ids
        .map((id) => Number(id))
        .filter((id) => Number.isInteger(id)),
    }));
  } catch (error) {
    throw new Error(`Export Items JSON is invalid: ${error.message}`);
  }
}

function exportRequestTarget() {
  const exportItems = parseExportItems();

  if (exportItems && exportItems.length) {
    return {
      dcsId: exportItems[0].dcs_id,
      body: {
        export_items: exportItems,
      },
    };
  }

  const dcsId = Number(dcsIdInput.value.trim());
  const dihIds = parseIds(dihIdsInput.value);

  if (!Number.isInteger(dcsId) || !dihIds.length) {
    throw new Error("Enter a valid DCS ID and at least one message/DIH ID, or provide Export Items JSON.");
  }

  return {
    dcsId,
    body: {
      dih_ids: dihIds,
    },
  };
}

function selectedPreviewSlides() {
  return previewSlides.filter((slide) => selectedSlideIndices.has(slide.index));
}

function updateExportSummary() {
  const selectedCount = selectedPreviewSlides().length;
  const skippedCount = skippedMessages.length;
  const colorText = useTemplateDefaultColors
    ? "template default colors"
    : `${selectedColorList().length} custom colors`;
  const templateText = templateLayoutsAvailable
    ? "template layouts"
    : "fallback config";

  exportSummary.textContent = `${selectedCount} slide(s) selected for export, ${skippedCount} message(s) skipped, using ${templateText} and ${colorText}.`;
  exportBtn.disabled = selectedCount === 0;
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      const value = String(reader.result || "");
      resolve(value.split(",").pop());
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function renderSlidePreviewList() {
  slidePreviewList.replaceChildren();

  if (!previewSlides.length) {
    applyLayoutToSlidesBtn.disabled = true;
    exportBtn.disabled = true;

    if (skippedMessages.length) {
      const skippedRow = document.createElement("div");
      skippedRow.className = "skipped-row";
      skippedRow.textContent = `Skipped messages with no chart: ${skippedMessageSummary()}`;
      slidePreviewList.replaceChildren(skippedRow);
    }

    return;
  }

  applyLayoutToSlidesBtn.disabled = !selectedLayout || !templateLayoutsAvailable;
  exportBtn.disabled = false;

  const rows = previewSlides.map((slide) => {
    const row = document.createElement("div");
    row.className = "slide-row";
    row.classList.toggle("disabled", !selectedSlideIndices.has(slide.index));

    const includeRow = document.createElement("label");
    includeRow.className = "slide-select-row";

    const includeInput = document.createElement("input");
    includeInput.type = "checkbox";
    includeInput.checked = selectedSlideIndices.has(slide.index);
    includeInput.addEventListener("change", () => {
      if (includeInput.checked) {
        selectedSlideIndices.add(slide.index);
      } else {
        selectedSlideIndices.delete(slide.index);
      }

      renderSlidePreviewList();
      updateExportSummary();
    });

    includeRow.append(includeInput, "Include in export");

    const details = document.createElement("div");
    details.className = "slide-row-details";

    const title = document.createElement("strong");
    title.textContent = `Slide ${slide.index + 1}: ${slide.title || "Key Insights"}`;

    const meta = document.createElement("span");
    const chartText = slide.has_chart ? "chart" : "no chart";
    const eventMeta = slide.event_metadata || {};
    const eventText = eventMeta.content_index !== undefined
      ? ` - content ${eventMeta.content_index}`
      : "";
    meta.textContent = `DCS ${slide.dcs_id || "-"} - DIH ${slide.dih_id || "-"} - ${slide.bullet_count} bullets - ${chartText}${eventText}`;

    details.append(title, meta);

    let select = null;
    const key = String(slide.index);

    if (templateLayoutsAvailable) {
      select = document.createElement("select");
      select.className = "slide-layout-select";

      layouts.forEach((layout) => {
        const option = document.createElement("option");
        option.value = String(layout.index);
        option.textContent = `${layout.index}: ${layout.name}`;
        select.append(option);
      });

      if (!Object.prototype.hasOwnProperty.call(slideSelections, key)) {
        slideSelections[key] = selectedLayout
          ? selectedLayout.index
          : defaultContentLayout().index;
      }

      select.value = String(slideSelections[key]);
      select.addEventListener("change", () => {
        slideSelections[key] = Number(select.value);
        renderSlidePreviewList();
      });
    } else {
      select = document.createElement("p");
      select.className = "fallback-layout-note";
      select.textContent = "Fallback PPT config layout will be used.";
    }

    row.append(includeRow, details, select);
    return row;
  });

  if (skippedMessages.length) {
    const skippedRow = document.createElement("div");
    skippedRow.className = "skipped-row";
    skippedRow.textContent = `Skipped messages with no chart: ${skippedMessageSummary()}`;
    rows.push(skippedRow);
  }

  slidePreviewList.replaceChildren(...rows);
}

async function previewExport() {
  let target = null;

  try {
    target = exportRequestTarget();
  } catch (error) {
    setStatus(error.message);
    return false;
  }

  setBusy(true, "Building slide preview...");

  try {
    const response = await fetch(`/export/preview/${target.dcsId}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(target.body),
    });

    if (!response.ok) {
      throw new Error(await response.text());
    }

    const payload = await response.json();
    previewSlides = payload.slides || [];
    skippedMessages = payload.skipped_messages || [];
    selectedSlideIndices = new Set(previewSlides.map((slide) => slide.index));
    slideSelections = {};

    const defaultIndex = templateLayoutsAvailable
      ? (
        selectedLayout
          ? selectedLayout.index
          : defaultContentLayout().index
      )
      : null;

    previewSlides.forEach((slide) => {
      if (defaultIndex !== null) {
        slideSelections[String(slide.index)] = defaultIndex;
      }
    });

    const skippedText = skippedMessages.length
      ? ` ${skippedMessages.length} message(s) with no chart will be skipped.`
      : "";

    slidePreviewMeta.textContent = `${payload.slide_count || previewSlides.length} slides will be generated.${skippedText}`;
    renderSlidePreviewList();
    updateExportSummary();

    if (!previewSlides.length) {
      setStatus("No chart slides found. Selected messages will be skipped.");
      return false;
    }

    setStatus(
      skippedMessages.length
        ? "Preview ready. Some messages with no chart will be skipped."
        : "Preview ready. Choose a layout per slide, then export."
    );
    return true;
  } catch (error) {
    resetPreviewState();
    setStatus(`Preview failed: ${error.message}`);
    return false;
  } finally {
    previewBtn.disabled = templateLayoutsAvailable
      ? !selectedLayout
      : false;
  }
}

async function exportMessages() {
  if (templateLayoutsAvailable && !selectedLayout) {
    return;
  }

  let target = null;

  try {
    target = exportRequestTarget();
  } catch (error) {
    setStatus(error.message);
    return;
  }

  let pptConfig = null;

  try {
    pptConfig = parsePptConfig();
  } catch (error) {
    setStatus(error.message);
    return;
  }

  if (!previewSlides.length) {
    const previewOk = await previewExport();

    if (!previewOk) {
      return;
    }
  }

  if (!selectedSlideIndices.size) {
    setStatus("Select at least one chart/slide to export.");
    return;
  }

  const slideLayouts = {};

  if (templateLayoutsAvailable) {
    selectedPreviewSlides().forEach((slide) => {
      const key = String(slide.index);
      slideLayouts[key] = Object.prototype.hasOwnProperty.call(slideSelections, key)
        ? slideSelections[key]
        : selectedLayout.index;
    });

    slideLayouts.default = selectedLayout.index;
  }

  const body = {
    ...target.body,
    include_charts: chartToggle.checked,
    selected_colors: selectedColorList(),
    selected_slide_indices: Array.from(selectedSlideIndices),
  };

  if (templateLayoutsAvailable) {
    body.slide_layouts = slideLayouts;
  }

  if (pptConfig) {
    body.ppt_config = pptConfig;
  }

  if (uploadedTemplateId) {
    body.template_id = uploadedTemplateId;
  }

  if (filenameInput.value.trim()) {
    body.filename = filenameInput.value.trim();
  }

  setBusy(true, "Exporting selected slides...");

  try {
    const response = await fetch(`/export/ppt/${target.dcsId}`, {
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
    link.download = filenameInput.value.trim() || "analysis_export.pptx";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setStatus("Downloaded PPT export.");
  } catch (error) {
    setStatus(`Export failed: ${error.message}`);
  } finally {
    updateExportSummary();
  }
}

refreshBtn.addEventListener("click", () => {
  loadLayouts().catch((error) => setStatus(`Load failed: ${error.message}`));
});

refreshCacheBtn.addEventListener("click", async () => {
  const params = new URLSearchParams();

  params.set("refresh_cache", "true");

  if (uploadedTemplateId) {
    params.set("template_id", uploadedTemplateId);
  }

  try {
    setBusy(true, "Refreshing template thumbnails...");
    const response = await fetch(`${layoutsEndpoint}?${params.toString()}`);

    if (!response.ok) {
      throw new Error(await response.text());
    }

    await loadLayouts();
  } catch (error) {
    setStatus(`Thumbnail refresh failed: ${error.message}`);
  } finally {
    setBusy(false);
  }
});

templateUploadInput.addEventListener("change", async () => {
  const file = templateUploadInput.files && templateUploadInput.files[0];

  if (!file) {
    return;
  }

  try {
    setBusy(true, "Uploading template...");
    const contentBase64 = await fileToBase64(file);
    const response = await fetch("/export/templates", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        filename: file.name,
        content_base64: contentBase64,
      }),
    });

    if (!response.ok) {
      throw new Error(await response.text());
    }

    const payload = await response.json();
    uploadedTemplateId = payload.template_id;
    resetPreviewState();
    await loadLayouts();
    setStatus(`Uploaded template: ${file.name}`);
  } catch (error) {
    setStatus(`Template upload failed: ${error.message}`);
  } finally {
    setBusy(false);
  }
});

configUploadInput.addEventListener("change", async () => {
  const file = configUploadInput.files && configUploadInput.files[0];

  if (!file) {
    return;
  }

  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    pptConfigInput.value = JSON.stringify(parsed, null, 2);
    setStatus(`Loaded config JSON: ${file.name}`);
  } catch (error) {
    setStatus(`Config upload failed: ${error.message}`);
  }
});

defaultColorsBtn.addEventListener("click", () => {
  useTemplateDefaultColors = true;
  renderPalette();
  updateExportSummary();
});

selectAllColorsBtn.addEventListener("click", () => {
  useTemplateDefaultColors = false;
  selectedColors = new Set(palette.map((color) => color.hex));
  renderPalette();
  updateExportSummary();
});

clearColorsBtn.addEventListener("click", () => {
  useTemplateDefaultColors = false;
  selectedColors = new Set();
  renderPalette();
  updateExportSummary();
});

previewBtn.addEventListener("click", previewExport);

applyLayoutToSlidesBtn.addEventListener("click", () => {
  if (!templateLayoutsAvailable || !selectedLayout || !previewSlides.length) {
    return;
  }

  previewSlides.forEach((slide) => {
    slideSelections[String(slide.index)] = selectedLayout.index;
  });

  renderSlidePreviewList();
  setStatus("Selected layout applied to all previewed slides.");
});

dcsIdInput.addEventListener("input", resetPreviewState);
dihIdsInput.addEventListener("input", resetPreviewState);
exportItemsInput.addEventListener("input", resetPreviewState);
filenameInput.addEventListener("input", updateExportSummary);

layoutDropdownBtn.addEventListener("click", () => {
  if (!templateLayoutsAvailable) {
    return;
  }

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

exportBtn.addEventListener("click", exportMessages);

loadLayouts().catch((error) => setStatus(`Load failed: ${error.message}`));
