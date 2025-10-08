const NODE_WIDTH = 220;
const NODE_HEIGHT = 120;
const HORIZONTAL_GAP = 260;
const VERTICAL_GAP = 200;
const MIN_SCALE = 0.4;
const MAX_SCALE = 2.4;
const SCALE_STEP = 0.15;

const COLUMN_ALIASES = {
  name: ["employee name", "name", "full name"],
  tribe: ["tribe", "chapter", "department", "group"],
  squad: ["squad", "team", "unit"],
  expertise: ["expertise", "role", "function", "speciality", "specialty"],
  level: ["job level", "level", "grade"],
  manager: ["manager name", "manager", "reporting to", "reports to"],
};

const SAMPLE_DATA = [
  {
    name: "Aria Chen",
    tribe: "Product",
    squad: "Discovery",
    expertise: "Product Manager",
    level: "Senior",
    manager: "Daniel Rivera",
  },
  {
    name: "Jonas Patel",
    tribe: "Product",
    squad: "Discovery",
    expertise: "UX Researcher",
    level: "Mid",
    manager: "Aria Chen",
  },
  {
    name: "Lila Mensah",
    tribe: "Product",
    squad: "Discovery",
    expertise: "Product Designer",
    level: "Mid",
    manager: "Aria Chen",
  },
  {
    name: "Mateo García",
    tribe: "Product",
    squad: "Growth",
    expertise: "Product Manager",
    level: "Lead",
    manager: "Daniel Rivera",
  },
  {
    name: "Sofia Novak",
    tribe: "Engineering",
    squad: "Growth",
    expertise: "Frontend Engineer",
    level: "Mid",
    manager: "Mateo García",
  },
  {
    name: "Emily Stone",
    tribe: "Engineering",
    squad: "Growth",
    expertise: "Backend Engineer",
    level: "Senior",
    manager: "Mateo García",
  },
  {
    name: "Kai Nakamura",
    tribe: "Engineering",
    squad: "Platform",
    expertise: "DevOps Engineer",
    level: "Senior",
    manager: "Priya Desai",
  },
  {
    name: "Priya Desai",
    tribe: "Engineering",
    squad: "Platform",
    expertise: "Engineering Manager",
    level: "Lead",
    manager: "Daniel Rivera",
  },
  {
    name: "Daniel Rivera",
    tribe: "Executive",
    squad: "Leadership",
    expertise: "VP of Product & Engineering",
    level: "Executive",
    manager: "",
  },
  {
    name: "Noah Jensen",
    tribe: "People",
    squad: "Operations",
    expertise: "People Partner",
    level: "Senior",
    manager: "Daniel Rivera",
  },
  {
    name: "Amelia Fox",
    tribe: "People",
    squad: "Operations",
    expertise: "Recruiter",
    level: "Mid",
    manager: "Noah Jensen",
  },
];

const state = {
  rows: [],
  view: "tribe",
  collapsed: new Set(),
  search: "",
  scale: 1,
  layout: null,
};

const elements = {
  viewToggle: document.getElementById("view-toggle"),
  themeToggle: document.getElementById("theme-toggle"),
  fileInput: document.getElementById("file-input"),
  sheetUrl: document.getElementById("sheet-url"),
  loadSheet: document.getElementById("load-sheet"),
  loadSample: document.getElementById("load-sample"),
  searchInput: document.getElementById("search-input"),
  clearSearch: document.getElementById("clear-search"),
  chartContainer: document.getElementById("chart-container"),
  chartContent: document.getElementById("chart-content"),
  chartZoom: document.getElementById("chart-zoom"),
  chartLinks: document.getElementById("chart-links"),
  chartNodes: document.getElementById("chart-nodes"),
  minimap: document.getElementById("minimap"),
  emptyState: document.getElementById("empty-state"),
  statusBar: document.getElementById("status-bar"),
  statusMessage: document.getElementById("status-message"),
  fab: document.getElementById("fab"),
};

const themePreference = localStorage.getItem("org-nav-theme");
if (themePreference) {
  document.documentElement.setAttribute("data-theme", themePreference);
  updateThemeButton(themePreference);
}

elements.themeToggle.addEventListener("click", () => {
  const current = document.documentElement.getAttribute("data-theme") || "light";
  const next = current === "light" ? "dark" : "light";
  document.documentElement.setAttribute("data-theme", next);
  localStorage.setItem("org-nav-theme", next);
  updateThemeButton(next);
});

elements.viewToggle.addEventListener("change", () => {
  state.view = elements.viewToggle.checked ? "manager" : "tribe";
  state.collapsed.clear();
  render();
});

elements.fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    updateStatus(`Loading ${file.name}…`);
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) throw new Error("No sheets detected in the workbook");
    const worksheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    state.rows = mapRows(rawRows);
    state.collapsed.clear();
    state.search = "";
    elements.searchInput.value = "";
    render();
    updateStatus(`Loaded ${state.rows.length} rows from ${sheetName}`);
  } catch (error) {
    console.error(error);
    updateStatus(`Failed to load spreadsheet: ${error.message}`, true);
  } finally {
    event.target.value = "";
  }
});

elements.loadSheet.addEventListener("click", async () => {
  const url = elements.sheetUrl.value.trim();
  if (!url) {
    updateStatus("Please paste a Google Sheets URL", true);
    return;
  }
  const sheetId = extractGoogleSheetId(url);
  if (!sheetId) {
    updateStatus("Could not parse the Google Sheets link", true);
    return;
  }
  try {
    updateStatus("Fetching Google Sheet…");
    const response = await fetch(
      `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv`
    );
    if (!response.ok) {
      throw new Error(`Google Sheets request failed (${response.status})`);
    }
    const csvText = await response.text();
    const rows = parseCsv(csvText);
    state.rows = mapRows(rows);
    state.collapsed.clear();
    state.search = "";
    elements.searchInput.value = "";
    render();
    updateStatus(`Loaded ${state.rows.length} rows from Google Sheets`);
  } catch (error) {
    console.error(error);
    updateStatus(`Failed to fetch Google Sheets data: ${error.message}`, true);
  }
});

elements.loadSample.addEventListener("click", () => {
  state.rows = [...SAMPLE_DATA];
  state.collapsed.clear();
  state.search = "";
  elements.searchInput.value = "";
  render();
  updateStatus("Loaded sample organization data");
});

elements.searchInput.addEventListener("input", (event) => {
  state.search = event.target.value.trim();
  render();
});

elements.clearSearch.addEventListener("click", () => {
  if (!state.search) return;
  state.search = "";
  elements.searchInput.value = "";
  render();
});

elements.fab.addEventListener("click", (event) => {
  const button = event.target.closest(".fab-button");
  if (!button) return;
  const action = button.dataset.action;
  switch (action) {
    case "zoom-in":
      setScale(state.scale + SCALE_STEP);
      break;
    case "zoom-out":
      setScale(state.scale - SCALE_STEP);
      break;
    case "zoom-fit":
      zoomToFit();
      break;
    case "expand":
      state.collapsed.clear();
      render();
      break;
    case "collapse":
      collapseAll();
      break;
    default:
      break;
  }
});

elements.chartContainer.addEventListener("scroll", () => {
  drawMinimap(state.layout);
});

window.addEventListener("resize", () => {
  drawMinimap(state.layout);
});

elements.minimap.addEventListener("click", (event) => {
  if (!state.layout) return;
  const { widthPx, heightPx } = state.layout;
  if (!widthPx || !heightPx) return;

  const rect = elements.minimap.getBoundingClientRect();
  const x = event.clientX - rect.left;
  const y = event.clientY - rect.top;

  const { scale, offsetX, offsetY } = getMinimapScaling(widthPx, heightPx);
  const chartX = (x - offsetX) / scale;
  const chartY = (y - offsetY) / scale;

  const targetX = chartX * state.scale - elements.chartContainer.clientWidth / 2;
  const targetY = chartY * state.scale - elements.chartContainer.clientHeight / 2;
  elements.chartContainer.scrollTo({
    left: clamp(targetX, 0, elements.chartContainer.scrollWidth),
    top: clamp(targetY, 0, elements.chartContainer.scrollHeight),
    behavior: "smooth",
  });
});

function setScale(next) {
  const clamped = clamp(next, MIN_SCALE, MAX_SCALE);
  if (Math.abs(clamped - state.scale) < 0.001) return;
  state.scale = clamped;
  render();
}

function zoomToFit() {
  if (!state.layout) return;
  const containerWidth = elements.chartContainer.clientWidth;
  const containerHeight = elements.chartContainer.clientHeight;
  if (containerWidth === 0 || containerHeight === 0) return;

  const { widthPx, heightPx } = state.layout;
  if (!widthPx || !heightPx) return;

  const scaleX = containerWidth / widthPx;
  const scaleY = containerHeight / heightPx;
  const nextScale = clamp(Math.min(scaleX, scaleY) * 0.92, MIN_SCALE, MAX_SCALE);
  state.scale = nextScale;
  render();
}

function collapseAll() {
  const tree = buildTree();
  const ids = [];
  const collect = (node) => {
    if (node.children?.length) {
      ids.push(node.id);
      node.children.forEach(collect);
    }
  };
  tree.forEach(collect);
  state.collapsed = new Set(ids);
  render();
}

function render() {
  if (!state.rows.length) {
    elements.chartNodes.innerHTML = "";
    elements.chartLinks.innerHTML = "";
    state.layout = null;
    elements.emptyState.style.display = "flex";
    drawMinimap(null);
    return;
  }

  const tree = buildTree();
  const filteredTree = applySearch(tree, state.search);
  const layout = computeLayout(filteredTree);
  state.layout = layout;

  renderLinks(layout);
  renderNodes(layout);
  elements.emptyState.style.display = "none";
  drawMinimap(layout);
}

function buildTree() {
  return state.view === "manager"
    ? buildManagerTree(state.rows)
    : buildTribeTree(state.rows);
}

function buildManagerTree(rows) {
  const employeeNodes = [];
  const employeesByName = new Map();
  const placeholderManagers = new Map();

  rows.forEach((row, index) => {
    const label = row.name || `Employee ${index + 1}`;
    const node = {
      id: `emp-${index}`,
      key: `${label}-${index}`,
      label,
      type: "employee",
      data: row,
      children: [],
      originalChildren: [],
      collapsed: false,
      hasChildren: false,
      isMatch: false,
      visible: true,
      searchTerms: buildSearchTerms(row, label),
    };
    employeeNodes.push(node);
    const list = employeesByName.get(label) || [];
    list.push(node);
    employeesByName.set(label, list);
  });

  const getManagerNode = (name) => {
    if (!name) return null;
    const trimmed = name.trim();
    if (!trimmed) return null;
    const employeeList = employeesByName.get(trimmed);
    if (employeeList && employeeList.length) {
      return employeeList[0];
    }
    if (!placeholderManagers.has(trimmed)) {
      const placeholder = {
        id: `placeholder-${placeholderManagers.size}`,
        key: trimmed,
        label: trimmed,
        type: "placeholder",
        data: {
          name: trimmed,
          tribe: "",
          squad: "",
          expertise: "",
          level: "",
          manager: "",
        },
        children: [],
        originalChildren: [],
        collapsed: false,
        hasChildren: false,
        isMatch: false,
        visible: true,
        searchTerms: buildSearchTerms({ name: trimmed }, trimmed),
      };
      placeholderManagers.set(trimmed, placeholder);
    }
    return placeholderManagers.get(trimmed);
  };

  const roots = [];

  employeeNodes.forEach((node) => {
    const managerName = node.data.manager?.trim();
    if (!managerName) {
      roots.push(node);
      return;
    }
    const managerNode = getManagerNode(managerName);
    if (!managerNode) {
      roots.push(node);
      return;
    }
    if (managerNode === node) {
      roots.push(node);
      return;
    }
    managerNode.children.push(node);
    node.parent = managerNode;
  });

  placeholderManagers.forEach((node) => {
    if (!node.parent) {
      roots.push(node);
    }
  });

  const collapsedSet = state.search ? new Set() : state.collapsed;
  const finalize = (node) => {
    node.originalChildren = [...node.children];
    node.hasChildren = node.originalChildren.length > 0;
    if (collapsedSet.has(node.id)) {
      node.children = [];
      node.collapsed = true;
    } else {
      node.children = node.originalChildren.map(finalize);
      node.collapsed = false;
    }
    return node;
  };

  return roots.map(finalize);
}

function buildTribeTree(rows) {
  const tribeMap = new Map();

  const getTribeNode = (tribeName) => {
    const key = tribeName || "Unassigned Tribe";
    if (!tribeMap.has(key)) {
      tribeMap.set(key, {
        id: `tribe-${key}`,
        key,
        label: key,
        type: "tribe",
        data: { tribe: key },
        children: [],
        originalChildren: [],
        collapsed: false,
        hasChildren: false,
        isMatch: false,
        visible: true,
        searchTerms: buildSearchTerms({ tribe: key }, key),
      });
    }
    return tribeMap.get(key);
  };

  const squadMap = new Map();
  const getSquadNode = (tribeNode, squadName) => {
    const key = `${tribeNode.key}__${squadName || "Unassigned Squad"}`;
    if (!squadMap.has(key)) {
      const label = squadName || "Unassigned Squad";
      const node = {
        id: `squad-${key}`,
        key,
        label,
        type: "squad",
        data: { tribe: tribeNode.label, squad: label },
        children: [],
        originalChildren: [],
        collapsed: false,
        hasChildren: false,
        isMatch: false,
        visible: true,
        searchTerms: buildSearchTerms({ tribe: tribeNode.label, squad: label }, label),
      };
      squadMap.set(key, node);
      tribeNode.children.push(node);
      node.parent = tribeNode;
    }
    return squadMap.get(key);
  };

  rows.forEach((row) => {
    const tribeNode = getTribeNode(row.tribe);
    const squadNode = getSquadNode(tribeNode, row.squad);
    const employeeNode = {
      id: `emp-${row.name || crypto.randomUUID?.() || Math.random()}`,
      key: row.name || Math.random().toString(36).slice(2),
      label: row.name || "Unnamed",
      type: "employee",
      data: row,
      children: [],
      originalChildren: [],
      collapsed: false,
      hasChildren: false,
      isMatch: false,
      visible: true,
      searchTerms: buildSearchTerms(row, row.name || "Unnamed"),
      parent: squadNode,
    };
    squadNode.children.push(employeeNode);
  });

  const collapsedSet = state.search ? new Set() : state.collapsed;
  const finalize = (node) => {
    node.originalChildren = [...node.children];
    node.hasChildren = node.originalChildren.length > 0;
    if (collapsedSet.has(node.id)) {
      node.children = [];
      node.collapsed = true;
    } else {
      node.children = node.originalChildren.map(finalize);
      node.collapsed = false;
    }
    return node;
  };

  return Array.from(tribeMap.values()).map(finalize);
}

function applySearch(roots, query) {
  if (!query) {
    const reset = (node) => {
      node.isMatch = false;
      node.visible = true;
      node.children.forEach(reset);
    };
    roots.forEach(reset);
    return roots;
  }
  const lower = query.toLowerCase();

  const filterNode = (node) => {
    const selfMatch = node.searchTerms.some((term) => term.includes(lower));
    const filteredChildren = node.children
      .map(filterNode)
      .filter((child) => child.visible);
    node.children = filteredChildren;
    node.isMatch = selfMatch;
    node.visible = selfMatch || filteredChildren.length > 0;
    return node;
  };

  return roots
    .map(filterNode)
    .filter((node) => node.visible);
}

function computeLayout(roots) {
  const positions = [];
  const links = [];
  let maxDepth = 0;

  const subtreeWidth = (node) => {
    if (!node.children || node.children.length === 0) {
      node._width = 1;
      return 1;
    }
    let width = 0;
    node.children.forEach((child) => {
      width += subtreeWidth(child);
    });
    node._width = Math.max(width, 1);
    return node._width;
  };

  const place = (node, depth, offset) => {
    const width = node._width || 1;
    const xCenter = offset + (width - 1) / 2;
    positions.push({ node, depth, x: xCenter });
    maxDepth = Math.max(maxDepth, depth);
    let childOffset = offset;
    node.children.forEach((child) => {
      place(child, depth + 1, childOffset);
      links.push({ from: node, to: child });
      childOffset += child._width || 1;
    });
  };

  let offset = 0;
  roots.forEach((root) => {
    subtreeWidth(root);
    place(root, 0, offset);
    offset += root._width || 1;
  });

  const widthUnits = offset;
  const widthPx = Math.max(widthUnits * HORIZONTAL_GAP, elements.chartContainer.clientWidth / state.scale);
  const heightPx = Math.max((maxDepth + 1) * VERTICAL_GAP, elements.chartContainer.clientHeight / state.scale);

  const positionMap = new Map();
  positions.forEach((entry) => {
    positionMap.set(entry.node, entry);
  });

  return {
    nodes: positions,
    nodePositions: positionMap,
    links,
    widthPx,
    heightPx,
  };
}

function renderNodes(layout) {
  const { nodes, widthPx, heightPx } = layout;
  const scaledWidth = Math.max(widthPx * state.scale, elements.chartContainer.clientWidth);
  const scaledHeight = Math.max(heightPx * state.scale, elements.chartContainer.clientHeight);
  elements.chartContent.style.width = `${scaledWidth}px`;
  elements.chartContent.style.height = `${scaledHeight}px`;
  elements.chartZoom.style.width = `${widthPx}px`;
  elements.chartZoom.style.height = `${heightPx}px`;
  elements.chartZoom.style.transform = `scale(${state.scale})`;
  elements.chartNodes.innerHTML = "";

  nodes.forEach(({ node, depth, x }) => {
    const card = document.createElement("article");
    card.classList.add("node-card");
    if (node.type !== "employee") {
      card.classList.add("group");
    }
    if (node.isMatch) {
      card.classList.add("match");
    }

    const top = depth * VERTICAL_GAP + VERTICAL_GAP / 2 - NODE_HEIGHT / 2;
    const left = x * HORIZONTAL_GAP + HORIZONTAL_GAP / 2 - NODE_WIDTH / 2;

    card.style.transform = `translate(${left}px, ${top}px)`;

    card.innerHTML = buildCardContent(node);

    if (node.hasChildren) {
      const toggle = document.createElement("button");
      toggle.classList.add("collapse-toggle");
      toggle.type = "button";
      toggle.textContent = node.collapsed ? "+" : "−";
      toggle.addEventListener("click", (event) => {
        event.stopPropagation();
        toggleNode(node);
      });
      card.appendChild(toggle);
    }

    card.addEventListener("click", () => {
      if (node.hasChildren) {
        toggleNode(node);
      }
    });

    elements.chartNodes.appendChild(card);
  });
}

function renderLinks(layout) {
  const { links, nodePositions } = layout;
  const svgNS = "http://www.w3.org/2000/svg";
  elements.chartLinks.setAttribute("width", "100%");
  elements.chartLinks.setAttribute("height", "100%");
  elements.chartLinks.innerHTML = "";

  links.forEach(({ from, to }) => {
    const parentPos = nodePositions.get(from);
    const childPos = nodePositions.get(to);
    if (!parentPos || !childPos) return;

    const startX = parentPos.x * HORIZONTAL_GAP + HORIZONTAL_GAP / 2;
    const startY = parentPos.depth * VERTICAL_GAP + VERTICAL_GAP / 2 + NODE_HEIGHT / 2;
    const endX = childPos.x * HORIZONTAL_GAP + HORIZONTAL_GAP / 2;
    const endY = childPos.depth * VERTICAL_GAP + VERTICAL_GAP / 2 - NODE_HEIGHT / 2;
    const controlY = (startY + endY) / 2;

    const path = document.createElementNS(svgNS, "path");
    path.setAttribute(
      "d",
      `M${startX},${startY} C ${startX},${controlY} ${endX},${controlY} ${endX},${endY}`
    );
    path.setAttribute("fill", "none");
    path.setAttribute("stroke", "rgba(148, 163, 184, 0.5)");
    path.setAttribute("stroke-width", "2");
    elements.chartLinks.appendChild(path);
  });
}

function toggleNode(node) {
  if (state.search) return; // disable toggling while filtering
  if (state.collapsed.has(node.id)) {
    state.collapsed.delete(node.id);
  } else {
    state.collapsed.add(node.id);
  }
  render();
}

function drawMinimap(layout) {
  const ctx = elements.minimap.getContext("2d");
  ctx.clearRect(0, 0, elements.minimap.width, elements.minimap.height);

  if (!layout || !layout.nodes.length) {
    ctx.fillStyle = "rgba(148, 163, 184, 0.3)";
    ctx.font = "12px Inter, sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("No data", elements.minimap.width / 2, elements.minimap.height / 2);
    return;
  }

  const { widthPx, heightPx } = layout;
  const { scale, offsetX, offsetY } = getMinimapScaling(widthPx, heightPx);

  ctx.save();
  ctx.translate(offsetX, offsetY);
  ctx.scale(scale, scale);

  ctx.strokeStyle = "rgba(148, 163, 184, 0.35)";
  ctx.lineWidth = 1 / scale;
  layout.links.forEach(({ from, to }) => {
    const parentPos = layout.nodePositions.get(from);
    const childPos = layout.nodePositions.get(to);
    if (!parentPos || !childPos) return;
    const startX = parentPos.x * HORIZONTAL_GAP + HORIZONTAL_GAP / 2;
    const startY = parentPos.depth * VERTICAL_GAP + VERTICAL_GAP / 2;
    const endX = childPos.x * HORIZONTAL_GAP + HORIZONTAL_GAP / 2;
    const endY = childPos.depth * VERTICAL_GAP + VERTICAL_GAP / 2;
    ctx.beginPath();
    ctx.moveTo(startX, startY);
    ctx.lineTo(endX, endY);
    ctx.stroke();
  });

  layout.nodes.forEach(({ node, depth, x }) => {
    const cx = x * HORIZONTAL_GAP + HORIZONTAL_GAP / 2;
    const cy = depth * VERTICAL_GAP + VERTICAL_GAP / 2;
    ctx.fillStyle = node.type === "employee" ? "#2563eb" : "#0ea5e9";
    ctx.fillRect(cx - 6, cy - 6, 12, 12);
  });

  ctx.restore();

  // viewport rectangle
  const visibleWidth = elements.chartContainer.clientWidth / state.scale;
  const visibleHeight = elements.chartContainer.clientHeight / state.scale;
  const scrollLeft = elements.chartContainer.scrollLeft / state.scale;
  const scrollTop = elements.chartContainer.scrollTop / state.scale;

  ctx.save();
  ctx.translate(offsetX, offsetY);
  ctx.scale(scale, scale);
  ctx.strokeStyle = "rgba(37, 99, 235, 0.8)";
  ctx.lineWidth = 1 / scale;
  ctx.strokeRect(scrollLeft, scrollTop, visibleWidth, visibleHeight);
  ctx.restore();
}

function getMinimapScaling(widthPx, heightPx) {
  const padding = 16;
  const maxWidth = elements.minimap.width - padding * 2;
  const maxHeight = elements.minimap.height - padding * 2;
  const scale = Math.min(maxWidth / widthPx, maxHeight / heightPx, 1);
  const offsetX = (elements.minimap.width - widthPx * scale) / 2;
  const offsetY = (elements.minimap.height - heightPx * scale) / 2;
  return { scale, offsetX, offsetY };
}

function mapRows(rawRows) {
  return rawRows
    .map((row) => normalizeRow(row))
    .filter((row) => row.name);
}

function normalizeRow(row) {
  const normalized = {};
  Object.entries(row).forEach(([key, value]) => {
    normalized[normalizeColumnName(key)] = typeof value === "string" ? value.trim() : value;
  });

  const result = {
    name: pickAlias(normalized, "name"),
    tribe: pickAlias(normalized, "tribe"),
    squad: pickAlias(normalized, "squad"),
    expertise: pickAlias(normalized, "expertise"),
    level: pickAlias(normalized, "level"),
    manager: pickAlias(normalized, "manager"),
  };

  return Object.fromEntries(
    Object.entries(result).map(([key, value]) => [key, typeof value === "string" ? value.trim() : value])
  );
}

function pickAlias(row, field) {
  const aliases = COLUMN_ALIASES[field] || [];
  for (const alias of aliases) {
    if (row.hasOwnProperty(alias)) {
      return row[alias];
    }
  }
  return row[field] ?? "";
}

function normalizeColumnName(name) {
  return name
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/[^a-z0-9 ]+/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function buildSearchTerms(row, fallback) {
  return [
    fallback?.toLowerCase?.() || "",
    row.name?.toLowerCase?.() || "",
    row.tribe?.toLowerCase?.() || "",
    row.squad?.toLowerCase?.() || "",
    row.expertise?.toLowerCase?.() || "",
    row.level?.toLowerCase?.() || "",
    row.manager?.toLowerCase?.() || "",
  ].filter(Boolean);
}

function buildCardContent(node) {
  if (node.type === "tribe") {
    return `
      <h3 class="title">${escapeHtml(node.label)}</h3>
      <p class="meta">Tribe · ${node.originalChildren.length} squads</p>
    `;
  }
  if (node.type === "squad") {
    return `
      <div class="badge">Squad</div>
      <h3 class="title">${escapeHtml(node.label)}</h3>
      <p class="meta">${node.originalChildren.length} people · ${escapeHtml(node.data.tribe)}</p>
    `;
  }
  if (node.type === "placeholder") {
    return `
      <div class="badge">Manager</div>
      <h3 class="title">${escapeHtml(node.label)}</h3>
      <p class="meta">${node.originalChildren.length} direct reports</p>
    `;
  }

  const { expertise, tribe, squad, level, manager } = node.data;
  const teamLine = [tribe, squad].filter(Boolean).map(escapeHtml).join(" · ");
  return `
    <div class="badge">${escapeHtml(level || "Employee")}</div>
    <h3 class="title">${escapeHtml(node.label)}</h3>
    <p class="meta">
      <span>${escapeHtml(expertise || "")}</span>
      ${teamLine ? `<span>${teamLine}</span>` : ""}
      <span>Manager: ${escapeHtml(manager || "–")}</span>
    </p>
  `;
}

function toggleLoader(active) {
  elements.chartContainer.classList.toggle("loading", active);
}

function updateStatus(message, isError = false) {
  elements.statusMessage.textContent = message;
  elements.statusBar.style.color = isError ? "#ef4444" : "var(--color-subtle)";
}

function escapeHtml(value) {
  if (value == null) return "";
  return value
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function extractGoogleSheetId(url) {
  const patterns = [
    /spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
    /spreadsheets\?id=([a-zA-Z0-9-_]+)/,
  ];
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match) return match[1];
  }
  return null;
}

function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter(Boolean);
  if (!lines.length) return [];
  const headers = splitCsvLine(lines[0]);
  return lines.slice(1).map((line) => {
    const cells = splitCsvLine(line);
    const row = {};
    headers.forEach((header, index) => {
      row[header] = cells[index] ?? "";
    });
    return row;
  });
}

function splitCsvLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      result.push(current);
      current = "";
    } else {
      current += char;
    }
  }
  result.push(current);
  return result.map((value) => value.trim());
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function updateThemeButton(theme) {
  elements.themeToggle.querySelector(".icon").textContent = theme === "dark" ? "☀️" : "🌙";
}

// initial render with sample data for a quick start
if (!state.rows.length) {
  state.rows = [...SAMPLE_DATA];
  render();
  updateStatus("Showing sample data. Upload a file to replace.");
}
