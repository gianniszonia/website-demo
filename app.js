(function () {
  "use strict";

  var workbookPath = "./sample_-_superstore.xls";
  var accentByShipMode = {
    "Standard Class": "#4996b2",
    "Second Class": "#22c55e",
    "First Class": "#ef4444",
    "Same Day": "#f59e0b"
  };

  var state = {
    rows: [],
    cardHoverByShipMode: {}
  };

  var dataStatusEl = document.getElementById("dataStatus");
  var regionFilterEl = document.getElementById("regionFilter");
  var segmentFilterEl = document.getElementById("segmentFilter");
  var categoryFilterEl = document.getElementById("categoryFilter");
  var cardsGridEl = document.getElementById("cardsGrid");

  function parseWorkbookDate(value) {
    if (value instanceof Date && !isNaN(value.getTime())) return value;

    if (typeof value === "number" && isFinite(value)) {
      var parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
    }

    if (typeof value === "string" && value.trim()) {
      var normalized = new Date(value);
      if (!isNaN(normalized.getTime())) return normalized;
    }

    return null;
  }

  function formatCurrencyK(value) {
    var abs = Math.abs(value || 0);
    if (abs >= 1000) return Math.round((value || 0) / 1000) + "K";
    return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(value || 0);
  }

  function formatDelta(delta) {
    if (delta === null || !isFinite(delta)) return "—";
    var rounded = Math.round(delta * 10) / 10;
    if (Object.is(rounded, -0)) rounded = 0;
    return (rounded > 0 ? "+" : "") + rounded.toFixed(0) + "%";
  }

  function monthKey(date) {
    return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");
  }

  function monthShortLabel(date) {
    return date.toLocaleString("en-US", { month: "short" }).toUpperCase();
  }

  function monthLongLabel(date) {
    return date.toLocaleString("en-US", { month: "short" }) + " " + String(date.getFullYear()).slice(-2);
  }

  function populateFilter(selectEl, values, label) {
    selectEl.innerHTML = "";
    var allOption = document.createElement("option");
    allOption.value = "";
    allOption.textContent = "All " + label;
    selectEl.appendChild(allOption);

    values.forEach(function (value) {
      var option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      selectEl.appendChild(option);
    });
  }

  function normalizeRows(rawRows) {
    return rawRows
      .map(function (row) {
        var orderDate = parseWorkbookDate(row["Order Date"]);
        var sales = typeof row.Sales === "number" ? row.Sales : parseFloat(String(row.Sales || "").replace(/,/g, ""));

        return {
          Region: row.Region || "",
          Segment: row.Segment || "",
          Category: row.Category || "",
          ShipMode: row["Ship Mode"] || "",
          orderDate: orderDate,
          sales: sales
        };
      })
      .filter(function (row) {
        return row.orderDate && isFinite(row.sales) && row.ShipMode;
      });
  }

  function getFilteredRows() {
    return state.rows.filter(function (row) {
      if (regionFilterEl.value && row.Region !== regionFilterEl.value) return false;
      if (segmentFilterEl.value && row.Segment !== segmentFilterEl.value) return false;
      if (categoryFilterEl.value && row.Category !== categoryFilterEl.value) return false;
      return true;
    });
  }

  function buildMonthlyData(rows) {
    var byMonth = new Map();

    rows.forEach(function (row) {
      var key = monthKey(row.orderDate);
      if (!byMonth.has(key)) {
        byMonth.set(key, {
          key: key,
          date: new Date(row.orderDate.getFullYear(), row.orderDate.getMonth(), 1),
          sales: 0
        });
      }
      byMonth.get(key).sales += row.sales;
    });

    return Array.from(byMonth.values()).sort(function (a, b) {
      return a.date - b.date;
    });
  }

  function buildShipModeSeries(rows) {
    var byShipMode = new Map();

    rows.forEach(function (row) {
      if (!byShipMode.has(row.ShipMode)) byShipMode.set(row.ShipMode, []);
      byShipMode.get(row.ShipMode).push(row);
    });

    return Array.from(byShipMode.entries())
      .map(function (entry) {
        return {
          shipMode: entry[0],
          months: buildMonthlyData(entry[1]).slice(-8)
        };
      })
      .sort(function (a, b) {
        return a.shipMode.localeCompare(b.shipMode);
      });
  }

  function getSelectedMonth(months, shipMode) {
    if (!months.length) return null;
    var hoverKey = state.cardHoverByShipMode[shipMode];
    if (!hoverKey) return months[months.length - 1];

    for (var i = 0; i < months.length; i += 1) {
      if (months[i].key === hoverKey) return months[i];
    }

    return months[months.length - 1];
  }

  function findMonthByOffset(months, currentIndex, offset) {
    var index = currentIndex - offset;
    if (index < 0 || index >= months.length) return null;
    return months[index];
  }

  function buildDeltaRow(label, deltaValue) {
    var row = document.createElement("div");
    row.className = "delta-row";

    var pill = document.createElement("span");
    pill.className = "delta-pill";

    if (deltaValue === null || !isFinite(deltaValue)) {
      pill.classList.add("neutral");
    } else if (deltaValue >= 0) {
      pill.classList.add("positive");
    } else {
      pill.classList.add("negative");
    }

    pill.textContent = formatDelta(deltaValue);

    var text = document.createElement("span");
    text.className = "delta-label";
    text.textContent = label;

    row.appendChild(pill);
    row.appendChild(text);
    return row;
  }

  function renderCard(series) {
    var card = document.createElement("article");
    card.className = "split-card";

    var accentColor = accentByShipMode[series.shipMode] || "#4996b2";
    var selected = getSelectedMonth(series.months, series.shipMode);

    if (!selected) {
      card.innerHTML = '<div class="card-header"><div class="metric-row"><div class="metric-meta"><span class="metric-accent" style="background:' + accentColor + '"></span><span class="metric-title">Sales – ' + series.shipMode + '</span></div></div><div class="divider-line"></div></div><div class="card-body"><div class="empty-state">No rows match the current filters for this ship mode.</div></div>';
      return card;
    }

    var selectedIndex = series.months.findIndex(function (month) {
      return month.key === selected.key;
    });
    var previousMonth = findMonthByOffset(series.months, selectedIndex, 1);
    var previousYear = findMonthByOffset(series.months, selectedIndex, 12);

    var previousMonthDelta = previousMonth && previousMonth.sales !== 0
      ? ((selected.sales - previousMonth.sales) / Math.abs(previousMonth.sales)) * 100
      : null;
    var previousYearDelta = previousYear && previousYear.sales !== 0
      ? ((selected.sales - previousYear.sales) / Math.abs(previousYear.sales)) * 100
      : null;

    var header = document.createElement("div");
    header.className = "card-header";
    header.innerHTML =
      '<div class="metric-row">' +
        '<div class="metric-meta">' +
          '<span class="metric-accent" style="background:' + accentColor + '"></span>' +
          '<span class="metric-title">Sales – ' + series.shipMode + '</span>' +
        '</div>' +
      '</div>' +
      '<div class="divider-line"></div>';

    var body = document.createElement("div");
    body.className = "card-body";

    var summaryTop = document.createElement("div");
    summaryTop.className = "summary-top";
    summaryTop.innerHTML =
      '<div class="summary-value">' + formatCurrencyK(selected.sales) + '</div>' +
      '<div class="summary-date">' + monthLongLabel(selected.date) + '</div>';

    var deltas = document.createElement("div");
    deltas.className = "summary-deltas";
    deltas.appendChild(buildDeltaRow("vs previous month", previousMonthDelta));
    deltas.appendChild(buildDeltaRow("vs previous year", previousYearDelta));

    var chart = document.createElement("div");
    chart.className = "mini-bar-chart";

    var maxValue = Math.max.apply(null, series.months.map(function (month) {
      return month.sales;
    }));

    series.months.forEach(function (month) {
      var column = document.createElement("div");
      column.className = "bar-column";
      if (month.key === selected.key) column.classList.add("active");

      var value = document.createElement("div");
      value.className = "bar-value";
      value.textContent = formatCurrencyK(month.sales);

      var plot = document.createElement("div");
      plot.className = "bar-plot";

      var fill = document.createElement("div");
      fill.className = "bar-fill";
      fill.style.height = Math.max(10, Math.round((month.sales / maxValue) * 250)) + "px";

      var label = document.createElement("div");
      label.className = "bar-label";
      label.textContent = monthShortLabel(month.date);

      plot.appendChild(fill);
      column.appendChild(value);
      column.appendChild(plot);
      column.appendChild(label);

      column.addEventListener("mouseenter", function () {
        state.cardHoverByShipMode[series.shipMode] = month.key;
        render();
      });

      column.addEventListener("mouseleave", function () {
        delete state.cardHoverByShipMode[series.shipMode];
        render();
      });

      chart.appendChild(column);
    });

    var note = document.createElement("div");
    note.className = "card-note";
    note.textContent = "Hover bars to preview a month within this card.";

    body.appendChild(summaryTop);
    body.appendChild(deltas);
    body.appendChild(chart);
    body.appendChild(note);
    card.appendChild(header);
    card.appendChild(body);

    return card;
  }

  function render() {
    var filteredRows = getFilteredRows();
    var series = buildShipModeSeries(filteredRows);
    cardsGridEl.innerHTML = "";

    if (!series.length) {
      cardsGridEl.innerHTML = '<div class="controls-card"><div class="empty-state">No data matches the current filters.</div></div>';
      return;
    }

    series.forEach(function (entry) {
      cardsGridEl.appendChild(renderCard(entry));
    });
  }

  function hydrateFilters() {
    populateFilter(regionFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.Region; }))).sort(), "Regions");
    populateFilter(segmentFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.Segment; }))).sort(), "Segments");
    populateFilter(categoryFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.Category; }))).sort(), "Categories");

    [regionFilterEl, segmentFilterEl, categoryFilterEl].forEach(function (el) {
      el.addEventListener("change", function () {
        state.cardHoverByShipMode = {};
        render();
      });
    });
  }

  function loadWorkbook() {
    fetch(workbookPath)
      .then(function (response) {
        if (!response.ok) throw new Error("Could not load sample workbook.");
        return response.arrayBuffer();
      })
      .then(function (buffer) {
        var workbook = XLSX.read(buffer, { type: "array" });
        var firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        var rawRows = XLSX.utils.sheet_to_json(firstSheet, {
          raw: true,
          defval: ""
        });

        state.rows = normalizeRows(rawRows);
        hydrateFilters();
        render();
        dataStatusEl.textContent = "Loaded Superstore sample data successfully.";
      })
      .catch(function (error) {
        console.error(error);
        dataStatusEl.textContent = "Failed to load sample data.";
        cardsGridEl.innerHTML = '<div class="controls-card"><div class="empty-state">Check that <code>sample_-_superstore.xls</code> is present in this folder when you publish the demo.</div></div>';
      });
  }

  loadWorkbook();
})();
