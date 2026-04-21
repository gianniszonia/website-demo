(function () {
  "use strict";

  var workbookPath = "./sample_-_superstore.xls";
  var cardPages = [
    { key: "bar", label: "Bar Chart" },
    { key: "line", label: "Line Chart" },
    { key: "waterfall", label: "Waterfall" },
    { key: "radial", label: "Radial" },
    { key: "funnel", label: "Funnel" }
  ];
  var accentByShipMode = {
    "Standard Class": "#4996b2",
    "Second Class": "#22c55e",
    "First Class": "#ef4444",
    "Same Day": "#f59e0b"
  };

  var state = {
    rows: [],
    currentPage: 0
  };

  var dataStatusEl = document.getElementById("dataStatus");
  var regionFilterEl = document.getElementById("regionFilter");
  var categoryFilterEl = document.getElementById("categoryFilter");
  var subCategoryFilterEl = document.getElementById("subCategoryFilter");
  var stateFilterEl = document.getElementById("stateFilter");
  var cardsGridEl = document.getElementById("cardsGrid");
  var tooltipEl = document.getElementById("chartTooltip");

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

  function formatCompact(value) {
    var numeric = value || 0;
    var abs = Math.abs(numeric);
    if (abs >= 1000000) return (numeric / 1000000).toFixed(abs >= 10000000 ? 0 : 1).replace(".0", "") + "M";
    if (abs >= 1000) return Math.round(numeric / 1000) + "K";
    return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(numeric);
  }

  function formatCurrency(value) {
    return "$" + formatCompact(value);
  }

  function formatCurrencyFull(value) {
    return new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
      maximumFractionDigits: 0
    }).format(value || 0);
  }

  function formatPercent(delta) {
    if (delta === null || !isFinite(delta)) return "-";
    var rounded = Math.round(delta);
    if (Object.is(rounded, -0)) rounded = 0;
    return (rounded > 0 ? "+" : "") + rounded + "%";
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function monthKey(date) {
    return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");
  }

  function previousYearKey(date) {
    return (date.getFullYear() - 1) + "-" + String(date.getMonth() + 1).padStart(2, "0");
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
          Category: row.Category || "",
          SubCategory: row["Sub-Category"] || "",
          State: row.State || "",
          Segment: row.Segment || "",
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
      if (categoryFilterEl.value && row.Category !== categoryFilterEl.value) return false;
      if (subCategoryFilterEl.value && row.SubCategory !== subCategoryFilterEl.value) return false;
      if (stateFilterEl.value && row.State !== stateFilterEl.value) return false;
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
          rows: entry[1],
          months: buildMonthlyData(entry[1])
        };
      })
      .sort(function (a, b) {
        var aLatest = a.months.length ? a.months[a.months.length - 1].sales : 0;
        var bLatest = b.months.length ? b.months[b.months.length - 1].sales : 0;
        return bLatest - aLatest;
      });
  }

  function buildSummary(series) {
    if (!series.months.length) return null;
    var selected = series.months[series.months.length - 1];
    var previousMonth = series.months.length > 1 ? series.months[series.months.length - 2] : null;
    var byKey = {};
    series.months.forEach(function (month) {
      byKey[month.key] = month;
    });
    var previousYear = byKey[previousYearKey(selected.date)] || null;
    return {
      selected: selected,
      previousMonthDelta: previousMonth && previousMonth.sales !== 0
        ? ((selected.sales - previousMonth.sales) / Math.abs(previousMonth.sales)) * 100
        : null,
      previousYearDelta: previousYear && previousYear.sales !== 0
        ? ((selected.sales - previousYear.sales) / Math.abs(previousYear.sales)) * 100
        : null
    };
  }

  function tooltipHtml(label, value, extra) {
    var lines = [
      '<div class="tt-label">' + escapeHtml(label) + "</div>",
      "<div>" + escapeHtml(value) + "</div>"
    ];
    if (extra) lines.push('<div class="tt-label">' + escapeHtml(extra) + "</div>");
    return lines.join("");
  }

  function showTooltip(event, html) {
    tooltipEl.innerHTML = html;
    tooltipEl.classList.add("visible");
    tooltipEl.style.left = event.clientX + 14 + "px";
    tooltipEl.style.top = event.clientY + 14 + "px";
  }

  function hideTooltip() {
    tooltipEl.classList.remove("visible");
  }

  function attachTooltip(el, getHtml) {
    el.addEventListener("mouseenter", function (event) {
      showTooltip(event, getHtml());
    });
    el.addEventListener("mousemove", function (event) {
      showTooltip(event, getHtml());
    });
    el.addEventListener("mouseleave", hideTooltip);
  }

  function buildDeltaRow(label, deltaValue) {
    var row = document.createElement("div");
    row.className = "delta-row";
    var badge = document.createElement("span");
    badge.className = "kpi-delta";
    if (deltaValue === null || !isFinite(deltaValue)) badge.classList.add("neutral");
    else if (deltaValue >= 0) badge.classList.add("positive");
    else badge.classList.add("negative");
    badge.textContent = formatPercent(deltaValue);
    var text = document.createElement("span");
    text.className = "delta-text";
    text.textContent = label;
    row.appendChild(badge);
    row.appendChild(text);
    return row;
  }

  function createSvg(tag) {
    return document.createElementNS("http://www.w3.org/2000/svg", tag);
  }

  function recentMonthsDescending(series, count) {
    return series.months.slice(-count).reverse();
  }

  function renderBarChart(series, mount) {
    var months = recentMonthsDescending(series, 8);
    var wrap = document.createElement("div");
    wrap.className = "w-bars-shell";
    var chart = document.createElement("div");
    chart.className = "w-bars";
    var maxValue = Math.max.apply(null, months.map(function (month) { return month.sales; }).concat([1]));

    months.forEach(function (month) {
      var column = document.createElement("div");
      column.className = "bar-col";

      var plot = document.createElement("div");
      plot.className = "bar-plot";

      var value = document.createElement("span");
      value.className = "bar-val";
      value.textContent = formatCurrency(month.sales);
      plot.appendChild(value);

      var bar = document.createElement("div");
      bar.className = "bar";
      bar.style.background = "#4d97b2";
      bar.style.top = Math.max(0, 240 - Math.max(8, Math.round((month.sales / maxValue) * 240))) + "px";
      bar.style.height = Math.max(8, Math.round((month.sales / maxValue) * 240)) + "px";
      plot.appendChild(bar);

      var label = document.createElement("span");
      label.className = "bar-month";
      label.textContent = monthShortLabel(month.date);

      attachTooltip(column, function () {
        return tooltipHtml(monthLongLabel(month.date), formatCurrencyFull(month.sales));
      });

      column.appendChild(plot);
      column.appendChild(label);
      chart.appendChild(column);
    });

    wrap.appendChild(chart);
    mount.appendChild(wrap);
  }

  function renderLineChart(series, mount) {
    var months = recentMonthsDescending(series, 8);
    var wrap = document.createElement("div");
    wrap.className = "w-line";
    var svg = createSvg("svg");
    var width = 360;
    var height = 320;
    var pad = { top: 22, right: 16, bottom: 34, left: 28 };
    var plotWidth = width - pad.left - pad.right;
    var plotHeight = height - pad.top - pad.bottom;
    var maxValue = Math.max.apply(null, months.map(function (month) { return month.sales; }).concat([1]));
    var minValue = Math.min.apply(null, months.map(function (month) { return month.sales; }).concat([0]));
    var range = Math.max(1, maxValue - minValue);

    svg.setAttribute("viewBox", "0 0 " + width + " " + height);

    var points = months.map(function (month, index) {
      return {
        month: month,
        x: pad.left + (plotWidth / Math.max(1, months.length - 1)) * index,
        y: pad.top + plotHeight - (((month.sales - minValue) / range) * plotHeight)
      };
    });

    var pathData = points.map(function (point, index) {
      return (index === 0 ? "M" : "L") + point.x.toFixed(1) + "," + point.y.toFixed(1);
    }).join(" ");

    var area = createSvg("path");
    area.setAttribute("d", pathData + " L " + points[points.length - 1].x.toFixed(1) + "," + (height - pad.bottom) + " L " + points[0].x.toFixed(1) + "," + (height - pad.bottom) + " Z");
    area.setAttribute("fill", "rgba(73,150,178,0.12)");

    var line = createSvg("path");
    line.setAttribute("d", pathData);
    line.setAttribute("fill", "none");
    line.setAttribute("stroke", "#4996b2");
    line.setAttribute("stroke-width", "3");
    line.setAttribute("stroke-linecap", "round");
    line.setAttribute("stroke-linejoin", "round");

    svg.appendChild(area);
    svg.appendChild(line);

    points.forEach(function (point) {
      var dot = createSvg("circle");
      dot.setAttribute("cx", point.x);
      dot.setAttribute("cy", point.y);
      dot.setAttribute("r", "4");
      dot.setAttribute("fill", "#4996b2");
      attachTooltip(dot, function () {
        return tooltipHtml(monthLongLabel(point.month.date), formatCurrencyFull(point.month.sales));
      });
      svg.appendChild(dot);

      var text = createSvg("text");
      text.setAttribute("x", point.x);
      text.setAttribute("y", height - 10);
      text.setAttribute("text-anchor", "middle");
      text.setAttribute("font-size", "9");
      text.setAttribute("font-weight", "600");
      text.setAttribute("fill", "#666666");
      text.textContent = monthShortLabel(point.month.date);
      svg.appendChild(text);
    });

    wrap.appendChild(svg);
    mount.appendChild(wrap);
  }

  function renderWaterfallChart(series, mount) {
    var months = recentMonthsDescending(series, 8);
    var changes = months.map(function (month, index) {
      var nextMonth = months[index + 1];
      return {
        date: month.date,
        sales: month.sales,
        delta: nextMonth ? month.sales - nextMonth.sales : month.sales
      };
    });

    var wrap = document.createElement("div");
    wrap.className = "w-bars-shell";
    var chart = document.createElement("div");
    chart.className = "w-bars";
    var maxValue = Math.max.apply(null, changes.map(function (item) { return Math.abs(item.delta); }).concat([1]));

    changes.forEach(function (item) {
      var column = document.createElement("div");
      column.className = "bar-col";

      var plot = document.createElement("div");
      plot.className = "bar-plot";

      var value = document.createElement("span");
      value.className = "bar-val";
      value.textContent = formatCurrency(item.delta);
      plot.appendChild(value);

      var bar = document.createElement("div");
      bar.className = "bar";
      bar.style.background = item.delta >= 0 ? "#22c55e" : "#ef4444";
      bar.style.top = Math.max(0, 240 - Math.max(8, Math.round((Math.abs(item.delta) / maxValue) * 240))) + "px";
      bar.style.height = Math.max(8, Math.round((Math.abs(item.delta) / maxValue) * 240)) + "px";
      plot.appendChild(bar);

      var label = document.createElement("span");
      label.className = "bar-month";
      label.textContent = monthShortLabel(item.date);

      attachTooltip(column, function () {
        return tooltipHtml(
          monthLongLabel(item.date),
          "Change: " + formatCurrencyFull(item.delta),
          "Sales: " + formatCurrencyFull(item.sales)
        );
      });

      column.appendChild(plot);
      column.appendChild(label);
      chart.appendChild(column);
    });

    wrap.appendChild(chart);
    mount.appendChild(wrap);
  }

  function aggregateBySegment(rows) {
    var map = new Map();
    rows.forEach(function (row) {
      if (!map.has(row.Segment)) map.set(row.Segment, 0);
      map.set(row.Segment, map.get(row.Segment) + row.sales);
    });
    return Array.from(map.entries())
      .map(function (entry) {
        return { label: entry[0], value: entry[1] };
      })
      .sort(function (a, b) {
        return b.value - a.value;
      });
  }

  function renderRadialChart(series, mount) {
    var data = aggregateBySegment(series.rows);
    var wrap = document.createElement("div");
    wrap.className = "w-radial";
    var svg = createSvg("svg");
    var width = 320;
    var height = 320;
    var cx = 180;
    var cy = 224;
    var outerRadius = 92;
    var stroke = 12;
    var gap = 14;
    var maxValue = Math.max.apply(null, data.map(function (item) { return item.value; }).concat([1]));
    var colors = ["#4996b2", "#22c55e", "#ef4444"];

    svg.setAttribute("viewBox", "0 0 " + width + " " + height);

    data.forEach(function (item, index) {
      var radius = outerRadius - index * (stroke + gap);
      var circumference = 2 * Math.PI * radius;
      var trackLength = circumference * 0.75;
      var valueLength = Math.max(0, (item.value / maxValue) * trackLength);
      var transform = "rotate(-90 " + cx + " " + cy + ")";

      var track = createSvg("circle");
      track.setAttribute("cx", cx);
      track.setAttribute("cy", cy);
      track.setAttribute("r", radius);
      track.setAttribute("fill", "none");
      track.setAttribute("stroke", "#eeeeee");
      track.setAttribute("stroke-width", stroke);
      track.setAttribute("stroke-dasharray", trackLength.toFixed(2) + " " + circumference.toFixed(2));
      track.setAttribute("stroke-linecap", "round");
      track.setAttribute("transform", transform);
      svg.appendChild(track);

      var arc = createSvg("circle");
      arc.setAttribute("cx", cx);
      arc.setAttribute("cy", cy);
      arc.setAttribute("r", radius);
      arc.setAttribute("fill", "none");
      arc.setAttribute("stroke", colors[index % colors.length]);
      arc.setAttribute("stroke-width", stroke);
      arc.setAttribute("stroke-dasharray", valueLength.toFixed(2) + " " + circumference.toFixed(2));
      arc.setAttribute("stroke-linecap", "round");
      arc.setAttribute("transform", transform);
      attachTooltip(arc, function () {
        return tooltipHtml(item.label, formatCurrencyFull(item.value));
      });
      svg.appendChild(arc);

      var text = createSvg("text");
      text.setAttribute("x", "24");
      text.setAttribute("y", String(68 + index * 22));
      text.setAttribute("font-size", "12");
      text.setAttribute("font-weight", "600");
      text.setAttribute("fill", "#666666");
      text.textContent = item.label + ": " + formatCurrency(item.value);
      svg.appendChild(text);
    });

    wrap.appendChild(svg);
    mount.appendChild(wrap);
  }

  function renderFunnelChart(series, mount) {
    var data = aggregateBySegment(series.rows);
    var wrap = document.createElement("div");
    wrap.className = "w-funnel";
    var svg = createSvg("svg");
    var width = 320;
    var height = 320;
    var centerX = 160;
    var colors = ["#4996b2", "#63abc6", "#102127"];
    var maxValue = Math.max.apply(null, data.map(function (item) { return item.value; }).concat([1]));

    svg.setAttribute("viewBox", "0 0 " + width + " " + height);

    data.forEach(function (item, index) {
      var widthTop = 220 - index * 52;
      var widthBottom = index === data.length - 1 ? widthTop - 34 : 220 - (index + 1) * 52;
      var topY = 30 + index * 86;
      var bottomY = topY + 68;
      var path = createSvg("path");
      path.setAttribute("d", [
        "M", centerX - widthTop / 2, topY,
        "L", centerX + widthTop / 2, topY,
        "L", centerX + widthBottom / 2, bottomY,
        "L", centerX - widthBottom / 2, bottomY,
        "Z"
      ].join(" "));
      path.setAttribute("fill", colors[index % colors.length]);
      attachTooltip(path, function () {
        return tooltipHtml(
          item.label,
          formatCurrencyFull(item.value),
          Math.round((item.value / maxValue) * 100) + "% of top stage"
        );
      });
      svg.appendChild(path);

      var text = createSvg("text");
      text.setAttribute("x", centerX);
      text.setAttribute("y", topY + 37);
      text.setAttribute("text-anchor", "middle");
      text.setAttribute("font-size", "12");
      text.setAttribute("font-weight", "700");
      text.setAttribute("fill", "#ffffff");
      text.textContent = item.label + " " + formatCurrency(item.value);
      svg.appendChild(text);
    });

    wrap.appendChild(svg);
    mount.appendChild(wrap);
  }

  function renderWidgetByPage(series, mount) {
    if (state.currentPage === 0) return renderBarChart(series, mount);
    if (state.currentPage === 1) return renderLineChart(series, mount);
    if (state.currentPage === 2) return renderWaterfallChart(series, mount);
    if (state.currentPage === 3) return renderRadialChart(series, mount);
    return renderFunnelChart(series, mount);
  }

  function buildPager(card) {
    var pager = document.createElement("div");
    pager.className = "pagination";

    var prev = document.createElement("button");
    prev.className = "pag-btn" + (state.currentPage === 0 ? " invisible" : "");
    prev.type = "button";
    prev.setAttribute("aria-label", "Previous page");
    prev.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"></polyline></svg>';
    prev.addEventListener("click", function () {
      if (state.currentPage > 0) {
        state.currentPage -= 1;
        render();
      }
    });

    var indicator = document.createElement("span");
    indicator.className = "pag-indicator";
    indicator.textContent = state.currentPage + 1 + "/" + cardPages.length;

    var next = document.createElement("button");
    next.className = "pag-btn" + (state.currentPage === cardPages.length - 1 ? " invisible" : "");
    next.type = "button";
    next.setAttribute("aria-label", "Next page");
    next.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 6 15 12 9 18"></polyline></svg>';
    next.addEventListener("click", function () {
      if (state.currentPage < cardPages.length - 1) {
        state.currentPage += 1;
        render();
      }
    });

    pager.appendChild(prev);
    pager.appendChild(indicator);
    pager.appendChild(next);
    card.appendChild(pager);
  }

  function renderCard(series) {
    var summary = buildSummary(series);
    if (!summary) return null;

    var card = document.createElement("article");
    card.className = "card";
    card.style.setProperty("--metric-accent", accentByShipMode[series.shipMode] || "#4996b2");

    var header = document.createElement("div");
    header.className = "card-header";
    header.innerHTML =
      '<div class="metric-row">' +
        '<div class="metric-meta">' +
          '<span class="metric-accent"></span>' +
          '<span class="metric-title">Sales - ' + escapeHtml(series.shipMode) + "</span>" +
        "</div>" +
      "</div>" +
      '<div class="divider-line"></div>';

    var viewport = document.createElement("div");
    viewport.className = "carousel-viewport";

    var strip = document.createElement("div");
    strip.className = "carousel-strip";
    strip.style.transform = "translateX(-" + (state.currentPage * 100) + "%)";

    cardPages.forEach(function (page, pageIndex) {
      var pageEl = document.createElement("div");
      pageEl.className = "carousel-page";

      var summarySlot = document.createElement("div");
      summarySlot.className = "widget-slot";
      var summaryWrap = document.createElement("div");
      summaryWrap.className = "w-kpi";

      var topRow = document.createElement("div");
      topRow.className = "w-kpi-top";
      topRow.innerHTML = '<div class="w-kpi-last-date">' + monthLongLabel(summary.selected.date) + "</div>";

      var value = document.createElement("div");
      value.className = "w-kpi-value";
      value.textContent = formatCurrency(summary.selected.sales);

      var deltas = document.createElement("div");
      deltas.className = "w-kpi-deltas";
      deltas.appendChild(buildDeltaRow("vs previous month", summary.previousMonthDelta));
      deltas.appendChild(buildDeltaRow("vs previous year", summary.previousYearDelta));

      summaryWrap.appendChild(topRow);
      summaryWrap.appendChild(value);
      summaryWrap.appendChild(deltas);
      summarySlot.appendChild(summaryWrap);

      var chartSlot = document.createElement("div");
      chartSlot.className = "widget-slot chart-slot";
      if (pageIndex === state.currentPage) renderWidgetByPage(series, chartSlot);

      pageEl.appendChild(summarySlot);
      pageEl.appendChild(chartSlot);
      strip.appendChild(pageEl);
    });

    viewport.appendChild(strip);
    card.appendChild(header);
    card.appendChild(viewport);
    buildPager(card);
    return card;
  }

  function render() {
    hideTooltip();
    var filteredRows = getFilteredRows();
    var seriesList = buildShipModeSeries(filteredRows);
    cardsGridEl.innerHTML = "";
    if (!seriesList.length) {
      cardsGridEl.innerHTML = '<div class="empty-state">No data matches the current filters.</div>';
      return;
    }
    seriesList.forEach(function (series) {
      var card = renderCard(series);
      if (card) cardsGridEl.appendChild(card);
    });
  }

  function hydrateFilters() {
    populateFilter(regionFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.Region; }))).sort(), "Regions");
    populateFilter(categoryFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.Category; }))).sort(), "Categories");
    populateFilter(subCategoryFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.SubCategory; }))).sort(), "Sub-Categories");
    populateFilter(stateFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.State; }))).sort(), "States");

    [regionFilterEl, categoryFilterEl, subCategoryFilterEl, stateFilterEl].forEach(function (el) {
      el.addEventListener("change", render);
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
        cardsGridEl.innerHTML = '<div class="empty-state">Check that <code>sample_-_superstore.xls</code> is present in this folder when you publish the demo.</div>';
      });
  }

  loadWorkbook();
})();
