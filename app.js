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
    currentPage: 0
  };

  var dataStatusEl = document.getElementById("dataStatus");
  var regionFilterEl = document.getElementById("regionFilter");
  var categoryFilterEl = document.getElementById("categoryFilter");
  var subCategoryFilterEl = document.getElementById("subCategoryFilter");
  var stateFilterEl = document.getElementById("stateFilter");
  var cardsGridEl = document.getElementById("cardsGrid");
  var tooltipEl = document.getElementById("chartTooltip");
  var pageButtons = Array.prototype.slice.call(document.querySelectorAll(".page-btn"));

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

  function formatCurrencyFull(value) {
    return new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
      maximumFractionDigits: 0
    }).format(value || 0);
  }

  function formatDelta(delta) {
    if (delta === null || !isFinite(delta)) return "—";
    var rounded = Math.round(delta);
    if (Object.is(rounded, -0)) rounded = 0;
    return (rounded > 0 ? "+" : "") + rounded + "%";
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
        return a.shipMode.localeCompare(b.shipMode);
      });
  }

  function buildSummary(series) {
    var months = series.months;
    if (!months.length) return null;
    var selected = months[months.length - 1];
    var previousMonth = months.length > 1 ? months[months.length - 2] : null;
    var previousYear = months.length > 12 ? months[months.length - 13] : null;
    var previousMonthDelta = previousMonth && previousMonth.sales !== 0
      ? ((selected.sales - previousMonth.sales) / Math.abs(previousMonth.sales)) * 100
      : null;
    var previousYearDelta = previousYear && previousYear.sales !== 0
      ? ((selected.sales - previousYear.sales) / Math.abs(previousYear.sales)) * 100
      : null;
    return {
      selected: selected,
      previousMonthDelta: previousMonthDelta,
      previousYearDelta: previousYearDelta
    };
  }

  function showTooltip(event, html) {
    tooltipEl.innerHTML = html;
    tooltipEl.classList.add("visible");
    tooltipEl.style.left = (event.clientX + 14) + "px";
    tooltipEl.style.top = (event.clientY + 14) + "px";
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
    var pill = document.createElement("span");
    pill.className = "delta-pill";
    if (deltaValue === null || !isFinite(deltaValue)) pill.classList.add("neutral");
    else if (deltaValue >= 0) pill.classList.add("positive");
    else pill.classList.add("negative");
    pill.textContent = formatDelta(deltaValue);
    var text = document.createElement("span");
    text.className = "delta-label";
    text.textContent = label;
    row.appendChild(pill);
    row.appendChild(text);
    return row;
  }

  function renderBarChart(series, mount) {
    var months = series.months.slice(-8);
    var maxValue = Math.max.apply(null, months.map(function (month) { return month.sales; }));
    var chart = document.createElement("div");
    chart.className = "bar-chart";
    months.forEach(function (month) {
      var column = document.createElement("div");
      column.className = "chart-column";
      var value = document.createElement("div");
      value.className = "chart-value";
      value.textContent = formatCurrencyK(month.sales);
      var plot = document.createElement("div");
      plot.className = "chart-plot";
      var fill = document.createElement("div");
      fill.className = "bar-fill";
      fill.style.height = Math.max(10, Math.round((month.sales / maxValue) * 250)) + "px";
      var label = document.createElement("div");
      label.className = "chart-label";
      label.textContent = monthShortLabel(month.date);
      plot.appendChild(fill);
      column.appendChild(value);
      column.appendChild(plot);
      column.appendChild(label);
      attachTooltip(fill, function () {
        return "<strong>" + monthLongLabel(month.date) + "</strong><br>Sales: " + formatCurrencyFull(month.sales);
      });
      chart.appendChild(column);
    });
    mount.appendChild(chart);
  }

  function renderLineChart(series, mount) {
    var months = series.months.slice(-12);
    var svgWidth = 320;
    var svgHeight = 360;
    var padding = { top: 20, right: 16, bottom: 28, left: 10 };
    var maxValue = Math.max.apply(null, months.map(function (month) { return month.sales; }));
    var minValue = Math.min.apply(null, months.map(function (month) { return month.sales; }));
    var range = Math.max(1, maxValue - minValue);
    var width = svgWidth - padding.left - padding.right;
    var height = svgHeight - padding.top - padding.bottom;
    var points = months.map(function (month, index) {
      return {
        x: padding.left + (width / Math.max(1, months.length - 1)) * index,
        y: padding.top + height - (((month.sales - minValue) / range) * (height - 20)),
        month: month
      };
    });
    var path = points.map(function (point, index) {
      return (index === 0 ? "M" : "L") + point.x.toFixed(1) + " " + point.y.toFixed(1);
    }).join(" ");
    var wrap = document.createElement("div");
    wrap.className = "line-chart";
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 " + svgWidth + " " + svgHeight);
    var area = document.createElementNS("http://www.w3.org/2000/svg", "path");
    area.setAttribute("d", path + " L " + points[points.length - 1].x + " " + (svgHeight - padding.bottom) + " L " + points[0].x + " " + (svgHeight - padding.bottom) + " Z");
    area.setAttribute("fill", "rgba(73, 150, 178, 0.12)");
    var line = document.createElementNS("http://www.w3.org/2000/svg", "path");
    line.setAttribute("d", path);
    line.setAttribute("fill", "none");
    line.setAttribute("stroke", "#4996b2");
    line.setAttribute("stroke-width", "3");
    line.setAttribute("stroke-linecap", "round");
    svg.appendChild(area);
    svg.appendChild(line);
    points.forEach(function (point) {
      var dot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      dot.setAttribute("cx", point.x);
      dot.setAttribute("cy", point.y);
      dot.setAttribute("r", "4.5");
      dot.setAttribute("fill", "#4996b2");
      dot.setAttribute("class", "line-dot");
      attachTooltip(dot, function () {
        return "<strong>" + monthLongLabel(point.month.date) + "</strong><br>Sales: " + formatCurrencyFull(point.month.sales);
      });
      svg.appendChild(dot);
      var label = document.createElementNS("http://www.w3.org/2000/svg", "text");
      label.setAttribute("x", point.x);
      label.setAttribute("y", svgHeight - 8);
      label.setAttribute("text-anchor", "middle");
      label.setAttribute("font-size", "10");
      label.setAttribute("font-weight", "600");
      label.setAttribute("fill", "#666666");
      label.textContent = monthShortLabel(point.month.date);
      svg.appendChild(label);
    });
    wrap.appendChild(svg);
    mount.appendChild(wrap);
  }

  function renderWaterfallChart(series, mount) {
    var months = series.months.slice(-8);
    var chart = document.createElement("div");
    chart.className = "waterfall-chart";
    var running = 0;
    var maxAbs = Math.max.apply(null, months.map(function (month) {
      return Math.abs(month.sales);
    }));
    months.forEach(function (month, index) {
      running += month.sales;
      var column = document.createElement("div");
      column.className = "chart-column";
      var value = document.createElement("div");
      value.className = "chart-value";
      value.textContent = formatCurrencyK(month.sales);
      var plot = document.createElement("div");
      plot.className = "chart-plot";
      var fill = document.createElement("div");
      fill.className = "waterfall-fill" + (month.sales < 0 ? " negative" : "");
      fill.style.height = Math.max(10, Math.round((Math.abs(month.sales) / maxAbs) * 220)) + "px";
      var label = document.createElement("div");
      label.className = "chart-label";
      label.textContent = monthShortLabel(month.date);
      plot.appendChild(fill);
      column.appendChild(value);
      column.appendChild(plot);
      column.appendChild(label);
      attachTooltip(fill, function () {
        return "<strong>" + monthLongLabel(month.date) + "</strong><br>Sales: " + formatCurrencyFull(month.sales) + "<br>Running total: " + formatCurrencyFull(running);
      });
      chart.appendChild(column);
      if (index === months.length - 1) {
        var totalColumn = document.createElement("div");
        totalColumn.className = "chart-column";
        var totalValue = document.createElement("div");
        totalValue.className = "chart-value";
        totalValue.textContent = formatCurrencyK(running);
        var totalPlot = document.createElement("div");
        totalPlot.className = "chart-plot";
        var totalFill = document.createElement("div");
        totalFill.className = "waterfall-fill total";
        totalFill.style.height = Math.max(10, Math.round((Math.abs(running) / Math.max(Math.abs(running), maxAbs)) * 220)) + "px";
        var totalLabel = document.createElement("div");
        totalLabel.className = "chart-label";
        totalLabel.textContent = "TOTAL";
        totalPlot.appendChild(totalFill);
        totalColumn.appendChild(totalValue);
        totalColumn.appendChild(totalPlot);
        totalColumn.appendChild(totalLabel);
        attachTooltip(totalFill, function () {
          return "<strong>Total</strong><br>Sales: " + formatCurrencyFull(running);
        });
        chart.appendChild(totalColumn);
      }
    });
    mount.appendChild(chart);
  }

  function aggregateBySegment(rows) {
    var map = new Map();
    rows.forEach(function (row) {
      if (!map.has(row.Segment)) map.set(row.Segment, 0);
      map.set(row.Segment, map.get(row.Segment) + row.sales);
    });
    return Array.from(map.entries()).map(function (entry) {
      return { label: entry[0], value: entry[1] };
    }).sort(function (a, b) { return b.value - a.value; });
  }

  function renderRadialChart(series, mount) {
    var data = aggregateBySegment(series.rows);
    var wrap = document.createElement("div");
    wrap.className = "radial-chart";
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 320 360");
    var cx = 180;
    var cy = 220;
    var baseRadius = 88;
    var maxValue = Math.max.apply(null, data.map(function (item) { return item.value; }));
    data.forEach(function (item, index) {
      var radius = baseRadius - (index * 26);
      var circumference = 2 * Math.PI * radius;
      var visible = Math.max(0.12, item.value / maxValue) * circumference * 0.75;
      var track = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      track.setAttribute("cx", cx);
      track.setAttribute("cy", cy);
      track.setAttribute("r", radius);
      track.setAttribute("fill", "none");
      track.setAttribute("stroke", "#eeeeee");
      track.setAttribute("stroke-width", "12");
      track.setAttribute("stroke-dasharray", (circumference * 0.75) + " " + circumference);
      track.setAttribute("transform", "rotate(-135 " + cx + " " + cy + ")");
      svg.appendChild(track);
      var bar = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      bar.setAttribute("cx", cx);
      bar.setAttribute("cy", cy);
      bar.setAttribute("r", radius);
      bar.setAttribute("fill", "none");
      bar.setAttribute("stroke", ["#4996b2", "#22c55e", "#ef4444"][index % 3]);
      bar.setAttribute("stroke-width", "12");
      bar.setAttribute("stroke-linecap", "round");
      bar.setAttribute("stroke-dasharray", visible + " " + circumference);
      bar.setAttribute("transform", "rotate(-135 " + cx + " " + cy + ")");
      bar.setAttribute("class", "radial-hit");
      attachTooltip(bar, function () {
        return "<strong>" + item.label + "</strong><br>Sales: " + formatCurrencyFull(item.value);
      });
      svg.appendChild(bar);
      var label = document.createElementNS("http://www.w3.org/2000/svg", "text");
      label.setAttribute("x", 26);
      label.setAttribute("y", 72 + (index * 22));
      label.setAttribute("font-size", "12");
      label.setAttribute("font-weight", "600");
      label.setAttribute("fill", "#666666");
      label.textContent = item.label + ": " + formatCurrencyK(item.value);
      svg.appendChild(label);
    });
    wrap.appendChild(svg);
    mount.appendChild(wrap);
  }

  function renderFunnelChart(series, mount) {
    var data = aggregateBySegment(series.rows);
    var wrap = document.createElement("div");
    wrap.className = "funnel-chart";
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 320 360");
    var colors = ["#4996b2", "#63abc6", "#102127"];
    var maxValue = Math.max.apply(null, data.map(function (item) { return item.value; }));
    data.forEach(function (item, index) {
      var width = 200 - (index * 48);
      var nextWidth = index === data.length - 1 ? width - 34 : 200 - ((index + 1) * 48);
      var topY = 42 + (index * 88);
      var bottomY = topY + 68;
      var centerX = 160;
      var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      var d = [
        "M", centerX - (width / 2), topY,
        "L", centerX + (width / 2), topY,
        "L", centerX + (nextWidth / 2), bottomY,
        "L", centerX - (nextWidth / 2), bottomY,
        "Z"
      ].join(" ");
      path.setAttribute("d", d);
      path.setAttribute("fill", colors[index % colors.length]);
      path.setAttribute("class", "funnel-segment");
      attachTooltip(path, function () {
        return "<strong>" + item.label + "</strong><br>Sales: " + formatCurrencyFull(item.value) + "<br>% of top stage: " + formatDelta((item.value / maxValue) * 100);
      });
      svg.appendChild(path);
      var label = document.createElementNS("http://www.w3.org/2000/svg", "text");
      label.setAttribute("x", centerX);
      label.setAttribute("y", topY + 34);
      label.setAttribute("text-anchor", "middle");
      label.setAttribute("font-size", "12");
      label.setAttribute("font-weight", "700");
      label.setAttribute("fill", "#ffffff");
      label.textContent = item.label + " " + formatCurrencyK(item.value);
      svg.appendChild(label);
    });
    wrap.appendChild(svg);
    mount.appendChild(wrap);
  }

  function renderPageWidget(series, mount) {
    if (state.currentPage === 0) return renderBarChart(series, mount);
    if (state.currentPage === 1) return renderLineChart(series, mount);
    if (state.currentPage === 2) return renderWaterfallChart(series, mount);
    if (state.currentPage === 3) return renderRadialChart(series, mount);
    return renderFunnelChart(series, mount);
  }

  function renderCard(series) {
    var summary = buildSummary(series);
    if (!summary) return null;
    var card = document.createElement("article");
    card.className = "split-card";
    var accentColor = accentByShipMode[series.shipMode] || "#4996b2";

    var header = document.createElement("div");
    header.className = "card-header";
    header.innerHTML =
      '<div class="metric-row">' +
        '<div class="metric-meta">' +
          '<span class="metric-accent" style="background:' + accentColor + '"></span>' +
          '<span class="metric-title">Sales – ' + series.shipMode + '</span>' +
        '</div>' +
        '<span class="metric-period">' + monthLongLabel(summary.selected.date) + '</span>' +
      '</div>' +
      '<div class="divider-line"></div>';

    var body = document.createElement("div");
    body.className = "card-body";

    var summaryTop = document.createElement("div");
    summaryTop.className = "summary-top";
    summaryTop.innerHTML =
      '<div class="summary-value">' + formatCurrencyK(summary.selected.sales) + '</div>' +
      '<div class="summary-date">' + monthLongLabel(summary.selected.date) + '</div>';

    var deltas = document.createElement("div");
    deltas.className = "summary-deltas";
    deltas.appendChild(buildDeltaRow("vs previous month", summary.previousMonthDelta));
    deltas.appendChild(buildDeltaRow("vs previous year", summary.previousYearDelta));

    var widgetArea = document.createElement("div");
    widgetArea.className = "widget-area";
    renderPageWidget(series, widgetArea);

    body.appendChild(summaryTop);
    body.appendChild(deltas);
    body.appendChild(widgetArea);
    card.appendChild(header);
    card.appendChild(body);
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
    pageButtons.forEach(function (button) {
      button.classList.toggle("active", Number(button.dataset.page) === state.currentPage);
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
    pageButtons.forEach(function (button) {
      button.addEventListener("click", function () {
        state.currentPage = Number(button.dataset.page);
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
        cardsGridEl.innerHTML = '<div class="empty-state">Check that <code>sample_-_superstore.xls</code> is present in this folder when you publish the demo.</div>';
      });
  }

  loadWorkbook();
})();
