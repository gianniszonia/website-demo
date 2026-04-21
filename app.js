(function () {
  "use strict";

  var dataPath = "./Sample - Superstore_Orders.csv";
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

  var regionFilterEl = document.getElementById("regionFilter");
  var categoryFilterEl = document.getElementById("categoryFilter");
  var subCategoryFilterEl = document.getElementById("subCategoryFilter");
  var monthFilterEl = document.getElementById("monthFilter");
  var cardsGridEl = document.getElementById("cardsGrid");
  var tooltipEl = document.getElementById("chartTooltip");

  function parseWorkbookDate(value) {
    if (value instanceof Date && !isNaN(value.getTime())) return value;
    if (typeof value === "number" && isFinite(value)) {
      var parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
    if (typeof value === "string" && value.trim()) {
      var trimmed = value.trim();
      var parts = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
      if (parts) {
        var year = Number(parts[3]);
        if (year < 100) year += 2000;
        var manual = new Date(year, Number(parts[1]) - 1, Number(parts[2]));
        if (!isNaN(manual.getTime())) return manual;
      }
      var normalized = new Date(trimmed);
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

  function formatPlainPercent(value) {
    if (value === null || !isFinite(value)) return "-";
    return Math.round(value) + "%";
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

  function monthFilterLabel(date) {
    return date.toLocaleString("en-US", { month: "long" }) + "-" + String(date.getFullYear()).slice(-2);
  }

  function dateSpanLabel(startDate, endDate) {
    if (!startDate && !endDate) return "";
    if (!startDate || !endDate || monthKey(startDate) === monthKey(endDate)) return monthLongLabel(endDate || startDate);
    return monthLongLabel(startDate) + " - " + monthLongLabel(endDate);
  }

  function tooltipRow(label, value, note) {
    var row = '<div style="display:flex;gap:6px;align-items:baseline;flex-wrap:wrap">';
    row += '<span class="tt-label">' + escapeHtml(label) + ":</span>";
    row += '<span style="font-weight:600">' + escapeHtml(value) + "</span>";
    if (note) row += '<span class="tt-label">(' + escapeHtml(note) + ")</span>";
    row += "</div>";
    return row;
  }

  function tooltipCard(title, rows) {
    return '<div style="font-size:13px">' +
      '<div style="margin-bottom:5px;font-weight:600;font-size:13px">' + escapeHtml(title) + "</div>" +
      rows.join("") +
      "</div>";
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

  function populateMonthFilter(months, selectedKey) {
    monthFilterEl.innerHTML = "";
    months.slice().reverse().forEach(function (month) {
      var option = document.createElement("option");
      option.value = month.key;
      option.textContent = monthFilterLabel(month.date);
      if (month.key === selectedKey) option.selected = true;
      monthFilterEl.appendChild(option);
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

  function getBaseFilteredRows() {
    return state.rows.filter(function (row) {
      if (regionFilterEl.value && row.Region !== regionFilterEl.value) return false;
      if (categoryFilterEl.value && row.Category !== categoryFilterEl.value) return false;
      if (subCategoryFilterEl.value && row.SubCategory !== subCategoryFilterEl.value) return false;
      return true;
    });
  }

  function buildGlobalMonths(rows) {
    if (!rows.length) return [];
    var minDate = rows[0].orderDate;
    var maxDate = rows[0].orderDate;
    rows.forEach(function (row) {
      if (row.orderDate < minDate) minDate = row.orderDate;
      if (row.orderDate > maxDate) maxDate = row.orderDate;
    });
    var cursor = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
    var end = new Date(maxDate.getFullYear(), maxDate.getMonth(), 1);
    var months = [];
    while (cursor <= end) {
      months.push({
        key: monthKey(cursor),
        date: new Date(cursor.getFullYear(), cursor.getMonth(), 1)
      });
      cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
    }
    return months;
  }

  function buildMonthlyData(rows, globalMonths) {
    var byMonth = new Map();
    globalMonths.forEach(function (month) {
      byMonth.set(month.key, {
        key: month.key,
        date: new Date(month.date.getFullYear(), month.date.getMonth(), 1),
        sales: 0
      });
    });
    rows.forEach(function (row) {
      var key = monthKey(row.orderDate);
      if (byMonth.has(key)) byMonth.get(key).sales += row.sales;
    });
    return globalMonths.map(function (month) {
      return byMonth.get(month.key);
    });
  }

  function buildShipModeSeries(rows, globalMonths) {
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
          months: buildMonthlyData(entry[1], globalMonths)
        };
      })
      .sort(function (a, b) {
        var aLatest = a.months.length ? a.months[a.months.length - 1].sales : 0;
        var bLatest = b.months.length ? b.months[b.months.length - 1].sales : 0;
        return bLatest - aLatest;
      });
  }

  function buildSummary(series, endKey) {
    if (!series.months.length) return null;
    var selectedIndex = series.months.findIndex(function (month) { return month.key === endKey; });
    if (selectedIndex < 0) return null;
    var selected = series.months[selectedIndex];
    var previousMonth = selectedIndex > 0 ? series.months[selectedIndex - 1] : null;
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

  function getBucketDeltaPct(months, currentKey, periodsBack) {
    var idx = months.findIndex(function (month) { return month.key === currentKey; });
    if (idx < 0 || idx < periodsBack) return null;
    var prev = months[idx - periodsBack];
    var curr = months[idx];
    if (!prev || !curr || prev.sales === 0) return null;
    return ((curr.sales - prev.sales) / Math.abs(prev.sales)) * 100;
  }

  function showTooltip(event, html) {
    tooltipEl.innerHTML = html;
    tooltipEl.classList.add("visible");
    var source = event.currentTarget || event.target;
    var card = source && source.closest ? source.closest(".card") : null;
    var tooltipRect = tooltipEl.getBoundingClientRect();
    var cardRect = card ? card.getBoundingClientRect() : {
      left: 8,
      top: 8,
      right: window.innerWidth - 8,
      bottom: window.innerHeight - 8
    };
    var gap = 12;
    var preferRight = event.clientX <= ((cardRect.left + cardRect.right) / 2);
    var left = preferRight
      ? event.clientX + gap
      : event.clientX - tooltipRect.width - gap;

    if (left < cardRect.left + 8) left = cardRect.left + 8;
    if (left + tooltipRect.width > cardRect.right - 8) left = cardRect.right - tooltipRect.width - 8;

    var top = event.clientY - (tooltipRect.height / 2);
    if (top < cardRect.top + 8) top = cardRect.top + 8;
    if (top + tooltipRect.height > cardRect.bottom - 8) top = cardRect.bottom - tooltipRect.height - 8;

    tooltipEl.style.left = left + "px";
    tooltipEl.style.top = top + "px";
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

  function interpolateHexColor(startHex, endHex, ratio) {
    function parse(hex) {
      var value = hex.replace("#", "");
      if (value.length === 3) value = value.split("").map(function (c) { return c + c; }).join("");
      return {
        r: parseInt(value.slice(0, 2), 16),
        g: parseInt(value.slice(2, 4), 16),
        b: parseInt(value.slice(4, 6), 16)
      };
    }
    function toHex(value) {
      return Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, "0");
    }
    var start = parse(startHex);
    var end = parse(endHex);
    return "#" + toHex(start.r + ((end.r - start.r) * ratio)) + toHex(start.g + ((end.g - start.g) * ratio)) + toHex(start.b + ((end.b - start.b) * ratio));
  }

  function recentMonthsDescending(series, count) {
    return function (endKey) {
      var endIndex = series.months.findIndex(function (month) { return month.key === endKey; });
      if (endIndex < 0) return [];
      return series.months.slice(Math.max(0, endIndex - count + 1), endIndex + 1);
    };
  }

  function renderBarChart(series, mount, endKey) {
    var months = recentMonthsDescending(series, 8)(endKey);
    var wrap = document.createElement("div");
    wrap.className = "w-bars-shell";
    var chart = document.createElement("div");
    chart.className = "w-bars";
    var maxValue = Math.max.apply(null, months.map(function (month) { return month.sales; }).concat([1]));
    var columns = [];

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
      bar.style.top = "0px";
      bar.style.height = "0px";
      plot.appendChild(bar);

      var label = document.createElement("span");
      label.className = "bar-month";
      label.textContent = monthShortLabel(month.date);

      attachTooltip(column, function () {
        return tooltipCard(monthLongLabel(month.date), [
          tooltipRow("Sales", formatCurrencyFull(month.sales)),
          tooltipRow("vs previous month", formatPercent(getBucketDeltaPct(series.months, month.key, 1))),
          tooltipRow("vs previous year", formatPercent(getBucketDeltaPct(series.months, month.key, 12)))
        ]);
      });

      column.appendChild(plot);
      column.appendChild(label);
      chart.appendChild(column);
      columns.push({ plot: plot, value: value, bar: bar, month: month });
    });

    wrap.appendChild(chart);
    mount.appendChild(wrap);

    requestAnimationFrame(function () {
      var chartHeight = chart.clientHeight || 320;
      var valueHeight = 18;
      var monthHeight = 18;
      var plotHeight = Math.max(chartHeight - monthHeight - 8, 120);
      var topPad = valueHeight + 8;
      var usableHeight = Math.max(plotHeight - topPad - 8, 40);
      columns.forEach(function (item) {
        item.plot.style.height = plotHeight + "px";
        var barHeight = Math.max(4, Math.round((item.month.sales / maxValue) * usableHeight));
        var top = topPad + (usableHeight - barHeight);
        item.bar.style.top = top + "px";
        item.bar.style.height = barHeight + "px";
        item.value.style.top = Math.max(0, top - valueHeight - 4) + "px";
        item.value.style.bottom = "auto";
      });
    });
  }

  function renderLineChart(series, mount, endKey) {
    var months = recentMonthsDescending(series, 8)(endKey);
    var wrap = document.createElement("div");
    wrap.className = "w-line";
    mount.appendChild(wrap);

    requestAnimationFrame(function () { requestAnimationFrame(function () {
      var width = wrap.clientWidth || 360;
      var height = wrap.clientHeight || 320;
      var pad = { top: 24, right: 18, bottom: 34, left: 28 };
      var plotWidth = width - pad.left - pad.right;
      var plotHeight = height - pad.top - pad.bottom;
      var maxValue = Math.max.apply(null, months.map(function (month) { return month.sales; }).concat([1]));
      var minValue = Math.min.apply(null, months.map(function (month) { return month.sales; }).concat([0]));
      var range = Math.max(1, maxValue - minValue);

      function runPath(points) {
        if (points.length < 2) return "M" + points[0].x.toFixed(1) + "," + points[0].y.toFixed(1);
        var d = "M" + points[0].x.toFixed(1) + "," + points[0].y.toFixed(1);
        for (var i = 0; i < points.length - 1; i += 1) {
          var p0 = points[i];
          var p1 = points[i + 1];
          var mx = (p0.x + p1.x) / 2;
          d += " C" + mx.toFixed(1) + "," + p0.y.toFixed(1) + " " + mx.toFixed(1) + "," + p1.y.toFixed(1) + " " + p1.x.toFixed(1) + "," + p1.y.toFixed(1);
        }
        return d;
      }

      var points = months.map(function (month, index) {
        return {
          month: month,
          x: pad.left + (plotWidth / Math.max(1, months.length - 1)) * index,
          y: pad.top + plotHeight - (((month.sales - minValue) / range) * plotHeight)
        };
      });

      var pathData = runPath(points);
      var areaData = pathData + " L " + points[points.length - 1].x.toFixed(1) + "," + (height - pad.bottom) + " L " + points[0].x.toFixed(1) + "," + (height - pad.bottom) + " Z";
      var gradientId = "lineGrad" + Math.random().toString(36).slice(2, 8);
      var svgContent = '<svg width="' + width + '" height="' + height + '" style="width:' + width + 'px;height:' + height + 'px;display:block" viewBox="0 0 ' + width + " " + height + '" xmlns="http://www.w3.org/2000/svg">' +
        '<defs><linearGradient id="' + gradientId + '" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#4996b2" stop-opacity="0.28"/><stop offset="100%" stop-color="#4996b2" stop-opacity="0.02"/></linearGradient></defs>' +
        '<path d="' + areaData + '" fill="url(#' + gradientId + ')"/>' +
        '<path d="' + pathData + '" fill="none" stroke="#4996b2" stroke-width="2.25" stroke-linecap="round" stroke-linejoin="round"/>';

      points.forEach(function (point) {
        var px = Math.round(point.x);
        var py = Math.round(point.y);
        svgContent += '<circle cx="' + px + '" cy="' + py + '" r="4" fill="#4996b2"/>';
        svgContent += '<text class="widget-label" x="' + px + '" y="' + (height - 8) + '" font-size="15" text-anchor="middle">' + monthShortLabel(point.month.date) + "</text>";
        svgContent += '<text class="widget-label" x="' + px + '" y="' + Math.round(point.y - 18) + '" font-size="22" text-anchor="middle">' + escapeHtml(formatCurrency(point.month.sales)) + "</text>";
        svgContent += '<circle class="line-hit" data-idx="' + points.indexOf(point) + '" cx="' + px + '" cy="' + py + '" r="12" fill="transparent"/>';
      });

      svgContent += "</svg>";
      wrap.innerHTML = svgContent;

      Array.prototype.forEach.call(wrap.querySelectorAll(".line-hit"), function (hit) {
        var idx = parseInt(hit.getAttribute("data-idx"), 10);
        var point = points[idx];
        attachTooltip(hit, function () {
          return tooltipCard(dateSpanLabel(point.month.date, point.month.date), [
            tooltipRow("Sales", formatCurrencyFull(point.month.sales)),
            tooltipRow("vs previous month", formatPercent(getBucketDeltaPct(series.months, point.month.key, 1))),
            tooltipRow("vs previous year", formatPercent(getBucketDeltaPct(series.months, point.month.key, 12)))
          ]);
        });
      });
    }); });
  }

  function renderWaterfallChart(series, mount, endKey) {
    var months = recentMonthsDescending(series, 8)(endKey);
    var fullMonths = series.months;
    var varianceSeries = months.map(function (month) {
      var absoluteIndex = fullMonths.findIndex(function (item) { return item.key === month.key; });
      var previous = absoluteIndex > 0 ? fullMonths[absoluteIndex - 1] : null;
      var prevSales = previous ? previous.sales : 0;
      return {
        key: month.key,
        date: month.date,
        sales: month.sales,
        delta: month.sales - prevSales,
        start: prevSales,
        end: month.sales,
        isTotal: false
      };
    });

    var wrap = document.createElement("div");
    wrap.className = "w-bars-shell";
    var chart = document.createElement("div");
    chart.className = "w-bars";
    var allValues = [];
    varianceSeries.forEach(function (item) {
      allValues.push(item.start, item.end, 0);
    });
    var maxValue = Math.max.apply(null, allValues.concat([1]));
    var minValue = Math.min.apply(null, allValues.concat([0]));
    if (maxValue === minValue) {
      maxValue += 1;
      minValue -= 1;
    }
    var range = maxValue - minValue;

    varianceSeries.forEach(function (item) {
      var column = document.createElement("div");
      column.className = "bar-col";

      var plot = document.createElement("div");
      plot.className = "bar-plot";

      var value = document.createElement("span");
      value.className = "bar-val";
      value.textContent = formatCurrency(item.sales);
      plot.appendChild(value);

      var bar = document.createElement("div");
      bar.className = "bar";
      bar.style.background = item.delta >= 0 ? "#22c55e" : "#ef4444";
      bar.style.top = "0px";
      bar.style.height = "0px";
      plot.appendChild(bar);

      var connector = document.createElement("div");
      connector.style.position = "absolute";
      connector.style.height = "0";
      connector.style.borderTop = "1px dashed #afafaf";
      connector.style.left = "90%";
      connector.style.width = "0";
      connector.style.pointerEvents = "none";
      connector.style.zIndex = "4";
      plot.appendChild(connector);

      var label = document.createElement("span");
      label.className = "bar-month";
      label.textContent = monthShortLabel(item.date);

      attachTooltip(column, function () {
        return tooltipCard(monthLongLabel(item.date), [
          tooltipRow("Sales", formatCurrencyFull(item.sales)),
          tooltipRow("Change", formatCurrencyFull(item.delta)),
          tooltipRow("vs previous month", formatPercent(getBucketDeltaPct(series.months, monthKey(item.date), 1)))
        ]);
      });

      column.appendChild(plot);
      column.appendChild(label);
      chart.appendChild(column);
      column.__wf = { plot: plot, value: value, bar: bar, connector: connector, item: item };
    });

    wrap.appendChild(chart);
    mount.appendChild(wrap);

    requestAnimationFrame(function () {
      var columns = Array.prototype.slice.call(chart.children).map(function (child) { return child.__wf; }).filter(Boolean);
      var chartHeight = chart.clientHeight || 320;
      var valueHeight = 18;
      var monthHeight = 18;
      var plotHeight = Math.max(chartHeight - monthHeight - 8, 120);
      var topPad = valueHeight + 8;
      var bottomPad = minValue < 0 ? valueHeight + 8 : 8;
      var usableHeight = Math.max(plotHeight - topPad - bottomPad, 40);
      var unit = usableHeight / range;
      columns.forEach(function (entry, idx) {
        entry.plot.style.height = plotHeight + "px";
        var topVal = Math.max(entry.item.start, entry.item.end);
        var bottomVal = Math.min(entry.item.start, entry.item.end);
        var top = topPad + ((maxValue - topVal) * unit);
        var bottom = topPad + ((maxValue - bottomVal) * unit);
        var barHeight = Math.max(4, Math.round(bottom - top));
        entry.bar.style.top = top + "px";
        entry.bar.style.height = barHeight + "px";
        if (entry.item.delta >= 0) {
          entry.value.style.top = Math.max(0, top - valueHeight - 4) + "px";
        } else {
          entry.value.style.top = (top + barHeight + 4) + "px";
        }
        entry.value.style.bottom = "auto";
        if (entry.connector) {
          var nextEntry = idx < columns.length - 1 ? columns[idx + 1] : null;
          if (nextEntry) {
            var currentLevel = topPad + ((maxValue - entry.item.end) * unit);
            entry.connector.style.top = currentLevel + "px";
            entry.connector.style.width = Math.max(0, entry.plot.clientWidth * 0.2 + 8) + "px";
          } else {
            entry.connector.style.width = "0";
          }
        }
      });
    });
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

  function renderRadialChart(series, mount, endKey) {
    var data = aggregateBySegment(series.rows.filter(function (row) {
      return monthKey(row.orderDate) === endKey;
    }));
    var wrap = document.createElement("div");
    wrap.className = "w-radial";
    mount.appendChild(wrap);

    requestAnimationFrame(function () { requestAnimationFrame(function () {
      var width = wrap.clientWidth || 320;
      var height = wrap.clientHeight || 320;
      var svg = createSvg("svg");
      var stroke = 12;
      var stride = stroke + 12;
      var labelAreaW = 162;
      var outerRadius = Math.min(Math.floor((width - labelAreaW - 32) / 2), Math.floor((height - 24) / 2));
      var layoutWidth = labelAreaW + (outerRadius * 2);
      var layoutLeft = Math.max(12, Math.round((width - layoutWidth) / 2));
      var cx = layoutLeft + labelAreaW + outerRadius - 34;
      var cy = Math.round(height / 2);
      var maxValue = Math.max.apply(null, data.map(function (item) { return item.value; }).concat([1]));
      var totalValue = data.reduce(function (sum, item) { return sum + item.value; }, 0);
      var colors = ["#4996b2", "#22c55e", "#ef4444"];

      svg.setAttribute("width", "100%");
      svg.setAttribute("height", "100%");
      svg.setAttribute("viewBox", "0 0 " + width + " " + height);

      data.forEach(function (item, index) {
        var radius = outerRadius - index * stride;
        var circumference = 2 * Math.PI * radius;
        var trackLength = circumference * 0.75;
        var valueLength = Math.max(0, (item.value / maxValue) * trackLength);
        var transform = "rotate(-90 " + cx + " " + cy + ")";
        var color = colors[index % colors.length];

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
        arc.setAttribute("stroke", color);
        arc.setAttribute("stroke-width", stroke);
        arc.setAttribute("stroke-dasharray", valueLength.toFixed(2) + " " + circumference.toFixed(2));
        arc.setAttribute("stroke-linecap", "round");
        arc.setAttribute("transform", transform);
        attachTooltip(arc, function () {
          var selectedMonth = series.months.find(function (month) { return month.key === endKey; });
          return tooltipCard(dateSpanLabel(selectedMonth && selectedMonth.date, selectedMonth && selectedMonth.date), [
            tooltipRow("Segment", item.label),
            tooltipRow("Sales", formatCurrencyFull(item.value)),
            tooltipRow("% of total", formatPlainPercent((item.value / totalValue) * 100))
          ]);
        });
        svg.appendChild(arc);

        // Label at arc's left opening — dot placed outside the stroke, text to its left
        var startX = cx;
        var startY = cy - radius;
        var labelY = Math.round(startY);
        var dotCx = Math.round(startX - stroke / 2 - 26);

        var dot = createSvg("circle");
        dot.setAttribute("cx", String(dotCx));
        dot.setAttribute("cy", String(labelY));
        dot.setAttribute("r", "4.5");
        dot.setAttribute("fill", color);
        dot.setAttribute("pointer-events", "none");
        svg.appendChild(dot);

        var text = createSvg("text");
        text.setAttribute("x", String(dotCx - 20));
        text.setAttribute("y", String(labelY));
        text.setAttribute("text-anchor", "end");
        text.setAttribute("dominant-baseline", "middle");
        text.setAttribute("class", "widget-label");
        text.setAttribute("font-size", "20");
        text.textContent = item.label + ": " + formatCurrency(item.value);
        svg.appendChild(text);
      });

      wrap.innerHTML = "";
      wrap.appendChild(svg);
    }); });
  }

  function renderFunnelChart(series, mount, endKey) {
    var data = aggregateBySegment(series.rows.filter(function (row) {
      return monthKey(row.orderDate) === endKey;
    }));
    var wrap = document.createElement("div");
    wrap.className = "w-funnel";
    mount.appendChild(wrap);

    requestAnimationFrame(function () {
      var width = wrap.clientWidth || 320;
      var height = wrap.clientHeight || 320;
      var pad = 12;
      var centerX = width / 2;
      var maxValue = Math.max.apply(null, data.map(function (item) { return item.value; }).concat([1]));
      var segCount = data.length;
      var mainSpan = height - (pad * 2);
      var crossSpan = width - (pad * 2);
      var segSize = mainSpan / Math.max(segCount, 1);
      var svg = createSvg("svg");
      svg.setAttribute("viewBox", "0 0 " + width + " " + height);

      data.forEach(function (item, index) {
        var nextItem = data[Math.min(index + 1, segCount - 1)];
        var curRatio = Math.max(0.12, Math.abs(item.value) / maxValue);
        var nextRatio = Math.max(0.12, Math.abs(nextItem.value) / maxValue);
        var curCross = crossSpan * curRatio;
        var nextCross = crossSpan * nextRatio;
        var topY = pad + (index * segSize);
        var bottomY = pad + ((index + 1) * segSize);
        var leftTop = centerX - (curCross / 2);
        var rightTop = centerX + (curCross / 2);
        var leftBottom = centerX - (nextCross / 2);
        var rightBottom = centerX + (nextCross / 2);
        var curve = Math.min(18, segSize * 0.38);
        var midY = (topY + bottomY) / 2;
        var color = interpolateHexColor("#4996b2", "#8fd3e8", segCount === 1 ? 0 : (index / (segCount - 1)));

        var path = createSvg("path");
        var d =
          "M " + leftTop.toFixed(1) + " " + topY.toFixed(1) +
          " C " + (leftTop + curve).toFixed(1) + " " + (topY - curve * 0.45).toFixed(1) +
          " " + (rightTop - curve).toFixed(1) + " " + (topY - curve * 0.45).toFixed(1) +
          " " + rightTop.toFixed(1) + " " + topY.toFixed(1) +
          " C " + (rightTop - curve * 0.25).toFixed(1) + " " + midY.toFixed(1) +
          " " + (rightBottom + curve * 0.25).toFixed(1) + " " + midY.toFixed(1) +
          " " + rightBottom.toFixed(1) + " " + bottomY.toFixed(1) +
          " C " + (rightBottom - curve).toFixed(1) + " " + (bottomY + curve * 0.45).toFixed(1) +
          " " + (leftBottom + curve).toFixed(1) + " " + (bottomY + curve * 0.45).toFixed(1) +
          " " + leftBottom.toFixed(1) + " " + bottomY.toFixed(1) +
          " C " + (leftBottom - curve * 0.25).toFixed(1) + " " + midY.toFixed(1) +
          " " + (leftTop + curve * 0.25).toFixed(1) + " " + midY.toFixed(1) +
          " " + leftTop.toFixed(1) + " " + topY.toFixed(1) + " Z";
        path.setAttribute("d", d);
        path.setAttribute("fill", color);
        path.setAttribute("stroke", color);
        path.setAttribute("stroke-width", "1.5");
        path.setAttribute("stroke-linejoin", "round");
        attachTooltip(path, function () {
          var selectedMonth = series.months.find(function (month) { return month.key === endKey; });
          return tooltipCard(dateSpanLabel(selectedMonth && selectedMonth.date, selectedMonth && selectedMonth.date), [
            tooltipRow("Stage", item.label),
            tooltipRow("Sales", formatCurrencyFull(item.value)),
            tooltipRow("% of first", Math.round((item.value / maxValue) * 100) + "%")
          ]);
        });
        svg.appendChild(path);

        var text = createSvg("text");
        text.setAttribute("x", centerX);
        text.setAttribute("y", topY + (segSize / 2));
        text.setAttribute("text-anchor", "middle");
        text.setAttribute("dominant-baseline", "middle");
        text.setAttribute("font-size", "10");
        text.setAttribute("font-weight", "600");
        text.setAttribute("fill", "#333333");
        text.textContent = item.label + " | " + formatCurrency(item.value);
        svg.appendChild(text);
      });

      wrap.innerHTML = "";
      wrap.appendChild(svg);
    });
  }

  function renderWidgetByPage(pageIndex, series, mount, endKey) {
    if (pageIndex === 0) return renderBarChart(series, mount, endKey);
    if (pageIndex === 1) return renderLineChart(series, mount, endKey);
    if (pageIndex === 2) return renderWaterfallChart(series, mount, endKey);
    if (pageIndex === 3) return renderRadialChart(series, mount, endKey);
    return renderFunnelChart(series, mount, endKey);
  }

  function applyPaginationState(card) {
    var strip = card.querySelector(".carousel-strip");
    var prev = card.querySelector(".pag-prev");
    var next = card.querySelector(".pag-next");
    var indicator = card.querySelector(".pag-indicator");
    if (strip) strip.style.transform = "translateX(-" + (state.currentPage * 100) + "%)";
    if (prev) prev.classList.toggle("invisible", state.currentPage === 0);
    if (next) next.classList.toggle("invisible", state.currentPage === cardPages.length - 1);
    if (indicator) indicator.textContent = state.currentPage + 1 + "/" + cardPages.length;
  }

  function updatePaginationAcrossCards() {
    Array.prototype.forEach.call(cardsGridEl.querySelectorAll(".card"), applyPaginationState);
  }

  function buildPager(card) {
    var pager = document.createElement("div");
    pager.className = "pagination";

    var prev = document.createElement("button");
    prev.className = "pag-btn pag-prev" + (state.currentPage === 0 ? " invisible" : "");
    prev.type = "button";
    prev.setAttribute("aria-label", "Previous page");
    prev.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"></polyline></svg>';
    prev.addEventListener("click", function () {
      if (state.currentPage > 0) {
        state.currentPage -= 1;
        updatePaginationAcrossCards();
      }
    });

    var indicator = document.createElement("span");
    indicator.className = "pag-indicator";
    indicator.textContent = state.currentPage + 1 + "/" + cardPages.length;

    var next = document.createElement("button");
    next.className = "pag-btn pag-next" + (state.currentPage === cardPages.length - 1 ? " invisible" : "");
    next.type = "button";
    next.setAttribute("aria-label", "Next page");
    next.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 6 15 12 9 18"></polyline></svg>';
    next.addEventListener("click", function () {
      if (state.currentPage < cardPages.length - 1) {
        state.currentPage += 1;
        updatePaginationAcrossCards();
      }
    });

    pager.appendChild(prev);
    pager.appendChild(indicator);
    pager.appendChild(next);
    card.appendChild(pager);
  }

  function renderCard(series, endKey) {
    var summary = buildSummary(series, endKey);
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
      renderWidgetByPage(pageIndex, series, chartSlot, endKey);

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
    var filteredRows = getBaseFilteredRows();
    var globalMonths = buildGlobalMonths(filteredRows);
    cardsGridEl.innerHTML = "";
    if (!globalMonths.length) {
      cardsGridEl.innerHTML = '<div class="empty-state">No data matches the current filters.</div>';
      return;
    }
    if (!state.currentEndKey || !globalMonths.some(function (month) { return month.key === state.currentEndKey; })) {
      state.currentEndKey = globalMonths[globalMonths.length - 1].key;
    }
    populateMonthFilter(globalMonths, state.currentEndKey);
    var cutoffRows = filteredRows.filter(function (row) {
      return monthKey(row.orderDate) <= state.currentEndKey;
    });
    var activeMonths = globalMonths.filter(function (month) {
      return month.key <= state.currentEndKey;
    });
    var seriesList = buildShipModeSeries(cutoffRows, activeMonths);
    if (!seriesList.length) {
      cardsGridEl.innerHTML = '<div class="empty-state">No data matches the current filters.</div>';
      return;
    }
    seriesList.forEach(function (series) {
      var card = renderCard(series, state.currentEndKey);
      if (card) cardsGridEl.appendChild(card);
    });
  }

  function hydrateFilters() {
    populateFilter(regionFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.Region; }))).sort(), "Regions");
    populateFilter(categoryFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.Category; }))).sort(), "Categories");
    populateFilter(subCategoryFilterEl, Array.from(new Set(state.rows.map(function (row) { return row.SubCategory; }))).sort(), "Sub-Categories");

    [regionFilterEl, categoryFilterEl, subCategoryFilterEl].forEach(function (el) {
      el.addEventListener("change", render);
    });
    monthFilterEl.addEventListener("change", function () {
      state.currentEndKey = monthFilterEl.value;
      render();
    });
  }

  function loadWorkbook() {
    fetch(dataPath)
      .then(function (response) {
        if (!response.ok) throw new Error("Could not load sample data.");
        return response.text();
      })
      .then(function (text) {
        var workbook = XLSX.read(text, { type: "string", raw: true, cellDates: true });
        var firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        var rawRows = XLSX.utils.sheet_to_json(firstSheet, {
          raw: true,
          defval: ""
        });
        state.rows = normalizeRows(rawRows);
        hydrateFilters();
        render();
      })
      .catch(function (error) {
        console.error(error);
        cardsGridEl.innerHTML = '<div class="empty-state">Check that <code>Sample - Superstore_Orders.csv</code> is present in this folder when you publish the demo.</div>';
      });
  }

  loadWorkbook();
})();
