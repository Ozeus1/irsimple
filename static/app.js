function parseBRL(text) {
  if (!text) return 0;
  return Number(String(text).replace(/[R$\s.]/g, "").replace(",", ".")) || 0;
}

function drawWealthChart() {
  const canvas = document.getElementById("wealthChart");
  const table = document.getElementById("wealthData");
  if (!canvas || !table) return;
  const rows = Array.from(table.querySelectorAll("tbody tr")).map((tr) => {
    const cells = Array.from(tr.children).map((td) => td.textContent.trim());
    return {
      periodo: cells[0],
      patrimonio: parseBRL(cells[1]),
      ibovespa: parseBRL(cells[2]),
      selic: parseBRL(cells[3]),
      poupanca: parseBRL(cells[4]),
    };
  });
  if (!rows.length) return;
  const rect = canvas.getBoundingClientRect();
  canvas.width = Math.max(600, rect.width);
  canvas.height = 280;
  const ctx = canvas.getContext("2d");
  const margin = { left: 72, top: 24, right: 20, bottom: 46 };
  const w = canvas.width - margin.left - margin.right;
  const h = canvas.height - margin.top - margin.bottom;
  const series = [
    ["Patrimonio", "patrimonio", "#2563eb"],
    ["Ibovespa", "ibovespa", "#f97316"],
    ["Selic", "selic", "#16a34a"],
    ["Poupanca", "poupanca", "#9333ea"],
  ];
  let values = [];
  rows.forEach((row) => series.forEach((s) => { if (row[s[1]] > 0) values.push(row[s[1]]); }));
  const min = Math.min(...values, 0);
  const max = Math.max(...values, 1);
  const scaleY = (value) => margin.top + h - ((value - min) / (max - min || 1)) * h;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.strokeStyle = "#cbd5e1";
  ctx.strokeRect(margin.left, margin.top, w, h);
  ctx.font = "11px Segoe UI";
  ctx.fillStyle = "#475569";
  for (let i = 0; i <= 4; i++) {
    const y = margin.top + (h / 4) * i;
    ctx.strokeStyle = "#e2e8f0";
    ctx.beginPath();
    ctx.moveTo(margin.left, y);
    ctx.lineTo(margin.left + w, y);
    ctx.stroke();
    const value = max - ((max - min) / 4) * i;
    ctx.fillText(value.toLocaleString("pt-BR", { maximumFractionDigits: 0 }), 8, y + 4);
  }
  series.forEach(([label, key, color], sidx) => {
    const points = rows.map((row, idx) => ({
      x: margin.left + (w / Math.max(rows.length - 1, 1)) * idx,
      y: row[key] > 0 ? scaleY(row[key]) : null,
    })).filter((p) => p.y !== null);
    if (points.length < 2) return;
    ctx.strokeStyle = color;
    ctx.lineWidth = key === "patrimonio" ? 3 : 2;
    ctx.setLineDash(key === "patrimonio" ? [] : [5, 4]);
    ctx.beginPath();
    points.forEach((p, idx) => idx ? ctx.lineTo(p.x, p.y) : ctx.moveTo(p.x, p.y));
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = color;
    ctx.fillRect(canvas.width - 180, 16 + sidx * 18, 20, 3);
    ctx.fillStyle = "#0f172a";
    ctx.fillText(label, canvas.width - 154, 22 + sidx * 18);
  });
  rows.forEach((row, idx) => {
    if (idx === 0 || idx === rows.length - 1 || idx % Math.max(Math.floor(rows.length / 8), 1) === 0) {
      const x = margin.left + (w / Math.max(rows.length - 1, 1)) * idx;
      ctx.fillStyle = "#475569";
      ctx.fillText(row.periodo, x - 18, canvas.height - 18);
    }
  });
}

window.addEventListener("load", drawWealthChart);
window.addEventListener("resize", drawWealthChart);
