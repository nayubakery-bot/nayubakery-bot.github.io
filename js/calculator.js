// Exact replica of the Excel Custom Cake Pricing Calculator (₹ INR).
// Formulas match the Calculator + Settings + Recipes + Raw Materials workbook.

export function costPerUnit(item) {
  const size = Number(item.packSize);
  const price = Number(item.packPrice);
  if (!isFinite(size) || size <= 0 || !isFinite(price)) return 0;
  return price / size;
}

export function materialMap(rawMaterials) {
  const map = new Map();
  for (const m of rawMaterials) map.set(m.name, m);
  return map;
}

export function recipeCostPerKg(recipeName, recipes, rawMaterials) {
  const mm = materialMap(rawMaterials);
  let total = 0;
  for (const line of recipes) {
    if (line.name !== recipeName) continue;
    if (!line.ingredient || !isFinite(Number(line.qtyPerKg))) continue;
    const mat = mm.get(line.ingredient);
    const rate = mat ? costPerUnit(mat) : 0;
    total += Number(line.qtyPerKg) * rate;
  }
  return total;
}

export function allRecipeCosts(recipes, rawMaterials) {
  const names = [...new Set(recipes.map((r) => r.name))];
  const out = {};
  for (const n of names) out[n] = recipeCostPerKg(n, recipes, rawMaterials);
  return out;
}

/** The add-on behind the "Candles" / "Carry bag" tick boxes (works after a rename too). */
export function findAddon(addons, which) {
  const list = addons || [];
  if (which === "candles") return list.find((a) => a.name === "Candles & knife set") || list.find((a) => /candle/i.test(a.name));
  if (which === "carryBag") return list.find((a) => a.name === "Carry bag") || list.find((a) => /carry\s*bag|bag/i.test(a.name));
  return null;
}

export function lookup(list, key, keyField, valueField, fallback = 0) {
  const row = list.find((x) => x[keyField] === key);
  if (!row) return fallback;
  return row[valueField];
}

function n(v) {
  const x = Number(v);
  return isFinite(x) ? x : 0;
}

/** Excel CEILING(number, significance) for positive numbers. */
export function excelCeiling(value, significance) {
  if (!significance || significance <= 0) return value;
  return Math.ceil(value / significance - 1e-12) * significance;
}

export function fixedOverheads(settings) {
  const o = settings.overheads || {};
  return n(o.gas) + n(o.electricity) + n(o.water) + n(o.rent) + n(o.equipment);
}

export function toppingUsed(name) {
  return !!(name && name !== "None");
}

export function suggestedHours(input, settings) {
  const kg = n(input.weightKg);
  const layers = n(input.layers);
  const tops =
    (toppingUsed(input.topping1) ? 1 : 0) +
    (toppingUsed(input.topping2) ? 1 : 0) +
    (toppingUsed(input.topping3) ? 1 : 0);
  return (
    n(settings.baseHours) +
    n(settings.hoursPerKg) * kg +
    n(settings.hoursPerLayer) * layers +
    n(settings.hoursPerTopping) * tops
  );
}

export function sizeInfo(sizeName, settings) {
  const s = (settings.sizes || []).find((x) => x.name === sizeName);
  return s || { name: sizeName, kg: 0, servings: "" };
}

export function toppingQty(toppingName, weightKg, settings) {
  const t = (settings.toppings || []).find((x) => x.name === toppingName);
  if (!t) return { qty: 0, unit: "", rate: 0, material: "" };
  return {
    qty: n(t.fixedQty) + n(t.qtyPerKg) * n(weightKg),
    unit: t.unit || "",
    material: t.material || "",
    fixedQty: n(t.fixedQty),
    qtyPerKg: n(t.qtyPerKg),
  };
}

/**
 * @param {object} input calculator inputs
 * @param {object} data { rawMaterials, recipes, settings }
 */
export function calculateQuote(input, data) {
  const { rawMaterials, recipes, settings } = data;
  const costs = allRecipeCosts(recipes, rawMaterials);
  const mm = materialMap(rawMaterials);
  const rateOf = (name) => {
    const m = mm.get(name);
    return m ? costPerUnit(m) : 0;
  };
  const unitOf = (name) => {
    const m = mm.get(name);
    return m ? m.unit : "";
  };

  const size = sizeInfo(input.size, settings);
  const weightKg = n(input.weightKg != null ? input.weightKg : size.kg);
  const servings = input.servings != null ? input.servings : size.servings;
  const layers = n(input.layers);
  const wastage = n(settings.wastage);

  const fillingQty = (weightKg * n(settings.fillingGPerKg) * Math.max(layers - 1, 0)) / 1000;
  const crumbQty = (weightKg * n(settings.crumbCoatGPerKgPerLayer) * layers) / 1000;
  const frostingQty = (weightKg * n(settings.finalFrostingGPerKgPerLayer) * layers) / 1000;
  const spongeQty = Math.max(weightKg - fillingQty - crumbQty - frostingQty, 0);

  const spongeRate = n(costs[input.sponge]);
  const fillingRate = n(costs[input.filling]);
  const frostingRate = n(costs[input.frosting]);

  const spongeCost = spongeQty * (1 + wastage) * spongeRate;
  const fillingCost = fillingQty * (1 + wastage) * fillingRate;
  const crumbCost = crumbQty * (1 + wastage) * frostingRate;
  const frostingCost = frostingQty * (1 + wastage) * frostingRate;

  const toppingLines = [input.topping1, input.topping2, input.topping3].map((name) => {
    const info = toppingQty(name, weightKg, settings);
    const material = info.material;
    const rate = material ? rateOf(material) : 0;
    const unit = material ? unitOf(material) : info.unit;
    const cost = info.qty * rate;
    return { name: name || "", material, qty: info.qty, unit, rate, cost };
  });

  const boxRate = input.box ? rateOf(input.box) : 0;
  const boardRate = input.board ? rateOf(input.board) : 0;
  const boxQty = input.box ? 1 : 0;
  const boardQty = input.board ? 1 : 0;
  const boxCost = boxQty * boxRate;
  const boardCost = boardQty * boardRate;

  // Add-ons: the two classic tick boxes plus any extra add-ons she created.
  const addons = settings.addons || [];
  const candleAddon = findAddon(addons, "candles");
  const bagAddon = findAddon(addons, "carryBag");
  const addonLine = (a) => {
    const qty = a ? n(a.qty) : 0;
    const rate = a ? rateOf(a.material) : 0;
    return { name: a ? a.name : "", material: a ? a.material : "", qty, unit: a ? unitOf(a.material) || "nos" : "nos", rate, cost: qty * rate };
  };
  const candlesLine = input.candles && candleAddon ? addonLine(candleAddon) : { qty: 0, rate: candleAddon ? rateOf(candleAddon.material) : 0, cost: 0 };
  const bagLine = input.carryBag && bagAddon ? addonLine(bagAddon) : { qty: 0, rate: bagAddon ? rateOf(bagAddon.material) : 0, cost: 0 };
  const candlesQty = candlesLine.qty, candlesRate = candlesLine.rate, candlesCost = candlesLine.cost;
  const bagQty = bagLine.qty, bagRate = bagLine.rate, bagCost = bagLine.cost;
  const extraAddonLines = (input.extraAddons || [])
    .map((nm) => addons.find((a) => a.name === nm))
    .filter((a) => a && a !== candleAddon && a !== bagAddon)
    .map(addonLine);
  const extraAddonsCost = extraAddonLines.reduce((s, l) => s + l.cost, 0);

  const ingredientCost = spongeCost + fillingCost + crumbCost + frostingCost;
  const toppingsCost = toppingLines.reduce((s, t) => s + t.cost, 0);
  const packagingCost = boxCost + boardCost + candlesCost + bagCost + extraAddonsCost;
  const rawMaterialCost = ingredientCost + toppingsCost + packagingCost;

  const hoursSuggested = suggestedHours({ ...input, weightKg }, settings);
  const hoursUsed =
    input.hoursOverride === "" || input.hoursOverride == null || input.hoursOverride === undefined
      ? hoursSuggested
      : n(input.hoursOverride);

  const overheads = fixedOverheads(settings);
  const labour = hoursUsed * n(settings.labourRate);
  const totalCost = rawMaterialCost + overheads + labour;
  const markup = n(settings.markup);
  const calcProfit = totalCost * markup;
  const calculatedBeforeDelivery = totalCost + calcProfit;

  // Menu price: when she has set a selling price for this size, the cake is priced
  // from her price list (size price + each extra at its own selling price, or at
  // cost + markup when it has none). Otherwise the workbook's cost-plus formula.
  const priceList = data.priceList || [];
  const sell = (kind, item) => {
    if (!item || item === "None") return null;
    const p = priceList.find((x) => x.kind === kind && x.item === item);
    return p && p.price !== "" && p.price != null && isFinite(Number(p.price)) && Number(p.price) > 0 ? Number(p.price) : null;
  };
  const sizePrice = sell("Size", input.size);
  let menu = null;
  if (sizePrice != null) {
    const withMarkup = (cost) => cost * (1 + markup);
    const parts = [{ label: "Cake " + (input.size || ""), price: sizePrice }];
    toppingLines.forEach((t) => { if (toppingUsed(t.name)) parts.push({ label: t.name, price: sell("Topping", t.name) ?? withMarkup(t.cost) }); });
    if (input.box) parts.push({ label: input.box, price: sell("Box", input.box) ?? withMarkup(boxCost) });
    if (input.board) parts.push({ label: input.board, price: sell("Board", input.board) ?? withMarkup(boardCost) });
    if (input.candles && candleAddon) parts.push({ label: candleAddon.name, price: sell("Add-on", candleAddon.name) ?? withMarkup(candlesCost) });
    if (input.carryBag && bagAddon) parts.push({ label: bagAddon.name, price: sell("Add-on", bagAddon.name) ?? withMarkup(bagCost) });
    extraAddonLines.forEach((a) => parts.push({ label: a.name, price: sell("Add-on", a.name) ?? withMarkup(a.cost) }));
    menu = { parts, total: parts.reduce((s, p) => s + p.price, 0) };
  }
  const baseBeforeDelivery = menu ? menu.total : calculatedBeforeDelivery;

  // Pricing rules (rush / express charges, festival discounts)
  const ruleDefs = data.rules || [];
  const ruleLines = (input.rules || [])
    .map((nm) => ruleDefs.find((r) => r.name === nm))
    .filter(Boolean)
    .map((r) => {
      const amt = r.type === "₹" ? n(r.amount) : (baseBeforeDelivery * n(r.amount)) / 100;
      return { name: r.name, effect: r.effect, type: r.type, amount: n(r.amount), value: r.effect === "Discount" ? -amt : amt };
    });
  const rulesTotal = ruleLines.reduce((s, r) => s + r.value, 0);
  const priceBeforeDelivery = Math.max(baseBeforeDelivery + rulesTotal, 0);
  let profit = priceBeforeDelivery - totalCost;
  if (!menu && !ruleLines.length) profit = calcProfit; // identical to the workbook

  let delivery = 0;
  let deliveryZone = null;
  if (input.delivery === "Home delivery") {
    const zone = input.zone ? (data.zones || []).find((z) => z.name === input.zone) : null;
    if (zone) {
      delivery = n(zone.charge);
      deliveryZone = zone.name;
    } else {
      delivery = Math.max(
        n(settings.minDelivery),
        n(input.distanceKm) * n(settings.deliveryRatePerKm) * n(settings.tripMultiplier)
      );
    }
  }
  let gst = (priceBeforeDelivery + delivery) * n(settings.gst);
  let beforeRounding = priceBeforeDelivery + delivery + gst;
  const calculatedQuote = excelCeiling(beforeRounding, n(settings.rounding));
  let finalQuote = calculatedQuote;
  const override = input.priceOverride === "" || input.priceOverride == null ? null : Number(input.priceOverride);
  if (override != null && isFinite(override) && override > 0) {
    // her own price for this order: GST is inside it, delivery is still passed on
    finalQuote = override;
    gst = n(settings.gst) ? (override * n(settings.gst)) / (1 + n(settings.gst)) : 0;
    beforeRounding = override;
    profit = override - gst - delivery - totalCost;
  }
  const pricePerKgExclDelivery = weightKg > 0 ? priceBeforeDelivery / weightKg : 0;
  const weightCheck = spongeQty + fillingQty + crumbQty + frostingQty;
  const weightOk = Math.abs(weightCheck - weightKg) < 0.001;

  return {
    weightKg,
    servings,
    layers,
    hoursSuggested,
    hoursUsed,
    lines: {
      sponge: { name: input.sponge, qty: spongeQty, unit: "kg", rate: spongeRate, cost: spongeCost },
      filling: { name: input.filling, qty: fillingQty, unit: "kg", rate: fillingRate, cost: fillingCost },
      crumbCoat: { name: input.frosting, qty: crumbQty, unit: "kg", rate: frostingRate, cost: crumbCost },
      frosting: { name: input.frosting, qty: frostingQty, unit: "kg", rate: frostingRate, cost: frostingCost },
      toppings: toppingLines,
      box: { name: input.box, qty: boxQty, unit: "nos", rate: boxRate, cost: boxCost },
      board: { name: input.board, qty: boardQty, unit: "nos", rate: boardRate, cost: boardCost },
      candles: {
        name: input.candles ? "Included" : "Not added",
        qty: candlesQty,
        unit: "nos",
        rate: candlesRate,
        cost: candlesCost,
      },
      carryBag: {
        name: input.carryBag ? "Included" : "Not added",
        qty: bagQty,
        unit: "nos",
        rate: bagRate,
        cost: bagCost,
      },
      extraAddons: extraAddonLines,
    },
    ingredientCost,
    toppingsCost,
    packagingCost,
    rawMaterialCost,
    overheads,
    labour,
    totalCost,
    profit,
    priceBeforeDelivery,
    delivery,
    gst,
    beforeRounding,
    finalQuote,
    calculatedQuote,
    calculatedBeforeDelivery,
    priceOverride: override != null && isFinite(override) && override > 0 ? override : null,
    menu,
    rules: ruleLines,
    rulesTotal,
    deliveryZone,
    pricePerKgExclDelivery,
    weightCheck,
    weightOk,
  };
}

export function scaleRecipe(recipeName, targetKg, recipes, rawMaterials) {
  const mm = materialMap(rawMaterials);
  const kg = n(targetKg);
  const lines = recipes
    .filter((r) => r.name === recipeName)
    .map((line) => {
      const mat = mm.get(line.ingredient);
      const unit = line.unit || (mat ? mat.unit : "");
      const qty = n(line.qtyPerKg) * kg;
      const rate = mat ? costPerUnit(mat) : 0;
      const cost = qty * rate;
      const packSize = mat ? n(mat.packSize) : 0;
      const packsToBuy = packSize > 0 ? Math.ceil(qty / packSize - 1e-12) : 0;
      const packPrice = mat ? n(mat.packPrice) : 0;
      return {
        ingredient: line.ingredient,
        unit,
        qtyPerKg: n(line.qtyPerKg),
        qty,
        rate,
        cost,
        packSize,
        packsToBuy,
        packPrice,
        packSpend: packsToBuy * packPrice,
        estimated: !!(mat && mat.estimated),
        missing: !mat,
      };
    });
  const totalCost = lines.reduce((s, l) => s + l.cost, 0);
  const packSpend = lines.reduce((s, l) => s + l.packSpend, 0);
  return { recipeName, targetKg: kg, lines, totalCost, packSpend, costPerKg: kg ? totalCost / kg : 0 };
}

export function rupees(nVal, digits = 2) {
  if (!isFinite(nVal)) return "₹0";
  return "₹" + Number(nVal).toLocaleString("en-IN", { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

export function rupeesWhole(nVal) {
  if (!isFinite(nVal)) return "₹0";
  return "₹" + Math.round(nVal).toLocaleString("en-IN");
}
