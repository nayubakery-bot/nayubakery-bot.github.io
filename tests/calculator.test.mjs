import { readFileSync } from "fs";
import { calculateQuote, recipeCostPerKg, allRecipeCosts, scaleRecipe } from "../js/calculator.js";

const seed = JSON.parse(readFileSync(new URL("../js/seed-data.json", import.meta.url)));
const data = { rawMaterials: seed.rawMaterials.filter(m => !m.estimated), recipes: seed.recipes, settings: seed.settings };

const expectedCosts = {
  "Vanilla Sponge": 100.035,
  "Chocolate Sponge": 133.7399,
  "Red Velvet Sponge": 137.6491,
  "Coconut Sponge": 117.578,
  "Whipped Cream": 315.1,
  "Chocolate Ganache": 422,
  "Strawberry Cream": 360.5,
  "Pineapple Cream": 354.5,
  "Cream Cheese Filling": 555.75,
  "Swiss Meringue Buttercream": 357.75,
  "Mascarpone Cream Filling": 815.5,
  "Whipped Cream Frosting": 353.1,
  "Chocolate Ganache Frosting": 422,
  "Vanilla Buttercream": 292.96,
  "Cream Cheese Frosting": 488.25,
  "Swiss Meringue Buttercream Frosting": 357.75,
  "Mascarpone Cream Frosting": 815.5,
};

let failed = 0;
function nearly(a, b, eps = 1e-6, label = "") {
  const ok = Math.abs(a - b) <= eps || Math.abs(a - b) / (Math.abs(b) || 1) <= 1e-9;
  if (!ok) {
    failed++;
    console.error("FAIL", label, "got", a, "expected", b, "diff", a - b);
  } else {
    console.log("OK", label, a);
  }
}

for (const [name, exp] of Object.entries(expectedCosts)) {
  nearly(recipeCostPerKg(name, data.recipes, data.rawMaterials), exp, 1e-6, "cost/kg " + name);
}

const s = seed.sampleOrder;
const q = calculateQuote(s, data);
const e = s.expected;
nearly(q.weightKg, e.weightKg, 1e-12, "weight");
nearly(q.hoursSuggested, e.suggestedHours, 1e-12, "suggested hours");
nearly(q.hoursUsed, e.hoursUsed, 1e-12, "hours used");
nearly(q.lines.sponge.qty, e.spongeQty, 1e-12, "sponge qty");
nearly(q.lines.filling.qty, e.fillingQty, 1e-12, "filling qty");
nearly(q.lines.crumbCoat.qty, e.crumbQty, 1e-12, "crumb qty");
nearly(q.lines.frosting.qty, e.frostingQty, 1e-12, "frosting qty");
nearly(q.lines.sponge.cost, e.spongeCost, 1e-6, "sponge cost");
nearly(q.lines.filling.cost, e.fillingCost, 1e-6, "filling cost");
nearly(q.lines.crumbCoat.cost, e.crumbCost, 1e-6, "crumb cost");
nearly(q.lines.frosting.cost, e.frostingCost, 1e-6, "frosting cost");
nearly(q.ingredientCost, e.ingredientCost, 1e-6, "ingredient cost");
nearly(q.toppingsCost, e.toppingsCost, 1e-6, "toppings");
nearly(q.packagingCost, e.packagingCost, 1e-6, "packaging");
nearly(q.rawMaterialCost, e.rawMaterialCost, 1e-6, "raw material");
nearly(q.overheads, e.overheads, 1e-12, "overheads");
nearly(q.labour, e.labour, 1e-12, "labour");
nearly(q.totalCost, e.totalCost, 1e-6, "total cost");
nearly(q.profit, e.profit, 1e-6, "profit");
nearly(q.priceBeforeDelivery, e.priceBeforeDelivery, 1e-6, "before delivery");
nearly(q.delivery, e.delivery, 1e-12, "delivery");
nearly(q.gst, e.gst, 1e-12, "gst");
nearly(q.beforeRounding, e.beforeRounding, 1e-6, "before rounding");
nearly(q.finalQuote, e.finalQuote, 1e-9, "FINAL QUOTE");
nearly(q.pricePerKgExclDelivery, e.pricePerKgExclDelivery, 1e-6, "per kg");
if (q.servings !== e.servings) { failed++; console.error("FAIL servings", q.servings, e.servings); }
else console.log("OK servings", q.servings);

const scaled = scaleRecipe("Chocolate Sponge", 0.37, data.recipes, data.rawMaterials);
nearly(scaled.costPerKg, expectedCosts["Chocolate Sponge"], 1e-6, "scaler cost/kg");

if (failed) {
  console.error("\n" + failed + " failures");
  process.exit(1);
}
console.log("\nAll checks passed.");
