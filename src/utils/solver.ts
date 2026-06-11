export interface CutItem {
  id: string;
  constructionId: string;
  constructionName: string;
  profile: string;
  length: number; // in mm
  qty: number;
  position: string;
}

export interface NestedBar {
  id: string;
  stockLength: number;
  trim: number;
  kerf: number;
  cuts: {
    id: string;
    length: number;
    position: string;
    constructionName: string;
    cutIndex: number; // index of cut in the bar
    startOffset: number; // mm from start of bar
    endOffset: number; // mm from start of bar
  }[];
  usedLength: number; // total length of cuts + kerfs (not including trim)
  totalConsumed: number; // usedLength + 2 * trim
  waste: number; // remaining length
  wastePercent: number;
}

export interface SolveResult {
  profile: string;
  stockLength: number;
  kerf: number;
  trim: number;
  bars: NestedBar[];
  totalBars: number;
  totalCuts: number;
  totalWaste: number;
  wastePercent: number;
  yieldPercent: number;
}

export interface SolverSettings {
  stockLength: number;
  kerf: number;
  trim: number;
  customProfileSettings: Record<string, { stockLength?: number; kerf?: number; trim?: number }>;
}

/**
 * Solve 1D nesting for a list of demands.
 * Uses a Best-Fit Decreasing (BFD) algorithm.
 * To optimize further, it runs a randomized shuffle on equivalent/similar parts
 * to find the absolute minimum stock bars.
 */
export function solveNesting(
  profile: string,
  cuts: CutItem[],
  settings: SolverSettings
): SolveResult {
  // Get active settings for this profile
  const profileSettings = settings.customProfileSettings[profile] || {};
  const stockLength = profileSettings.stockLength ?? settings.stockLength;
  const kerf = profileSettings.kerf ?? settings.kerf;
  const trim = profileSettings.trim ?? settings.trim;

  // Flatten the demands: if qty is 5, create 5 separate cut tasks
  const flatCuts: { id: string; length: number; position: string; constructionName: string }[] = [];
  for (const cut of cuts) {
    for (let i = 0; i < cut.qty; i++) {
      flatCuts.push({
        id: `${cut.id}-${i}`,
        length: cut.length,
        position: cut.position,
        constructionName: cut.constructionName,
      });
    }
  }

  if (flatCuts.length === 0) {
    return {
      profile,
      stockLength,
      kerf,
      trim,
      bars: [],
      totalBars: 0,
      totalCuts: 0,
      totalWaste: 0,
      wastePercent: 0,
      yieldPercent: 100,
    };
  }

  // Check if any cut exceeds the maximum possible cuts (stockLength - 2 * trim)
  const usableLength = stockLength - 2 * trim;
  const oversizedCuts = flatCuts.filter((c) => c.length > usableLength);
  if (oversizedCuts.length > 0) {
    // We will still handle them but put them in "impossible" bars
    // To avoid breaking the layout, we'll assign them their own bar which shows as error/overflow.
  }

  // Let's run a multi-start search:
  // 1. Sort descending (standard BFD) - usually optimal.
  // 2. Run randomized shuffles on items of similar sizes to see if we can get a better packing.
  // We'll run 50 iterations and choose the one with the fewest bars, and break ties with the lowest total waste variance or highest front-loading.
  let bestBars: NestedBar[] = [];
  let bestBarCount = Infinity;
  let bestWasteScore = Infinity; // Lower is better (front-loads bars, leaving larger scraps)

  // Iteration 0: Standard BFD (sorted strictly descending)
  const sortedCuts = [...flatCuts].sort((a, b) => b.length - a.length);
  
  for (let iter = 0; iter < 100; iter++) {
    let currentCuts = [...sortedCuts];
    if (iter > 0) {
      // Perturb the order slightly by shuffling items that are close in size,
      // or randomly swapping a few elements.
      currentCuts = perturbCuts(sortedCuts);
    }

    const bars = runBFD(currentCuts, stockLength, kerf, trim);
    
    // Evaluate the solution
    const barCount = bars.length;
    // Waste score: we want to consolidate cuts into fewer bars, leaving large remainders.
    // A solution with 1 bar with 90% waste is better than 2 bars with 45% waste each.
    // Sum of squared remaining waste encourages large remainders.
    const wasteScore = -bars.reduce((sum, bar) => sum + bar.waste * bar.waste, 0);

    if (barCount < bestBarCount || (barCount === bestBarCount && wasteScore < bestWasteScore)) {
      bestBars = bars;
      bestBarCount = barCount;
      bestWasteScore = wasteScore;
    }
  }

  // Calculate statistics
  const totalCuts = flatCuts.length;
  const totalBarsUsed = bestBars.length;
  const totalInputLength = totalBarsUsed * stockLength;
  
  // Total length of cuts (actual material in final products)
  const totalCutsLength = flatCuts.reduce((sum, c) => sum + c.length, 0);
  
  // Total waste: stock we bought - cuts we actually use.
  // (This includes kerfs, trims, and leftover pieces)
  const totalWaste = totalInputLength - totalCutsLength;
  const wastePercent = totalInputLength > 0 ? (totalWaste / totalInputLength) * 100 : 0;
  const yieldPercent = 100 - wastePercent;

  return {
    profile,
    stockLength,
    kerf,
    trim,
    bars: bestBars,
    totalBars: totalBarsUsed,
    totalCuts,
    totalWaste,
    wastePercent,
    yieldPercent,
  };
}

/**
 * Perturbs the sorted list of cuts to explore alternative packing solutions.
 */
function perturbCuts(
  cuts: { id: string; length: number; position: string; constructionName: string }[]
) {
  const result = [...cuts];
  // Randomly swap adjacent elements or elements close in size
  for (let i = 0; i < result.length - 1; i++) {
    if (Math.random() < 0.15) {
      // Swap with next element
      const temp = result[i];
      result[i] = result[i + 1];
      result[i + 1] = temp;
    }
  }
  return result;
}

/**
 * Standard Best-Fit Decreasing packing
 */
function runBFD(
  cuts: { id: string; length: number; position: string; constructionName: string }[],
  stockLength: number,
  kerf: number,
  trim: number
): NestedBar[] {
  const usableLength = stockLength - 2 * trim;
  const bars: { cuts: typeof cuts; remainingSpace: number }[] = [];

  for (const cut of cuts) {
    // For oversized cuts, allocate them to a special dedicated bar immediately
    if (cut.length > usableLength) {
      bars.push({
        cuts: [cut],
        remainingSpace: usableLength - cut.length, // Will be negative, indicates overflow
      });
      continue;
    }

    let bestBarIndex = -1;
    let minRemainingAfter = Infinity;

    for (let i = 0; i < bars.length; i++) {
      const bar = bars[i];
      
      // Calculate space required to add this cut to this bar.
      // If the bar already has cuts, we must add kerf.
      const neededSpace = cut.length + (bar.cuts.length > 0 ? kerf : 0);
      
      if (bar.remainingSpace >= neededSpace) {
        const remainingAfter = bar.remainingSpace - neededSpace;
        if (remainingAfter < minRemainingAfter) {
          minRemainingAfter = remainingAfter;
          bestBarIndex = i;
        }
      }
    }

    if (bestBarIndex !== -1) {
      // Place in the best fitting bar
      const bar = bars[bestBarIndex];
      const neededSpace = cut.length + (bar.cuts.length > 0 ? kerf : 0);
      bar.cuts.push(cut);
      bar.remainingSpace -= neededSpace;
    } else {
      // Create a new bar
      bars.push({
        cuts: [cut],
        remainingSpace: usableLength - cut.length,
      });
    }
  }

  // Convert to NestedBar format
  return bars.map((b, index) => {
    const nestedCuts: NestedBar['cuts'] = [];
    let currentOffset = trim;

    b.cuts.forEach((cut, cutIndex) => {
      const startOffset = currentOffset;
      const endOffset = startOffset + cut.length;
      
      nestedCuts.push({
        id: cut.id,
        length: cut.length,
        position: cut.position,
        constructionName: cut.constructionName,
        cutIndex,
        startOffset,
        endOffset,
      });

      currentOffset = endOffset + kerf;
    });

    const usedLength = nestedCuts.length > 0 
      ? nestedCuts[nestedCuts.length - 1].endOffset - trim 
      : 0;

    const totalConsumed = usedLength > 0 ? usedLength + 2 * trim : 0;
    const waste = stockLength - totalConsumed;
    const wastePercent = stockLength > 0 ? (waste / stockLength) * 100 : 0;

    return {
      id: `bar-${index + 1}`,
      stockLength,
      trim,
      kerf,
      cuts: nestedCuts,
      usedLength,
      totalConsumed,
      waste,
      wastePercent,
    };
  });
}
