/**
 * 基于偏好权重的优化算法
 * @param {number} a - 沪深300夏普比率
 * @param {number} b - 中证500夏普比率
 * @param {number} c - 中证1000夏普比率
 * @param {number} d - 中证2000夏普比率
 * @param {Object} options - 配置选项
 * @returns {Object} 优化结果
 */
function optimizeWithPreference(a, b, c, d, options = {}) {
  const {
    targetW = 0.4, // w的目标权重
    targetXYZ = 0.2, // x, y, z的目标权重
    alpha = 0.3, // 平衡参数：0-1，越大越注重夏普比率，越小越注重权重偏好
    iterations = 10000, // 迭代次数
    tolerance = 1e-8, // 容差
  } = options;

  // 夏普比率数组
  const sharpeRatios = [a, b, c, d];

  // 目标权重数组
  const targetWeights = [targetW, targetXYZ, targetXYZ, targetXYZ];

  // 归一化目标权重确保和为1
  const targetSum = targetWeights.reduce((sum, w) => sum + w, 0);
  for (let i = 0; i < targetWeights.length; i++) {
    targetWeights[i] /= targetSum;
  }

  // 初始化权重
  let weights = [...targetWeights];
  let bestWeights = [...weights];
  let bestObjective = calculateObjective(
    weights,
    sharpeRatios,
    targetWeights,
    alpha
  );

  // 模拟退火参数
  let temperature = 1.0;
  const coolingRate = 0.995;

  // 迭代优化
  for (let iter = 0; iter < iterations; iter++) {
    // 生成新权重
    const newWeights = generateNeighbor(weights);

    // 确保权重有效
    normalizeWeights(newWeights);

    // 计算目标函数值
    const newObjective = calculateObjective(
      newWeights,
      sharpeRatios,
      targetWeights,
      alpha
    );

    // 接受新解的条件
    const delta = newObjective - bestObjective;

    if (delta > 0 || Math.random() < Math.exp(delta / temperature)) {
      weights = newWeights;

      if (newObjective > bestObjective) {
        bestWeights = [...newWeights];
        bestObjective = newObjective;
      }
    }

    // 降温
    temperature *= coolingRate;

    // 检查收敛
    if (temperature < tolerance) break;
  }

  // 最终调整确保权重在有效范围内
  adjustWeightsToBounds(bestWeights);

  // 计算最终夏普比率
  const finalSharpe =
    a * bestWeights[0] +
    b * bestWeights[1] +
    c * bestWeights[2] +
    d * bestWeights[3];

  return {
    sharpeRatio: finalSharpe,
    weights: {
      w: bestWeights[0],
      x: bestWeights[1],
      y: bestWeights[2],
      z: bestWeights[3],
    },
    preferences: {
      targetW: targetW,
      targetXYZ: targetXYZ,
      achievedW: bestWeights[0],
      deviation:
        Math.abs(bestWeights[0] - targetW) +
        Math.abs(bestWeights[1] - targetXYZ) +
        Math.abs(bestWeights[2] - targetXYZ) +
        Math.abs(bestWeights[3] - targetXYZ),
    },
  };
}

/**
 * 计算目标函数值
 * 平衡夏普比率和权重偏好
 */
function calculateObjective(weights, sharpeRatios, targetWeights, alpha) {
  // 计算夏普比率部分
  let sharpePart = 0;
  for (let i = 0; i < weights.length; i++) {
    sharpePart += weights[i] * sharpeRatios[i];
  }

  // 计算权重偏差部分（使用负偏差，因为我们要最大化目标函数）
  let deviationPart = 0;
  for (let i = 0; i < weights.length; i++) {
    deviationPart -= Math.pow(weights[i] - targetWeights[i], 2);
  }

  // 加权组合：alpha控制对夏普比率的重视程度
  return alpha * sharpePart + (1 - alpha) * deviationPart;
}

/**
 * 生成邻域解（微调当前权重）
 */
function generateNeighbor(weights) {
  const newWeights = [...weights];
  const minAdjustment = -0.05; // 最大调整-5%
  const maxAdjustment = 0.05; // 最大调整+5%

  // 随机选择两个权重进行调整（一个增加，一个减少）
  const idx1 = Math.floor(Math.random() * weights.length);
  const idx2 = Math.floor(Math.random() * weights.length);

  if (idx1 !== idx2) {
    // 随机调整量
    const adjustment =
      minAdjustment + Math.random() * (maxAdjustment - minAdjustment);

    // 确保调整后仍在有效范围内
    const maxReduction = Math.max(0.0001 - newWeights[idx1], -0.05);
    const maxIncrease = Math.min(0.9999 - newWeights[idx2], 0.05);
    const actualAdjustment = Math.max(
      maxReduction,
      Math.min(adjustment, maxIncrease)
    );

    // 应用调整
    newWeights[idx1] += actualAdjustment;
    newWeights[idx2] -= actualAdjustment;
  }

  return newWeights;
}

/**
 * 归一化权重确保和为1
 */
function normalizeWeights(weights) {
  const sum = weights.reduce((s, w) => s + w, 0);
  if (Math.abs(sum - 1) > 1e-10) {
    const factor = 1 / sum;
    for (let i = 0; i < weights.length; i++) {
      weights[i] *= factor;
    }
  }
}

/**
 * 调整权重到有效范围内
 */
function adjustWeightsToBounds(weights) {
  const minWeight = 0.0001;
  const maxWeight = 0.9999;

  // 首先确保所有权重在有效范围内
  for (let i = 0; i < weights.length; i++) {
    if (weights[i] < minWeight) weights[i] = minWeight;
    if (weights[i] > maxWeight) weights[i] = maxWeight;
  }

  // 调整总和为1
  let sum = weights.reduce((s, w) => s + w, 0);
  if (Math.abs(sum - 1) > 1e-10) {
    // 计算需要调整的量
    let adjustment = (1 - sum) / weights.length;
    let remainingAdjustment = adjustment;

    // 应用调整，确保不超出边界
    for (let i = 0; i < weights.length; i++) {
      const newWeight = weights[i] + remainingAdjustment;
      if (newWeight >= minWeight && newWeight <= maxWeight) {
        weights[i] = newWeight;
        remainingAdjustment = 0;
        break;
      }
    }

    // 如果还有剩余调整量，继续调整
    if (Math.abs(remainingAdjustment) > 1e-10) {
      adjustWeightsToBounds(weights);
    }
  }
}

/**
 * 格式化为百分比显示
 */
function formatPercent(decimal, decimals = 4) {
  return (decimal * 100).toFixed(decimals) + "%";
}

/**
 * 主函数：提供多种优化策略
 */
function optimizePortfolio(a, b, c, d, strategy = "balanced") {
  const strategies = {
    // 策略1：平衡型（默认）
    balanced: {
      targetW: 0.4,
      targetXYZ: 0.2,
      alpha: 0.5, // 平衡夏普比率和权重偏好
    },

    // 策略2：更注重夏普比率
    sharpeFocused: {
      targetW: 0.4,
      targetXYZ: 0.2,
      alpha: 0.7, // 更注重夏普比率
    },

    // 策略3：更注重权重偏好
    weightFocused: {
      targetW: 0.4,
      targetXYZ: 0.2,
      alpha: 0.3, // 更注重权重偏好
    },

    // 策略4：灵活调整（根据夏普比率动态调整）
    adaptive: {
      targetW: 0.4,
      targetXYZ: 0.2,
      alpha: calculateAdaptiveAlpha(a, b, c, d),
    },
  };

  const config = strategies[strategy] || strategies.balanced;

  return optimizeWithPreference(a, b, c, d, config);
}

/**
 * 根据夏普比率动态计算alpha
 */
function calculateAdaptiveAlpha(a, b, c, d) {
  // 计算夏普比率的方差
  const sharpeRatios = [a, b, c, d];
  const mean = sharpeRatios.reduce((sum, r) => sum + r, 0) / 4;
  const variance =
    sharpeRatios.reduce((sum, r) => sum + Math.pow(r - mean, 2), 0) / 4;

  // 方差越大，越应该注重夏普比率（alpha越大）
  // 将方差映射到0.3-0.7之间
  const minAlpha = 0.3;
  const maxAlpha = 0.7;
  const normalizedVariance = Math.min(1, Math.max(0, variance / 2)); // 假设最大方差为2

  return minAlpha + normalizedVariance * (maxAlpha - minAlpha);
}

/**
 * 示例展示
 */
function runExamples() {
  console.log("=".repeat(60));
  console.log("投资组合优化示例");
  console.log("=".repeat(60));

  // 示例1：正常情况
  console.log("\n示例1: 正常夏普比率");
  const result1 = optimizePortfolio(0.5, 0.3, 0.4, 0.2, "balanced");
  console.log(`夏普比率: ${result1.sharpeRatio.toFixed(6)}`);
  console.log(`权重分配:`);
  console.log(
    `  沪深300(w): ${formatPercent(result1.weights.w)} (目标: ${formatPercent(
      0.4
    )})`
  );
  console.log(
    `  中证500(x): ${formatPercent(result1.weights.x)} (目标: ${formatPercent(
      0.2
    )})`
  );
  console.log(
    `  中证1000(y): ${formatPercent(result1.weights.y)} (目标: ${formatPercent(
      0.2
    )})`
  );
  console.log(
    `  中证2000(z): ${formatPercent(result1.weights.z)} (目标: ${formatPercent(
      0.2
    )})`
  );
  console.log(`总偏差: ${result1.preferences.deviation.toFixed(6)}`);

  console.log("\n" + "-".repeat(40));
  console.log("不同策略比较:");

  for (const strategy of [
    "balanced",
    "sharpeFocused",
    "weightFocused",
    "adaptive",
  ]) {
    const res = optimizePortfolio(0.5, 0.3, 0.4, 0.2, strategy);
    console.log(
      `${strategy.padEnd(15)}: 夏普=${res.sharpeRatio.toFixed(
        4
      )}, w=${formatPercent(res.weights.w, 2)}`
    );
  }

  console.log("\n" + "=".repeat(60));

  // 示例2：有负夏普比率
  console.log("\n示例2: 包含负夏普比率");
  const result2 = optimizePortfolio(0.2, -0.1, 0.3, 0.1, "balanced");
  console.log(`夏普比率: ${result2.sharpeRatio.toFixed(6)}`);
  console.log(`权重分配:`);
  console.log(`  沪深300: ${formatPercent(result2.weights.w)}`);
  console.log(`  中证500: ${formatPercent(result2.weights.x)}`);
  console.log(`  中证1000: ${formatPercent(result2.weights.y)}`);
  console.log(`  中证2000: ${formatPercent(result2.weights.z)}`);

  console.log("\n" + "=".repeat(60));

  // 示例3：所有夏普比率相等
  console.log("\n示例3: 所有夏普比率相等 (0.25)");
  const result3 = optimizePortfolio(0.25, 0.25, 0.25, 0.25, "balanced");
  console.log(`权重分配 (应接近目标权重):`);
  console.log(`  沪深300: ${formatPercent(result3.weights.w)}`);
  console.log(`  中证500: ${formatPercent(result3.weights.x)}`);
  console.log(`  中证1000: ${formatPercent(result3.weights.y)}`);
  console.log(`  中证2000: ${formatPercent(result3.weights.z)}`);
}

/**
 * 计算纯最大夏普比率（用于比较）
 */
function calculateMaxSharpeOnly(a, b, c, d) {
  const sharpeRatios = [a, b, c, d];
  const maxIndex = sharpeRatios.indexOf(Math.max(...sharpeRatios));
  const weights = [0.0001, 0.0001, 0.0001, 0.0001];
  weights[maxIndex] = 1 - 0.0001 * 3;

  return {
    sharpe: a * weights[0] + b * weights[1] + c * weights[2] + d * weights[3],
    weights: weights,
  };
}

/**
 * 用户友好的主函数
 */
function main(a, b, c, d, userOptions = {}) {
  // 合并用户选项
  const defaultOptions = {
    strategy: "balanced",
    targetW: 0.4,
    targetXYZ: 0.2,
  };

  const options = { ...defaultOptions, ...userOptions };

  // 计算纯最大夏普比率（作为参考）
  const maxSharpeResult = calculateMaxSharpeOnly(a, b, c, d);

  // 计算带偏好的优化结果
  const optimizedResult = optimizePortfolio(a, b, c, d, options.strategy);

  // 计算折损比例
  const sacrificeRatio =
    1 - optimizedResult.sharpeRatio / maxSharpeResult.sharpe;

  return {
    optimized: optimizedResult,
    comparison: {
      maxPossibleSharpe: maxSharpeResult.sharpe,
      sacrificeForBalance: sacrificeRatio,
      message: `为了权重平衡，牺牲了${(sacrificeRatio * 100).toFixed(
        2
      )}%的夏普比率`,
    },
    recommendations: generateRecommendations(optimizedResult, maxSharpeResult),
  };
}

/**
 * 生成投资建议
 */
function generateRecommendations(optimizedResult, maxSharpeResult) {
  const wDiff = optimizedResult.weights.w - 0.4;
  const recommendations = [];

  if (wDiff > 0.05) {
    recommendations.push("建议减少沪深300配置，增加其他指数配置以分散风险");
  } else if (wDiff < -0.05) {
    recommendations.push("建议增加沪深300配置，以获得更稳定的收益");
  }

  if (optimizedResult.sharpeRatio < 0) {
    recommendations.push("警告：当前配置的预期夏普比率为负，请谨慎投资");
  }

  return recommendations;
}

// 导出函数
module.exports = {
  optimizePortfolio,
  optimizeWithPreference,
  main,
  runExamples,
  formatPercent,
};

// 如果直接运行，执行示例
if (require.main === module) {
  runExamples();
}
