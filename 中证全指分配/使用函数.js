// // 导入模块
// const portfolioOptimizer = require("./计算沪深300-中证500-中证1000中证2000分配比例.js");

// // 使用示例
// const result = portfolioOptimizer.calculateOptimalPortfolio({
//   a: 2.568, // 沪深300夏普比率
//   b: 1.85, // 中证500夏普比率
//   c: 1.494, // 中证1000夏普比率
//   d: 2.324, // 中证2000夏普比率
// });

// console.log("最大夏普比率:", result.maxSharpe);
// console.log("权重分配:", result.weights);

// 或者直接运行示例
// portfolioOptimizer.example();

const optimizer = require("./计算沪深300-中证500-中证1000中证2000分配比例.js");

// 简单使用
const result = optimizer.main(2.568, 1.85, 1.494, 2.324);

console.log("优化结果:", result.optimized);
console.log("比较信息:", result.comparison);
console.log("投资建议:", result.recommendations);

// 自定义配置
// const customResult = optimizer.optimizeWithPreference(0.5, 0.3, 0.4, 0.2, {
//   targetW: 0.35, // 调整w的目标权重
//   targetXYZ: 0.2167, // 调整x,y,z的目标权重（总和为0.65）
//   alpha: 0.6, // 更注重夏普比率
//   iterations: 5000, // 减少迭代次数以加快计算
// });

// // 运行示例查看不同策略效果
// optimizer.runExamples();
