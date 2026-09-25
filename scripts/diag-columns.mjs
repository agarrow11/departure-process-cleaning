import { readFileSync } from "node:fs"
import { runPipeline } from "../lib/pipeline.ts"

function ab(path) {
  const buf = readFileSync(path)
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
}

const files = {
  oldSurvey: ab("data/Old-Survey_052826-597993.xlsx"),
  newSurvey: ab("data/New-Survey_070126-Copy-34c6c1.xlsx"),
  exitInterview: ab("data/Exit-Interview_052826-623ea3.xlsx"),
  beyondBain: ab("data/Beyond-Bain-Extract_060226-9e98f9.xlsx"),
  deptHierarchy: ab("data/dept_hierarchy_20260601-aa7426.xlsx"),
  geoHierarchy: ab("data/geographic_hierarchy_20260501-0691f7.xlsx"),
  pdGrade: ab("data/PD-Grade-Mapping_20260403.csv"),
}
const previousFiles = { ...files, newSurvey: ab("data/New-Survey_060226-1a7f5b.xlsx") }

const before = await runPipeline(previousFiles)
const after = await runPipeline(files)

const beforeCols = Object.keys(before.rows[0])
const afterCols = Object.keys(after.rows[0])
const priorSet = new Set(beforeCols)
const afterSet = new Set(afterCols)

console.log("before columns:", beforeCols.length)
console.log("after columns:", afterCols.length)
console.log("\nColumns ADDED (in after, not before):")
console.log(afterCols.filter((c) => !priorSet.has(c)))
console.log("\nColumns REMOVED (in before, not after):")
console.log(beforeCols.filter((c) => !afterSet.has(c)))

console.log("\nbefore populationRows:", before.stats.populationRows)
console.log("after populationRows:", after.stats.populationRows)
console.log("before newSurveyRows:", before.stats.newSurveyRows)
console.log("after newSurveyRows:", after.stats.newSurveyRows)
console.log("before combinedSurveyRows:", before.stats.combinedSurveyRows)
console.log("after combinedSurveyRows:", after.stats.combinedSurveyRows)
console.log("before bothSourcesCount/surveyOnlyCount/eiOnlyCount:", before.stats.bothSourcesCount, before.stats.surveyOnlyCount, before.stats.eiOnlyCount)
console.log("after bothSourcesCount/surveyOnlyCount/eiOnlyCount:", after.stats.bothSourcesCount, after.stats.surveyOnlyCount, after.stats.eiOnlyCount)
