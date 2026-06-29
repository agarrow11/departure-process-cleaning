import { readFileSync } from "node:fs"

// Proper CSV header parse (first record only), handling quotes + embedded newlines.
function firstRecord(text) {
  const fields = []
  let cur = ""
  let inQ = false
  for (let i = 0; i < text.length; i++) {
    const c = text[i]
    if (inQ) {
      if (c === '"') {
        if (text[i + 1] === '"') { cur += '"'; i++ }
        else inQ = false
      } else cur += c
    } else {
      if (c === '"') inQ = true
      else if (c === ",") { fields.push(cur); cur = "" }
      else if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++
        fields.push(cur)
        return fields
      } else cur += c
    }
  }
  fields.push(cur)
  return fields
}

const text = readFileSync("data/reference_output.csv", "utf8")
const cols = firstRecord(text)
console.log("TOTAL COLUMNS:", cols.length)
console.log("\n=== BB_ columns ===")
cols.forEach((c, i) => { if (c.startsWith("BB_")) console.log(i, JSON.stringify(c)) })
console.log("\n=== .1 / dup-suffixed columns ===")
cols.forEach((c, i) => { if (/\.\d+$/.test(c)) console.log(i, JSON.stringify(c)) })
console.log("\n=== last 15 columns ===")
cols.slice(-15).forEach((c, i) => console.log(cols.length - 15 + i, JSON.stringify(c)))
