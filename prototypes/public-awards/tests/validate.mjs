import fs from "node:fs";
import assert from "node:assert/strict";

const data=JSON.parse(fs.readFileSync(new URL("../awards.json",import.meta.url),"utf8"));
assert.equal(data.schema_version,"1.0");
assert.equal(data.status,"review_only");
assert.ok(Array.isArray(data.records)&&data.records.length===6);
const ids=new Set(),keys=new Set();
for(const row of data.records){
  for(const field of ["id","book","author","isbn","book_key","product_url","award","category","result","organization","source_url"]){
    assert.equal(typeof row[field],"string",field+" must be text");
    assert.ok(row[field].trim(),field+" cannot be blank");
  }
  assert.ok(Number.isInteger(row.year)&&row.year>=1900&&row.year<=2100);
  for(const urlField of ["source_url","product_url"]){
    assert.equal(new URL(row[urlField]).protocol,"https:",urlField+" must be HTTPS");
  }
  assert.ok(!ids.has(row.id),"duplicate id: "+row.id);ids.add(row.id);
  const key=[row.book_key,row.year,row.award,row.category,row.result].join("|").toLowerCase();
  assert.ok(!keys.has(key),"duplicate recognition: "+key);keys.add(key);
  assert.ok(/^97[89][0-9]{10}$/.test(row.isbn),"ISBN-13 expected: "+row.id);
}
assert.equal(new Set(data.records.map(row=>row.book_key)).size,4);
console.log("Awards pilot validated: "+data.records.length+" recognitions across four books.");
