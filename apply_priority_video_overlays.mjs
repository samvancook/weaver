import { readFile, writeFile } from "node:fs/promises";

const serverPath = new URL("./server.mjs", import.meta.url);
let source = await readFile(serverPath, "utf8");

if (source.includes('id: "northbeast-regional-2023-finals"')) {
  console.log("NorthBeast priority-video overlay already applied.");
  process.exit(0);
}

const northBeastSet = `  ["northbeast-regional-2023-finals", {
    id: "northbeast-regional-2023-finals",
    label: "NorthBeast Regional 2023 - Finals",
    eventName: "NorthBeast Regional 2023 - Finals",
    folderId: "1A-g6R5uJtEuunl7FAahNqkHqgRQ-IqMT",
    includeFileIds: [
      "1OIGkZoA0Jm-wza83JuUhG0ozNq4ywKVN",
      "1m6Xxk9Skbzwz4zv2xfJjeblF4F-Qz3UK",
      "19QjpNBGfX6ywTCrkdNu8hjo1E8YZrc74",
      "1kK55D45gltxFLt0WzGy35Y9tPwxxBagE",
      "1fmZmFZATkeiJoYIvXayn95_XUixUuGe0",
      "1BCIJvkZXrHKgfy-LHMEKKV0hkndBKWrm",
      "17ZiJa_RTrEMfLKhCZAZn9BWL_YuHnAgN",
      "1DHpjrKQvTJFOtOjQgyYUL_dtrz1KFBUk",
      "1obrpQAFEtY_1ec0AsHeDH1dZXKJHrF2B",
      "1hZCnqKhSvrKnY27yIe_Q_jA9Kvs-oSMU",
      "14GVeMnc75e-m0b45lLLcO5p-uHNpyQ0A"
    ],
    fileMetadataById: {
      "1OIGkZoA0Jm-wza83JuUhG0ozNq4ywKVN": { author: "Kai Wallin", poemTitle: "Love Poem" },
      "1m6Xxk9Skbzwz4zv2xfJjeblF4F-Qz3UK": { author: "Logan Lopez", poemTitle: "Road Trip With The Memory Of My Girlhood and My Imagined Boyhood" },
      "19QjpNBGfX6ywTCrkdNu8hjo1E8YZrc74": { author: "Art Collins", poemTitle: "Good Morning... no response." },
      "1kK55D45gltxFLt0WzGy35Y9tPwxxBagE": { author: "Will C.", poemTitle: "Love Me Some Me" },
      "1fmZmFZATkeiJoYIvXayn95_XUixUuGe0": { author: "Kenny Bradley", poemTitle: "To Pimp a Caterpillar" },
      "1BCIJvkZXrHKgfy-LHMEKKV0hkndBKWrm": { author: "Ren L[i]u", poemTitle: "A Second Coming" },
      "17ZiJa_RTrEMfLKhCZAZn9BWL_YuHnAgN": { author: "Seth Larbi", poemTitle: "Ode to Chief Keef, Ending In Survivor's Guilt" },
      "1DHpjrKQvTJFOtOjQgyYUL_dtrz1KFBUk": { author: "Meaghan Ford", poemTitle: "The Body As Reclamation" },
      "1obrpQAFEtY_1ec0AsHeDH1dZXKJHrF2B": { author: "Christopher Clauss", poemTitle: "Purple Shirt Yellow Balloon" },
      "1hZCnqKhSvrKnY27yIe_Q_jA9Kvs-oSMU": { author: "Matthew Richards", poemTitle: "Hospitality" },
      "14GVeMnc75e-m0b45lLLcO5p-uHNpyQ0A": { author: "Mica Rich", poemTitle: "Dirty Mouths" }
    }
  }],
`;

const helpers = `
function isPriorityVideoFileIncluded(file, set) {
  const fileId = cleanSheetWhitespace(file?.id);
  if (Array.isArray(set?.includeFileIds) && !set.includeFileIds.includes(fileId)) {
    return false;
  }
  if (set?.excludeNoPoem && PRIORITY_VIDEO_NO_POEM_PATTERN.test(cleanSheetWhitespace(file?.name))) {
    return false;
  }
  return true;
}

function getPriorityVideoIdentity(file, set) {
  const override = set?.fileMetadataById?.[cleanSheetWhitespace(file?.id)];
  if (override) {
    return {
      author: cleanSheetWhitespace(override.author),
      poemTitle: cleanSheetWhitespace(override.poemTitle)
    };
  }
  return parsePriorityVideoFileName(file?.name);
}
`;

function replaceExactlyOnce(label, before, after) {
  const first = source.indexOf(before);
  const last = source.lastIndexOf(before);
  if (first < 0) {
    throw new Error(`Overlay failed: ${label} target was not found.`);
  }
  if (first !== last) {
    throw new Error(`Overlay failed: ${label} target was not unique.`);
  }
  source = source.replace(before, after);
}

replaceExactlyOnce(
  "priority set insertion",
  '  ["publishers-poetry-slam-2026-camera-y", {',
  `${northBeastSet}  ["publishers-poetry-slam-2026-camera-y", {`
);

replaceExactlyOnce(
  "priority helper insertion",
  'const PRIORITY_VIDEO_NO_POEM_PATTERN = /\\bno\\s+poem\\b/i;',
  `${helpers}\nconst PRIORITY_VIDEO_NO_POEM_PATTERN = /\\bno\\s+poem\\b/i;`
);

replaceExactlyOnce(
  "priority set file filtering",
  `  const files = (await listDriveFolderVideoFiles(set.folderId)).filter(file => (\n    !set.excludeNoPoem || !PRIORITY_VIDEO_NO_POEM_PATTERN.test(cleanSheetWhitespace(file.name))\n  ));`,
  `  const files = (await listDriveFolderVideoFiles(set.folderId))\n    .filter(file => isPriorityVideoFileIncluded(file, set));`
);

replaceExactlyOnce(
  "priority item metadata",
  '      ...parsePriorityVideoFileName(file.name),',
  '      ...getPriorityVideoIdentity(file, set),'
);

replaceExactlyOnce(
  "priority progress filtering",
  `      files: (await listDriveFolderVideoFiles(set.folderId)).filter(file => (\n        !set.excludeNoPoem || !PRIORITY_VIDEO_NO_POEM_PATTERN.test(cleanSheetWhitespace(file.name))\n      ))`,
  `      files: (await listDriveFolderVideoFiles(set.folderId))\n        .filter(file => isPriorityVideoFileIncluded(file, set))`
);

replaceExactlyOnce(
  "priority review source validation",
  `  const file = folderFiles.find(candidate => (\n    cleanSheetWhitespace(candidate.id) === sourceFileId\n    && (!set.excludeNoPoem || !PRIORITY_VIDEO_NO_POEM_PATTERN.test(cleanSheetWhitespace(candidate.name)))\n  ));`,
  `  const file = folderFiles.find(candidate => (\n    cleanSheetWhitespace(candidate.id) === sourceFileId\n    && isPriorityVideoFileIncluded(candidate, set)\n  ));`
);

await writeFile(serverPath, source, "utf8");
console.log("Applied NorthBeast Regional 2023 - Finals priority-video overlay.");
