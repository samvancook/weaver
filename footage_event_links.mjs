import {validateEventManifest} from "./footage_event_intake.mjs";
// Reviewed event correspondence. These are asset relationships, not file aliases:
// never replace a review's sourceFileId or infer a canonical text ID from a title.
export const REVIEWED_EVENT_LINKS = [{
  inventoryFolderId: '1xeos0-8HxDBHdm8NP0r-KHgsex5SdrrP',
  prioritySetId: 'pages-matam-san-diego-2026-labc',
  inventoryMediaType: 'OM',
  curationMediaType: 'FV',
  mediaTypeEvidence: 'User confirmed horizontal PAGES files are OM and vertical files are FV on 2026-10-30; Button branding and Rudy credit visually checked on all four FVs.',
  evidence: 'Inventory source folder PAGES Matam - San Diego 2026; existing FV folder explicitly names PAGES footage from Rudy Francisco; four reviewed title correspondences.',
  pairs: [
    ['1pHdxNXRNgbdzRecfUYYsvNt1VyZS7Wws', '1Te0zeh7tmd820BBoUw2HOcDCbkBMjGwH'],
    ['1y7M7_eRYCO_FEtEvuERGRvYvrrvsiqyL', '1eoqefV0zfkzR9fEZ8Zr4HFY5rsGVw-FA'],
    ['1wHoJouPMLTtIEwv58FQp39gkDWpBZ5ni', '14NdOdzkV1F4dScR2QYk8k20WTAlyJF2g'],
    ['1sD5kbncP1z9G2xb_lHyjpLd9FLCHs-zf', '1MwOqcF4HWhB2ukcs58mzbbRU2xunipTO']
  ]

}, {
  inventoryFolderId: '1iRYP99bEhn8hs0DUzxatfJG6KqUWKGTy',
  prioritySetId: 'mpmu-2026-finals',
  inventoryMediaType: 'unknown',
  curationMediaType: 'FV',
  mediaTypeEvidence: 'All 16 curation IDs verified in the inventory FV Link Lookup Table and live Weaver set on 2026-11-03. Intake-source media role remains unclassified.',
  evidence: 'Closed Vertical Versions - MPMU 2026 Finals: 16 reviewed source/finished-video correspondences within the event, exact filenames except documented Tilt/Shift versus Tilt (slash) Shift; source and finished exports have different checksums. Closure note explicitly reuses the original curation form. These are relationships, not identical files.',
  pairs: [
    [
        "13ARarw1WrKRSy4lBC9ieGrD4o5jIzQUk",
        "1K4-dVsiUbHT-QjmDpdawt0onxBKjnSF9"
    ],
    [
        "1sTmz1uODYL7ci1IbvsHkkKGyS8nap0Qw",
        "1hzNfeC-mPE2guXcaTd8K6asdFVcGuesY"
    ],
    [
        "13_XemBuD9gfnlfznLGRZRG4LkkX7jRWp",
        "1N3bfVrUH84ducO2UZE_50PU4ndJkIrNW"
    ],
    [
        "1YLwGg3RW42qL9DGno4Iqorvp1JgZkMJL",
        "10_HxG-d5ibQQ27hVUIwQAtFy16yeQKL9"
    ],
    [
        "17t7F-3OOpi2QEKqMzgFkywP15XmgFThC",
        "1vFft2_pFOu7Fl3vuHXSI_4Z7098dQL9i"
    ],
    [
        "19VQ2EIEXaCCbrSNj8wvlHyud6SxpvH58",
        "1Oleo9uICFVFFOAfMSxkk2jvuQRUCexgN"
    ],
    [
        "1bT5k811M_JLJydA0qIOeyb1KEECSOXkW",
        "1vrQHrnnL4KTMPPDcsTqHWOetukM4be75"
    ],
    [
        "1P1KL1xxGJogw4rjW1kjeLNS7876YjTjc",
        "1gO9PBngB2lo0-8uC-TrDHuhxkk_QvZ5n"
    ],
    [
        "1y-s3bXElyONSdOmbUvQ0ktyW02Hq45wj",
        "1SqPZRP9-Iifz0sZBFY3XW_1I9LwUWtbj"
    ],
    [
        "1t2t1DASFpqDZ-PN8nnm7Bgf90mjippRB",
        "1vW708q89C8fKxxi2S5Gi2oOBTXtFOQG5"
    ],
    [
        "1jLeHN8LbwmJRcttwM6GFNej3KJDlvMu8",
        "1Eyegyiw6-iZm13Crp8G-EsDsnOQf7CW2"
    ],
    [
        "1D98qUMC6nyvW-PHDDftuOCYZBlWT1UCh",
        "1NnZYeC95uPLriHFgGwBQ7qf3Y0QrDWXx"
    ],
    [
        "1qHzd4DczQ_bhUPXmsc_CHO7vsDqGfxvQ",
        "1tB_er-htHAxFph7AhqwMD09aTYANWbZ4"
    ],
    [
        "1A-tCKF4HEqsjkseJja-jeXdu1e1Abr1N",
        "17De4UPUtWViGM5_nBjmLp0JxPiwWzVA6"
    ],
    [
        "1E08Mm0PDVHJyYUNPgkZqmVp0AzZWPYeH",
        "1OXuuyUgW1jYs1N3IfSOLWdpBcHRqsow9"
    ],
    [
        "1BNk2_B_dO1OYUYaXEqRVdUJRqt6MlNWI",
        "1Rj0QA-yW4uZjmrbjugh8fQaF4mVSinLT"
    ]
]
}];

export async function resolveEventTarget(manifest, sets, listFiles, parseFileId, registry = REVIEWED_EVENT_LINKS) {
  const reviewed = registry.find(link => link.inventoryFolderId === manifest.omFolderId);
  if (reviewed) {
    const set = sets.get(reviewed.prioritySetId);
    if (!set || set.inventoryEvent) throw new Error('The reviewed existing Weaver event is unavailable. No new queue was created.');
    const available = new Set((await listFiles(set)).map(f => f.id));
    const pairs = new Map(reviewed.pairs);
    return {mode:'existing_event', prioritySetId:set.id, label:set.label,
      evidence:reviewed.evidence, inventoryMediaType:reviewed.inventoryMediaType,
      curationMediaType:reviewed.curationMediaType, mediaTypeEvidence:reviewed.mediaTypeEvidence,
      pairs, available};
  }
  // Exact file IDs can safely discover an existing queue without fuzzy title joins.
  const ids = manifest.titles.map(t => parseFileId(t.contentLink)).filter(Boolean);
  const ownId = validateEventManifest(manifest).eventId;
  const candidates = [...sets.values()].filter(set => !(set.inventoryEvent && set.id === ownId));
  const matches = await Promise.all(candidates.map(async set => ({set,
    ids:new Set((await listFiles(set)).map(f => f.id))})));
  const touched = matches.filter(x => ids.some(id => x.ids.has(id)));
  if (!touched.length) return {mode:'new_event'};
  if (touched.some(x => x.set.inventoryEvent)) throw new Error('These files already belong to another inventory event. Existing work was preserved; connect the events explicitly.');
  const complete = touched.filter(x => ids.length === manifest.titles.length && ids.every(id => x.ids.has(id)));
  if (touched.length !== 1 || complete.length !== 1) {
    throw new Error('Some videos already belong to another Weaver queue, or match multiple queues. Existing work was preserved; preview needs an event connection.');
  }
  const {set,ids:available} = complete[0];
  return {mode:'existing_event',prioritySetId:set.id,label:set.label,
    evidence:'Exact Drive file IDs already present in this Weaver event.',
    pairs:new Map(ids.map(id => [id,id])),available};
}
