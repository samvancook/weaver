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
