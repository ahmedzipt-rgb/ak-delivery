function akV4ApplyTx_(data) {
  akV4EnsureInitialized_();
  if (!data || !data.opId || !Array.isArray(data.changes)) throw new Error('Invalid V4 transaction');
  const opId = String(data.opId);
  const existing = akV4FindOp_(opId);
  if (existing) return existing;

  const actor = { role: String(data.actorRole || ''), id: String(data.actorId || ''), name: String(data.actorName || '') };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const second = akV4FindOp_(opId);
    if (second) return second;

    const table = akV4EntityTable_();
    const conflicts = [];
    const prepared = [];
    const seen = new Set();

    data.changes.forEach((raw, idx) => {
      const type = String((raw && raw.type) || '');
      const id = String((raw && raw.id) || '');
      if (!(AKV4_TYPE_SET[type] || (type === AKV4_META_TYPE && id === AKV4_META_ID))) {
        throw new Error('Unsupported entity type ' + type);
      }
      if (!id) throw new Error('Missing entity id');
      const key = akV4Key_(type, id);
      if (seen.has(key)) return;
      seen.add(key);

      const current = table.map.get(key) || null;
      const currentRev = current ? current.rev : 0;
      const expected = akV4Num_(raw.expectedRev, 0);

      if (!raw.deleted && current && !current.deleted && current.entity && raw.entity) {
        if (JSON.stringify(current.entity) === JSON.stringify(raw.entity)) return;
      }

      if (expected !== currentRev) {
        if (actor.role === 'worker' && type === 'orders' && !raw.deleted && current && !current.deleted && current.entity) {
          const mergedEntity = akV4NormalizeWorkerOrder_(current.entity, raw.entity, actor.id);
          prepared.push({ idx, type, id, key, deleted: false, entity: mergedEntity, current });
          return;
        }
        conflicts.push({
          type,
          id,
          expectedRev: expected,
          currentRev,
          deleted: current ? current.deleted : false,
          entity: current && !current.deleted ? current.entity : null
        });
        return;
      }

      const safeEntity = akV4AuthorizeChange_(raw, current, actor);
      if (!raw.deleted) {
        if (!safeEntity || typeof safeEntity !== 'object') throw new Error('Invalid entity ' + type + ':' + id);
        if (type !== AKV4_META_TYPE && String(safeEntity.id) !== id) throw new Error('Entity id mismatch ' + type + ':' + id);
      }
      prepared.push({ idx, type, id, key, deleted: !!raw.deleted, entity: safeEntity, current });
    });

    if (conflicts.length) {
      const rev = akV4CurrentRev_();
      const details = { code: 'cloud_conflict', currentRev: rev, conflicts };
      akV4WriteOp_(opId, 'conflict', rev, 'Entity revision conflict', details);
      return akV4FindOp_(opId);
    }

    if (!prepared.length) {
      const rev = akV4CurrentRev_();
      akV4WriteOp_(opId, 'success', rev, '', { empty: true });
      return akV4FindOp_(opId);
    }

    const newRev = akV4CurrentRev_() + 1;
    const now = new Date();
    const nowIso = now.toISOString();
    const appendEntityRows = [];
    const eventRows = [];
    const responseChanges = [];

    prepared.forEach((p, seq) => {
      const json = p.deleted ? '' : JSON.stringify(p.entity);
      const row = [p.type, p.id, newRev, json, p.deleted, now];
      if (p.current) table.sheet.getRange(p.current.row, 1, 1, 6).setValues([row]);
      else appendEntityRows.push(row);
      eventRows.push([newRev, seq + 1, opId, now, actor.role, actor.id, String(data.kind || 'sync'), p.type, p.id, p.deleted, newRev, json]);
      responseChanges.push({ type: p.type, id: p.id, deleted: p.deleted, entityRev: newRev, entity: p.deleted ? null : akV4Clone_(p.entity) });
    });

    if (appendEntityRows.length) {
      const start = table.sheet.getLastRow() + 1;
      table.sheet.getRange(start, 1, appendEntityRows.length, 6).setValues(appendEntityRows);
    }
    const evs = akV4EventsSheet_();
    const evStart = evs.getLastRow() + 1;
    evs.getRange(evStart, 1, eventRows.length, 12).setValues(eventRows);

    akV4SetMetaMany_({ currentRev: String(newRev), updatedAt: nowIso, lastOpId: opId });
    akV4WriteOp_(opId, 'success', newRev, '', { changeCount: responseChanges.length });
    SpreadsheetApp.flush();
    return { opId, status: 'success', rev: newRev, error: '', details: { changeCount: responseChanges.length }, changes: responseChanges, updatedAt: nowIso };
  } catch (err) {
    const rev = akV4CurrentRev_();
    try { akV4WriteOp_(opId, 'error', rev, String((err && err.message) || err), { code: err && err.code ? err.code : null }); } catch (_) {}
    return akV4FindOp_(opId) || { opId, status: 'error', rev, error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
