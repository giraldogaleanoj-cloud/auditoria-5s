/* Almacenamiento local en el dispositivo (IndexedDB): auditorías y fotos. */
(function () {
  const DB_NAME = 'auditorias5s';
  const VERSION = 1;
  let dbp = null;

  function open() {
    if (dbp) return dbp;
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains('audits')) db.createObjectStore('audits', { keyPath: 'id' });
        if (!db.objectStoreNames.contains('photos')) {
          const s = db.createObjectStore('photos', { keyPath: 'id' });
          s.createIndex('auditId', 'auditId');
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbp;
  }

  function tx(store, mode, fn) {
    return open().then(db => new Promise((resolve, reject) => {
      const t = db.transaction(store, mode);
      const s = t.objectStore(store);
      let result;
      Promise.resolve(fn(s, t)).then(r => { result = r; });
      t.oncomplete = () => resolve(result);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    }));
  }

  const req2p = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });

  const DB = {
    allAudits: () => tx('audits', 'readonly', s => req2p(s.getAll())),
    getAudit: id => tx('audits', 'readonly', s => req2p(s.get(id))),
    putAudit: a => tx('audits', 'readwrite', s => { s.put(a); }),
    async deleteAudit(id) {
      const photos = await DB.photosFor(id);
      await tx('photos', 'readwrite', s => photos.forEach(p => s.delete(p.id)));
      await tx('audits', 'readwrite', s => { s.delete(id); });
    },
    photosFor: auditId => tx('photos', 'readonly', s => req2p(s.index('auditId').getAll(auditId))),
    allPhotos: () => tx('photos', 'readonly', s => req2p(s.getAll())),
    putPhoto: p => tx('photos', 'readwrite', s => { s.put(p); }),
    deletePhoto: id => tx('photos', 'readwrite', s => { s.delete(id); }),
  };

  window.DB = DB;
})();
