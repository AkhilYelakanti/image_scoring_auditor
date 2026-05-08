/**
 * Web Worker for heavy data processing
 * offloads comparison, filtering, and sorting of 50k+ records from the main thread
 */
self.onmessage = (e) => {
  const { prevData, newData, config } = e.data;
  const { filter = 'all', searchQuery = '', sortConfig = { key: 'upc', order: 'asc' }, userStatus = {} } = config || {};
  
  const getUpcValue = (item) => {
    if (!item) return '';
    const keys = Object.keys(item);
    // Prioritize exact matches or common names
    const exactMatch = keys.find(k => ['UPC', 'GTIN', 'EAN', 'ID', 'IDENTIFIER'].includes(k.toUpperCase()));
    if (exactMatch) return String(item[exactMatch]).trim();

    const upcKey = keys.find(k => {
      const lower = k.toLowerCase();
      return lower.includes('upc') || 
             lower.includes('gtin') || 
             lower.includes('identifier') || 
             lower.includes('ean') ||
             lower.includes('barcode') ||
             (lower.includes('item') && lower.includes('id')) ||
             (lower.includes('item') && lower.includes('code'));
    });
    const val = upcKey ? item[upcKey] : (item.UPC || item.GTIN || item.ID);
    return val ? String(val).trim() : '';
  };

  const getImageUrl = (item) => {
    if (!item) return '';
    const keys = Object.keys(item);
    
    // 1. Look for known image/url keys that have a value
    const urlKey = keys.find(k => {
      const lower = k.toLowerCase();
      const val = String(item[k] || '').trim();
      return val.length > 0 && 
             (lower.includes('url') || lower.includes('image') || lower.includes('link') || lower.includes('http') || lower.includes('path')) && 
             !lower.includes('name') && !lower.includes('type');
    });
    if (urlKey) return String(item[urlKey]).trim();

    // 2. Look for any key whose value starts with http
    const httpKey = keys.find(k => String(item[k] || '').trim().toLowerCase().startsWith('http'));
    if (httpKey) return String(item[httpKey]).trim();

    return String(item.SOURCE_IMAGE_URL || item.IMAGE_URL || '').trim();
  };

  const prevMap = new Map();
  if (prevData) {
    prevData.forEach(d => {
      const u = getUpcValue(d);
      if (u) {
        const foundUrl = getImageUrl(d);
        if (foundUrl) d.SOURCE_IMAGE_URL = foundUrl;
        prevMap.set(u, d);
      }
    });
  }

  const newMap = new Map();
  if (newData) {
    newData.forEach(d => {
      const u = getUpcValue(d);
      if (u) {
        const foundUrl = getImageUrl(d);
        if (foundUrl) d.SOURCE_IMAGE_URL = foundUrl;
        newMap.set(u, d);
      }
    });
  }

  const allUpcs = new Set([
    ...Array.from(prevMap.keys()),
    ...Array.from(newMap.keys())
  ]);

  let results = Array.from(allUpcs).map(upc => {
    const prev = prevMap.get(upc) || null;
    const current = newMap.get(upc) || null;
    
    const changes = [];
    if (prev && current) {
      const pName = String(prev.CSOR_IMAGE_NAME || '').trim();
      const cName = String(current.CSOR_IMAGE_NAME || '').trim();
      const pUrl = String(prev.SOURCE_IMAGE_URL || '').trim();
      const cUrl = String(current.SOURCE_IMAGE_URL || '').trim();

      if (pName !== cName) changes.push('IMAGE_NAME');
      if (pUrl !== cUrl) changes.push('URL');
      if (String(prev.IMAGE_STATUS || '').trim() !== String(current.IMAGE_STATUS || '').trim()) changes.push('STATUS');
    } else if (prev || current) {
      changes.push(prev ? 'DELETED' : 'CREATED');
    }

    const auditStatus = userStatus[upc] || 'pending';

    return {
      upc,
      previous: prev,
      current: current,
      hasChanged: changes.length > 0,
      changes,
      auditStatus
    };
  });

  // Calculate full stats before filtering
  const stats = {
    total: results.length,
    changed: results.filter(r => r.hasChanged).length,
    same: results.filter(r => !r.hasChanged && r.previous && r.current).length,
    reviewed: Object.values(userStatus).filter(s => s === 'reviewed').length,
    ignored: Object.values(userStatus).filter(s => s === 'ignored').length,
    totalPrev: prevMap.size,
    totalNew: newMap.size,
    totalMerged: results.length
  };

  // APPLY FILTERS & SORTING IN WORKER
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    results = results.filter(r => 
      r.upc.toLowerCase().includes(q) || 
      (r.current?.CSOR_IMAGE_NAME || '').toLowerCase().includes(q) ||
      (r.previous?.CSOR_IMAGE_NAME || '').toLowerCase().includes(q)
    );
  }

  if (filter !== 'all') {
    if (filter === 'changed') results = results.filter(r => r.hasChanged);
    else results = results.filter(r => r.auditStatus === filter);
  }

  if (sortConfig.key) {
    results.sort((a, b) => {
      let valA = '';
      let valB = '';

      switch (sortConfig.key) {
        case 'upc':
          valA = a.upc;
          valB = b.upc;
          break;
        case 'prev_name':
          valA = a.previous?.CSOR_IMAGE_NAME || '';
          valB = b.previous?.CSOR_IMAGE_NAME || '';
          break;
        case 'curr_name':
          valA = a.current?.CSOR_IMAGE_NAME || '';
          valB = b.current?.CSOR_IMAGE_NAME || '';
          break;
        case 'status':
          valA = a.hasChanged ? 'changed' : 'same';
          valB = b.hasChanged ? 'changed' : 'same';
          break;
        case 'auditStatus':
          valA = a.auditStatus;
          valB = b.auditStatus;
          break;
      }

      if (valA < valB) return sortConfig.order === 'asc' ? -1 : 1;
      if (valA > valB) return sortConfig.order === 'asc' ? 1 : -1;
      return 0;
    });
  }

  self.postMessage({ 
    results, 
    stats
  });
};
