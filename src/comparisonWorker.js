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
    const exactMatch = keys.find(k => ['UPC', 'GTIN', 'EAN', 'ID', 'IDENTIFIER', 'UPC / IDENTIFIER'].includes(k.toUpperCase()));
    if (exactMatch && item[exactMatch]) return String(item[exactMatch]).trim();

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
    
    // 1. Look for known image fields first
    const knownKeys = ['SOURCE_IMAGE_URL', 'IMAGE_URL', 'URL', 'IMAGE', 'SOURCE_URL'];
    for (const key of knownKeys) {
      const found = keys.find(k => k.toUpperCase() === key);
      if (found && String(item[found]).trim().startsWith('http')) {
        return String(item[found]).trim();
      }
    }

    // 2. Look for any value starting with http
    const httpKey = keys.find(k => {
      const val = String(item[k] || '').trim().toLowerCase();
      return val.startsWith('http');
    });
    if (httpKey) return String(item[httpKey]).trim();

    // 3. Look for anything with common image extensions or hosting keywords
    const imageKey = keys.find(k => {
      const val = String(item[k] || '').trim().toLowerCase();
      return val.length > 0 && 
             (val.includes('cloudinary') || val.includes('amazonaws') || val.endsWith('.jpg') || val.endsWith('.png') || val.endsWith('.jpeg')) &&
             !k.toLowerCase().includes('name') && !k.toLowerCase().includes('title');
    });
    if (imageKey) return String(item[imageKey]).trim();

    // Fallback to any URL segment
    const fallbackKey = keys.find(k => {
      const lower = k.toLowerCase();
      return (lower.includes('url') || lower.includes('link') || lower.includes('path')) && 
             !lower.includes('name') && !lower.includes('title');
    });
    if (fallbackKey) return String(item[fallbackKey]).trim();

    return '';
  };

  // Helper to extract Type and Version from image name
  const parseImageName = (name) => {
    if (!name) return { version: '', type: '' };
    // Handle multiple extensions like .jpg.jpg or .JPG
    const parts = String(name).split('.');
    const clean = parts[0]; 
    const segments = clean.split('_');
    
    // Pattern usually: UPC_VERSION_TYPE
    if (segments.length >= 3) {
      return {
        version: segments[segments.length - 2],
        type: segments[segments.length - 1]
      };
    }
    if (segments.length === 2) {
      return { version: segments[1], type: '' };
    }
    return { version: '', type: '' };
  };

  const prevMap = new Map();
  if (prevData) {
    prevData.forEach(d => {
      const u = getUpcValue(d);
      if (u) {
        const foundUrl = getImageUrl(d);
        if (foundUrl) d.SOURCE_IMAGE_URL = foundUrl;
        d.CSOR_IMAGE_NAME = d.CSOR_IMAGE_NAME || d.IMAGE_NAME || d.FILENAME || '';
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
        d.CSOR_IMAGE_NAME = d.CSOR_IMAGE_NAME || d.IMAGE_NAME || d.FILENAME || '';
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
      const pParts = parseImageName(prev.CSOR_IMAGE_NAME);
      const cParts = parseImageName(current.CSOR_IMAGE_NAME);

      if (pParts.type !== cParts.type) changes.push('TYPE');
      if (pParts.version !== cParts.version) changes.push('VERSION');
      
      // We are no longer tracking URL or general IMAGE_NAME changes per request
      if (String(prev.IMAGE_STATUS || '').trim() !== String(current.IMAGE_STATUS || '').trim()) {
        changes.push('STATUS');
      }
    }
    // CREATED and DELETED logic removed - these rows will now show as "Matched" since they have no changes in the tracked fields

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
    ignoredNoImages: 0,
    totalPrev: prevMap.size,
    totalNew: newMap.size,
    totalMerged: results.length
  };

  // Filter out results where BOTH versions have no image
  const initialCount = results.length;
  results = results.filter(r => {
    const hasPrevImg = !!r.previous?.SOURCE_IMAGE_URL;
    const hasCurrImg = !!r.current?.SOURCE_IMAGE_URL;
    return hasPrevImg || hasCurrImg;
  });
  stats.ignoredNoImages = initialCount - results.length;
  stats.totalMerged = results.length;

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
