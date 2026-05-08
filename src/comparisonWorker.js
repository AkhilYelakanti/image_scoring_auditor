/**
 * Web Worker for heavy data processing
 * offloads comparison of 50k+ records from the main thread
 */
self.onmessage = (e) => {
  const { prevData, newData } = e.data;
  
  const getUpcValue = (item) => {
    if (!item) return '';
    const keys = Object.keys(item);
    const upcKey = keys.find(k => {
      const lower = k.toLowerCase();
      return lower.includes('upc') || 
             lower.includes('gtin') || 
             lower.includes('identifier') || 
             lower.includes('ean') ||
             (lower.includes('item') && lower.includes('id')) ||
             (lower.includes('item') && lower.includes('code'));
    });
    const val = upcKey ? item[upcKey] : (item.UPC || item.GTIN || item.ID);
    return String(val || '').trim();
  };

  const prevMap = new Map();
  if (prevData) {
    prevData.forEach(d => {
      const u = getUpcValue(d);
      if (u) prevMap.set(u, d);
    });
  }

  const newMap = new Map();
  if (newData) {
    newData.forEach(d => {
      const u = getUpcValue(d);
      if (u) newMap.set(u, d);
    });
  }

  const allUpcs = new Set([
    ...Array.from(prevMap.keys()),
    ...Array.from(newMap.keys())
  ]);

  const results = Array.from(allUpcs).map(upc => {
    const prev = prevMap.get(upc) || null;
    const current = newMap.get(upc) || null;
    
    const changes = [];
    if (prev && current) {
      if (String(prev.CSOR_IMAGE_NAME || '').trim() !== String(current.CSOR_IMAGE_NAME || '').trim()) changes.push('IMAGE_NAME');
      if (String(prev.SOURCE_IMAGE_URL || '').trim() !== String(current.SOURCE_IMAGE_URL || '').trim()) changes.push('URL');
      if (String(prev.IMAGE_STATUS || '').trim() !== String(current.IMAGE_STATUS || '').trim()) changes.push('STATUS');
      if (String(prev.PKG_NAME || '').trim() !== String(current.PKG_NAME || '').trim()) changes.push('PACKAGE');
    } else if (prev || current) {
      changes.push(prev ? 'DELETED' : 'CREATED');
    }

    return {
      upc,
      previous: prev,
      current: current,
      hasChanged: changes.length > 0,
      changes,
      auditStatus: 'pending'
    };
  });

  self.postMessage(results);
};
