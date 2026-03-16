function normalizeValue_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function normalizeExtractionType_(value) {
  const normalized = String(value || 'tag')
    .replace(/["'\\]/g, '')
    .trim()
    .toLowerCase();

  if (normalized === 'account') return 'account';
  if (normalized === 'wallet') return 'wallet';
  if (normalized === 'tag') return 'tag';

  return normalized || 'tag';
}

function resolveExtractionType_(payload) {
  const hasSelection =
    payload &&
    payload.selectionValue !== undefined &&
    payload.selectionValue !== null &&
    String(payload.selectionValue).trim() !== '';

  if (!hasSelection) {
    return 'tag';
  }

  return normalizeExtractionType_(payload && payload.extractionType);
}

function matchesFilterValue_(rawValue, expectedValue) {
  return normalizeValue_(rawValue) === normalizeValue_(expectedValue);
}

function maxDefined_(max, value) {
  return value && value > max ? value : max;
}
