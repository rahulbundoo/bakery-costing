// Bakery Costing App script with repository template loading and CRUD flows.
(function () {
  const STORAGE = {
    setup: 'setupValues',
    ingredients: 'ingredientsData',
    orderLog: 'orderLog',
    lastOrderId: 'lastOrderId',
    templatePath: 'templatePath'
  };

  const TEMPLATE_CANDIDATES = [
    'bakery_costing_template_final.xlsx',
    'template.xlsx',
    'bakery_template.xlsx',
    'BakeryOrder.xlsx'
  ];

  const DEFAULT_SETUP = {
    laborRate: 0,
    deliveryRate: 0,
    defaultMargin: 0.35,
    defaultOverhead: 0,
    roundingIncrement: 0,
    templatePath: TEMPLATE_CANDIDATES[0]
  };

  let ingredientsData = [];
  let workbook = null;
  let setupValues = { ...DEFAULT_SETUP };
  let orderLog = [];
  let editingOrderId = null;
  let loadingCount = 0;
  let templateSource = { type: 'none', value: null, path: '' };

  function num(value, fallback = 0) {
    const parsed = Number.parseFloat(value);
    return Number.isFinite(parsed) ? parsed : fallback;
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  function txt(value) {
    return String(value == null ? '' : value).trim();
  }

  function money(value) {
    return num(value, 0).toFixed(2);
  }

  function toDateString(value) {
    const raw = txt(value);
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
    const parsed = new Date(raw);
    if (Number.isNaN(parsed.getTime())) return new Date().toISOString().slice(0, 10);
    return parsed.toISOString().slice(0, 10);
  }

  function showLoader() {
    const overlay = document.getElementById('loadingOverlay');
    if (!overlay) return;
    loadingCount += 1;
    overlay.style.display = 'flex';
  }

  function hideLoader() {
    const overlay = document.getElementById('loadingOverlay');
    if (!overlay) return;
    loadingCount = Math.max(0, loadingCount - 1);
    if (loadingCount === 0) overlay.style.display = 'none';
  }

  function showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
    if (!container || !window.bootstrap || !window.bootstrap.Toast) {
      if (type === 'error') console.error(message);
      else console.log(message);
      return;
    }

    const toastEl = document.createElement('div');
    const color = type === 'error' ? 'danger' : 'success';
    toastEl.className = `toast align-items-center text-white bg-${color} border-0`;
    toastEl.role = 'alert';
    toastEl.ariaLive = 'assertive';
    toastEl.ariaAtomic = 'true';
    toastEl.innerHTML = `\n<div class="d-flex">\n  <div class="toast-body">${message}</div>\n  <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>\n</div>\n`;
    container.appendChild(toastEl);
    const bsToast = new bootstrap.Toast(toastEl);
    bsToast.show();
    toastEl.addEventListener('hidden.bs.toast', () => toastEl.remove());
  }

  function getJson(key, fallback) {
    const raw = localStorage.getItem(key);
    if (!raw) return fallback;
    try {
      return JSON.parse(raw);
    } catch (error) {
      console.warn(`Invalid JSON for key ${key}`, error);
      return fallback;
    }
  }

  function setJson(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (error) {
      console.error(`Failed saving ${key}`, error);
      showToast('Unable to save local data. Browser storage may be full.', 'error');
      return false;
    }
  }

  function getString(key, fallback = '') {
    const value = localStorage.getItem(key);
    return value == null ? fallback : String(value);
  }

  function setString(key, value) {
    try {
      localStorage.setItem(key, String(value));
      return true;
    } catch (error) {
      console.error(`Failed saving ${key}`, error);
      showToast('Unable to save local data. Browser storage may be full.', 'error');
      return false;
    }
  }

  function setText(id, value) {
    const el = document.getElementById(id);
    if (!el) return;
    const nextText = String(value);
    const hasChanged = el.textContent !== nextText;
    el.textContent = nextText;
    if (!hasChanged) return;
    if (window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    el.classList.remove('value-changed');
    void el.offsetWidth;
    el.classList.add('value-changed');
  }

  function setOrderFeedback(message = '', type = 'success') {
    const feedback = document.getElementById('orderEntryFeedback');
    if (!feedback) return;
    feedback.classList.remove('error', 'success');
    if (!message) {
      feedback.textContent = '';
      feedback.style.display = 'none';
      return;
    }
    feedback.textContent = message;
    feedback.classList.add(type === 'error' ? 'error' : 'success');
    feedback.style.display = 'block';
  }

  function clearFieldErrors() {
    const invalid = document.querySelectorAll('.is-invalid-input');
    invalid.forEach((el) => el.classList.remove('is-invalid-input'));
  }

  function markFieldError(id) {
    const field = document.getElementById(id);
    if (field) field.classList.add('is-invalid-input');
  }

  function normalizeIngredient(item) {
    const purchaseQty = Math.max(0, num(item && item.purchaseQty, 0));
    const packageCost = Math.max(0, num(item && item.packageCost, 0));
    const wastePct = clamp(num(item && item.wastePct, 0), 0, 0.99);
    let costPerUnit = 0;
    if (purchaseQty > 0) {
      costPerUnit = packageCost / purchaseQty;
      if (wastePct > 0 && wastePct < 1) costPerUnit /= (1 - wastePct);
    }
    return {
      name: txt(item && item.name),
      purchaseQty,
      baseUnit: txt(item && item.baseUnit),
      packageCost,
      wastePct,
      costPerUnit
    };
  }

  function normalizeIngredients(list) {
    if (!Array.isArray(list)) return [];
    const seen = new Set();
    const result = [];
    list.forEach((item) => {
      const normalized = normalizeIngredient(item || {});
      if (!normalized.name) return;
      const key = normalized.name.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      result.push(normalized);
    });
    return result;
  }

  function normalizeSetup(input, fallback = DEFAULT_SETUP) {
    return {
      laborRate: Math.max(0, num(input && input.laborRate, fallback.laborRate)),
      deliveryRate: Math.max(0, num(input && input.deliveryRate, fallback.deliveryRate)),
      defaultMargin: clamp(num(input && input.defaultMargin, fallback.defaultMargin), 0, 0.99),
      defaultOverhead: Math.max(0, num(input && input.defaultOverhead, fallback.defaultOverhead)),
      roundingIncrement: Math.max(0, num(input && input.roundingIncrement, fallback.roundingIncrement)),
      templatePath: txt(input && input.templatePath) || txt(fallback.templatePath) || TEMPLATE_CANDIDATES[0]
    };
  }

  function normalizeOrderIngredient(item) {
    const name = txt(item && item.name);
    const qtyPerCake = Math.max(0, num(item && item.qtyPerCake, 0));
    const unitCost = Math.max(0, num(item && item.unitCost, 0));
    const lineCost = Math.max(0, num(item && item.lineCost, unitCost * qtyPerCake));
    return { name, qtyPerCake, unitCost, lineCost };
  }

  function normalizeOrder(order) {
    const totalCost = Math.max(0, num(order && order.totalCost, 0));
    const actualPrice = Math.max(0, num(order && order.actualPrice, 0));
    const profit = num(order && order.profit, actualPrice - totalCost);
    const margin = num(order && order.margin, actualPrice > 0 ? (profit / actualPrice) * 100 : 0);
    const ingredients = Array.isArray(order && order.ingredients)
      ? order.ingredients.map((item) => normalizeOrderIngredient(item)).filter((item) => item.name && item.qtyPerCake > 0)
      : [];
    return {
      date: toDateString(order && order.date),
      id: txt(order && order.id),
      customer: txt(order && order.customer),
      product: txt(order && order.product),
      qty: Math.max(0, num(order && order.qty, 0)),
      packagingCost: Math.max(0, num(order && order.packagingCost, 0)),
      laborHours: Math.max(0, num(order && order.laborHours, 0)),
      deliveryKm: Math.max(0, num(order && order.deliveryKm, 0)),
      extraOverhead: Math.max(0, num(order && order.extraOverhead, 0)),
      targetMargin: clamp(num(order && order.targetMargin, DEFAULT_SETUP.defaultMargin), 0, 0.99),
      totalCost,
      suggestedPrice: Math.max(0, num(order && order.suggestedPrice, 0)),
      actualPrice,
      profit,
      margin,
      ingredients
    };
  }

  function normalizeOrderLog(list) {
    if (!Array.isArray(list)) return [];
    const seen = new Set();
    const result = [];
    list.forEach((item) => {
      const normalized = normalizeOrder(item || {});
      if (!normalized.id || seen.has(normalized.id)) return;
      seen.add(normalized.id);
      result.push(normalized);
    });
    return result;
  }

  function persistSetup() {
    setJson(STORAGE.setup, setupValues);
    setString(STORAGE.templatePath, setupValues.templatePath || TEMPLATE_CANDIDATES[0]);
  }

  function persistIngredients() {
    setJson(STORAGE.ingredients, ingredientsData);
  }

  function persistOrderLog() {
    setJson(STORAGE.orderLog, orderLog);
  }

  function wbNumber(sheet, addr, fallback = 0) {
    if (!sheet || !sheet[addr]) return fallback;
    return num(sheet[addr].v, fallback);
  }

  function readSetupDefaults(wb) {
    const sheet = wb && wb.Sheets ? wb.Sheets.Setup : null;
    return {
      laborRate: wbNumber(sheet, 'B6', DEFAULT_SETUP.laborRate),
      deliveryRate: wbNumber(sheet, 'B7', DEFAULT_SETUP.deliveryRate),
      defaultMargin: clamp(wbNumber(sheet, 'B8', DEFAULT_SETUP.defaultMargin), 0, 0.99),
      defaultOverhead: wbNumber(sheet, 'B9', DEFAULT_SETUP.defaultOverhead),
      roundingIncrement: wbNumber(sheet, 'B10', DEFAULT_SETUP.roundingIncrement),
      templatePath: DEFAULT_SETUP.templatePath
    };
  }

  function readIngredientDefaults(wb) {
    const sheet = wb && wb.Sheets ? wb.Sheets.Ingredients : null;
    if (!sheet) return [];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const defaults = [];
    for (let i = 4; i < rows.length; i += 1) {
      const row = rows[i] || [];
      const name = txt(row[0]);
      if (!name) continue;
      defaults.push(normalizeIngredient({
        name,
        purchaseQty: row[3],
        baseUnit: row[4],
        packageCost: row[5],
        wastePct: row[6]
      }));
    }
    return defaults;
  }

  function buildTemplatePaths(preferredPath = '') {
    const paths = [];
    function add(path) {
      const clean = txt(path);
      if (clean && !paths.includes(clean)) paths.push(clean);
    }
    add(preferredPath);
    add(getString(STORAGE.templatePath, ''));
    TEMPLATE_CANDIDATES.forEach(add);
    return paths;
  }

  async function fetchTemplate(path) {
    if (!path) return null;
    try {
      const res = await fetch(path, { cache: 'no-store' });
      if (!res.ok) return null;
      const buffer = await res.arrayBuffer();
      if (!buffer || buffer.byteLength === 0) return null;
      return buffer;
    } catch (error) {
      return null;
    }
  }

  async function resolveTemplate(preferredPath = '') {
    const paths = buildTemplatePaths(preferredPath);
    for (const path of paths) {
      const buffer = await fetchTemplate(path);
      if (!buffer) continue;
      try {
        return {
          workbook: XLSX.read(buffer, { type: 'array', cellFormula: true, cellStyles: true }),
          source: { type: 'array', value: buffer.slice(0), path }
        };
      } catch (error) {
        console.warn(`Workbook parse failed for ${path}`, error);
      }
    }
    if (typeof TEMPLATE_B64 === 'string' && TEMPLATE_B64.trim()) {
      return {
        workbook: XLSX.read(TEMPLATE_B64, { type: 'base64', cellFormula: true, cellStyles: true }),
        source: { type: 'base64', value: TEMPLATE_B64, path: 'embedded' }
      };
    }
    throw new Error('No workbook template found. Add template.xlsx to repository root.');
  }

  function cloneWorkbookForExport() {
    if (templateSource.type === 'array' && templateSource.value) {
      return XLSX.read(templateSource.value.slice(0), { type: 'array', cellFormula: true, cellStyles: true });
    }
    if (templateSource.type === 'base64' && templateSource.value) {
      return XLSX.read(templateSource.value, { type: 'base64', cellFormula: true, cellStyles: true });
    }
    if (workbook) {
      const binary = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
      return XLSX.read(binary, { type: 'binary', cellFormula: true, cellStyles: true });
    }
    throw new Error('Template workbook is not loaded.');
  }

  function decodeBase64ToArrayBuffer(base64) {
    const binary = window.atob(base64);
    const length = binary.length;
    const bytes = new Uint8Array(length);
    for (let i = 0; i < length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes.buffer;
  }

  function binaryStringToArrayBuffer(binary) {
    const length = binary.length;
    const bytes = new Uint8Array(length);
    for (let i = 0; i < length; i += 1) bytes[i] = binary.charCodeAt(i) & 0xFF;
    return bytes.buffer;
  }

  function getTemplateArrayBufferForExport() {
    if (templateSource.type === 'array' && templateSource.value) {
      return templateSource.value.slice(0);
    }
    if (templateSource.type === 'base64' && templateSource.value) {
      return decodeBase64ToArrayBuffer(templateSource.value);
    }
    if (workbook) {
      const binary = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
      return binaryStringToArrayBuffer(binary);
    }
    throw new Error('Template workbook is not loaded.');
  }

  function isOrderLogSheetName(name) {
    return txt(name).toLowerCase().replace(/[\s_-]+/g, '') === 'orderlog';
  }

  function removeOrderLogSheetFromWorkbook(wb) {
    if (!wb || !Array.isArray(wb.SheetNames)) return;
    const toRemove = wb.SheetNames.filter((name) => isOrderLogSheetName(name));
    if (toRemove.length === 0) return;

    toRemove.forEach((name) => {
      if (wb.Sheets && wb.Sheets[name]) delete wb.Sheets[name];
    });
    wb.SheetNames = wb.SheetNames.filter((name) => !isOrderLogSheetName(name));

    if (wb.Workbook && Array.isArray(wb.Workbook.Sheets)) {
      wb.Workbook.Sheets = wb.Workbook.Sheets.filter((sheet) => !isOrderLogSheetName(sheet && sheet.name));
    }
  }

  function removeOrderLogSheetFromPopulateWorkbook(popWorkbook) {
    if (!popWorkbook) return;

    if (typeof popWorkbook.deleteSheet === 'function' && typeof popWorkbook.sheets === 'function') {
      try {
        const names = popWorkbook
          .sheets()
          .map((sheet) => (sheet && typeof sheet.name === 'function' ? sheet.name() : ''))
          .filter((name) => isOrderLogSheetName(name));
        names.forEach((name) => popWorkbook.deleteSheet(name));
        return;
      } catch (error) {
        console.warn('Unable to remove order log sheet using sheets().', error);
      }
    }

    if (typeof popWorkbook.deleteSheet === 'function') {
      ['Order_Log', 'Order Log'].forEach((name) => {
        try {
          popWorkbook.deleteSheet(name);
        } catch (error) {
          // Ignore missing sheet names for compatibility with variant templates.
        }
      });
    }
  }

  async function loadTemplate(options = {}) {
    const preferredPath = txt(options.preferredPath);
    const preserveLocal = options.preserveLocal !== false;
    const notify = options.notify === true;
    showLoader();
    try {
      const resolved = await resolveTemplate(preferredPath);
      const setupDefaults = readSetupDefaults(resolved.workbook);
      const ingredientDefaults = readIngredientDefaults(resolved.workbook);

      const storedSetup = preserveLocal ? getJson(STORAGE.setup, null) : null;
      const setupCandidate = storedSetup && typeof storedSetup === 'object'
        ? { ...setupDefaults, ...storedSetup }
        : setupDefaults;
      setupCandidate.templatePath = resolved.source.path === 'embedded'
        ? (preferredPath || setupCandidate.templatePath || getString(STORAGE.templatePath, TEMPLATE_CANDIDATES[0]))
        : resolved.source.path;
      setupValues = normalizeSetup(setupCandidate, setupDefaults);

      const storedIngredients = preserveLocal ? getJson(STORAGE.ingredients, null) : null;
      ingredientsData = normalizeIngredients(Array.isArray(storedIngredients) ? storedIngredients : ingredientDefaults);
      orderLog = normalizeOrderLog(getJson(STORAGE.orderLog, []));

      workbook = resolved.workbook;
      templateSource = resolved.source;

      persistSetup();
      persistIngredients();
      persistOrderLog();

      renderSetup();
      renderIngredientsList();
      updateAllIngredientSelects();
      renderOrderLog();
      calculateTotals();

      if (notify) {
        if (resolved.source.path === 'embedded') {
          showToast('Template loaded from embedded fallback data.', 'success');
        } else {
          showToast(`Template loaded from repository file ${resolved.source.path}.`, 'success');
        }
      }
      return true;
    } catch (error) {
      console.error(error);
      showToast(`Template load failed: ${error.message}`, 'error');
      return false;
    } finally {
      hideLoader();
    }
  }

  function generateOrderId() {
    const current = Number.parseInt(getString(STORAGE.lastOrderId, '0'), 10);
    const next = Number.isFinite(current) ? current + 1 : 1;
    setString(STORAGE.lastOrderId, next);
    return `ORD-${String(next).padStart(4, '0')}`;
  }

  function setEditingMode(orderId = null) {
    editingOrderId = orderId;
    const push = document.getElementById('pushLogBtn');
    const cancel = document.getElementById('cancelEditBtn');
    if (push) push.textContent = editingOrderId ? 'Update Log' : 'Add to Log';
    if (cancel) cancel.style.display = editingOrderId ? 'inline-block' : 'none';
  }

  function rebuildIngredientSelect(select, current = '') {
    if (!select) return;
    while (select.firstChild) select.removeChild(select.firstChild);
    const first = document.createElement('option');
    first.value = '';
    first.textContent = 'Select';
    select.appendChild(first);

    ingredientsData.forEach((ingredient) => {
      const opt = document.createElement('option');
      opt.value = ingredient.name;
      opt.textContent = ingredient.name;
      select.appendChild(opt);
    });

    if (current && !ingredientsData.some((item) => item.name === current)) {
      const missing = document.createElement('option');
      missing.value = current;
      missing.textContent = `${current} (missing)`;
      select.appendChild(missing);
    }

    select.value = current || '';
  }

  function addIngredientRow(ingredientName = '', qty = '') {
    const tbody = document.getElementById('ingredientsBody');
    if (!tbody) return;

    const row = document.createElement('tr');

    const ingredientTd = document.createElement('td');
    const select = document.createElement('select');
    rebuildIngredientSelect(select, ingredientName);
    ingredientTd.appendChild(select);

    const unitCostTd = document.createElement('td');
    unitCostTd.textContent = '0.00';

    const qtyTd = document.createElement('td');
    const qtyInput = document.createElement('input');
    qtyInput.type = 'number';
    qtyInput.min = '0';
    qtyInput.step = '0.01';
    qtyInput.style.width = '100%';
    qtyInput.style.minWidth = '90px';
    qtyInput.value = qty === '' ? '' : num(qty, 0);
    qtyTd.appendChild(qtyInput);

    const lineCostTd = document.createElement('td');
    lineCostTd.textContent = '0.00';

    const actionTd = document.createElement('td');
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.textContent = 'Remove';
    remove.classList.add('btn-danger');
    remove.addEventListener('click', () => {
      const reducedMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
      if (reducedMotion) {
        row.remove();
        calculateTotals();
        return;
      }
      row.classList.add('row-exit');
      window.setTimeout(() => {
        row.remove();
        calculateTotals();
      }, 180);
    });
    actionTd.appendChild(remove);

    select.addEventListener('change', calculateTotals);
    qtyInput.addEventListener('input', calculateTotals);

    row.appendChild(ingredientTd);
    row.appendChild(unitCostTd);
    row.appendChild(qtyTd);
    row.appendChild(lineCostTd);
    row.appendChild(actionTd);
    tbody.appendChild(row);
    row.classList.add('row-enter');
    window.setTimeout(() => row.classList.remove('row-enter'), 260);

    calculateTotals();
  }

  function updateAllIngredientSelects() {
    const selects = document.querySelectorAll('#ingredientsBody select');
    selects.forEach((select) => {
      const current = select.value;
      rebuildIngredientSelect(select, current);
    });
  }

  function currentOrderIngredients() {
    const rows = document.querySelectorAll('#ingredientsBody tr');
    const orderQty = Math.max(0, num(document.getElementById('quantity') && document.getElementById('quantity').value, 0));
    const result = [];

    rows.forEach((row) => {
      const select = row.children[0] && row.children[0].querySelector('select');
      const qtyInput = row.children[2] && row.children[2].querySelector('input');
      const unitCell = row.children[1];
      const lineCell = row.children[3];

      const name = txt(select && select.value);
      const qtyPerCake = Math.max(0, num(qtyInput && qtyInput.value, 0));
      const unitCost = Math.max(0, num(unitCell && unitCell.textContent, 0));
      const lineCost = Math.max(0, num(lineCell && lineCell.textContent, unitCost * qtyPerCake * orderQty));

      if (name && qtyPerCake > 0) {
        result.push({ name, qtyPerCake, unitCost, lineCost });
      }
    });

    return result;
  }

  function populateOrderIngredients(ingredients) {
    const tbody = document.getElementById('ingredientsBody');
    if (!tbody) return;

    tbody.innerHTML = '';
    if (Array.isArray(ingredients) && ingredients.length > 0) {
      ingredients.forEach((item) => addIngredientRow(item.name, item.qtyPerCake));
    } else {
      addIngredientRow();
    }
  }

  function calculateTotals() {
    const rows = document.querySelectorAll('#ingredientsBody tr');
    const orderQty = Math.max(0, num(document.getElementById('quantity') && document.getElementById('quantity').value, 0));
    let ingredientSubtotal = 0;

    rows.forEach((row) => {
      const select = row.children[0] && row.children[0].querySelector('select');
      const qtyInput = row.children[2] && row.children[2].querySelector('input');
      const unitCell = row.children[1];
      const lineCell = row.children[3];

      const ingredientName = txt(select && select.value);
      const ingredient = ingredientsData.find((item) => item.name === ingredientName);
      const unitCost = ingredient ? num(ingredient.costPerUnit, 0) : 0;
      const qtyPerCake = Math.max(0, num(qtyInput && qtyInput.value, 0));
      const lineCost = unitCost * qtyPerCake * orderQty;

      if (unitCell) unitCell.textContent = money(unitCost);
      if (lineCell) lineCell.textContent = money(lineCost);
      ingredientSubtotal += lineCost;
    });

    const packagingCost = Math.max(0, num(document.getElementById('packagingCost') && document.getElementById('packagingCost').value, 0));
    const laborHours = Math.max(0, num(document.getElementById('laborHours') && document.getElementById('laborHours').value, 0));
    const deliveryKm = Math.max(0, num(document.getElementById('deliveryKm') && document.getElementById('deliveryKm').value, 0));
    const extraOverhead = Math.max(0, num(document.getElementById('extraOverhead') && document.getElementById('extraOverhead').value, 0));
    const targetMargin = clamp(num(document.getElementById('targetMargin') && document.getElementById('targetMargin').value, setupValues.defaultMargin), 0, 0.99);
    const actualPrice = Math.max(0, num(document.getElementById('actualPrice') && document.getElementById('actualPrice').value, 0));

    const laborCost = setupValues.laborRate * laborHours;
    const deliveryCost = setupValues.deliveryRate * deliveryKm;
    const totalCost = ingredientSubtotal + packagingCost + laborCost + deliveryCost + extraOverhead;
    const suggestedPrice = targetMargin >= 1 ? totalCost : totalCost / (1 - targetMargin);

    let roundedQuote = suggestedPrice;
    if (setupValues.roundingIncrement > 0) {
      roundedQuote = Math.ceil(suggestedPrice / setupValues.roundingIncrement) * setupValues.roundingIncrement;
    }

    const profit = actualPrice - totalCost;
    const marginActual = actualPrice > 0 ? (profit / actualPrice) * 100 : 0;

    setText('ingSubtotal', money(ingredientSubtotal));
    setText('summaryPackaging', money(packagingCost));
    setText('laborCost', money(laborCost));
    setText('deliveryCost', money(deliveryCost));
    setText('summaryOverhead', money(extraOverhead));
    setText('totalCost', money(totalCost));
    setText('suggestedPrice', money(suggestedPrice));
    setText('roundedQuote', money(roundedQuote));
    setText('profitActual', money(profit));
    setText('marginActual', `${money(marginActual)}%`);

    setText('barTotalCost', money(totalCost));
    setText('barSuggestedPrice', money(roundedQuote));
    setText('barProfit', money(profit));
    setText('barMargin', `${money(marginActual)}%`);

    updateActionState();
  }

  function updateActionState() {
    const saveBtn = document.getElementById('pushLogBtn');
    if (!saveBtn) return;
    const customer = txt(document.getElementById('customer') && document.getElementById('customer').value);
    const product = txt(document.getElementById('product') && document.getElementById('product').value);
    const qty = Math.max(0, num(document.getElementById('quantity') && document.getElementById('quantity').value, 0));
    const hasIngredients = currentOrderIngredients().length > 0;
    saveBtn.disabled = !(customer && product && qty > 0 && hasIngredients);
  }

  function collectOrder() {
    calculateTotals();
    clearFieldErrors();
    setOrderFeedback('', 'success');

    const idInput = document.getElementById('orderId');
    const id = txt(idInput && idInput.value) || generateOrderId();
    if (idInput && !idInput.value) idInput.value = id;

    const customer = txt(document.getElementById('customer') && document.getElementById('customer').value);
    const product = txt(document.getElementById('product') && document.getElementById('product').value);
    const date = toDateString(document.getElementById('orderDate') && document.getElementById('orderDate').value);
    const qty = Math.max(0, num(document.getElementById('quantity') && document.getElementById('quantity').value, 0));
    const actualPrice = Math.max(0, num(document.getElementById('actualPrice') && document.getElementById('actualPrice').value, 0));
    const packagingCost = Math.max(0, num(document.getElementById('packagingCost') && document.getElementById('packagingCost').value, 0));
    const laborHours = Math.max(0, num(document.getElementById('laborHours') && document.getElementById('laborHours').value, 0));
    const deliveryKm = Math.max(0, num(document.getElementById('deliveryKm') && document.getElementById('deliveryKm').value, 0));
    const extraOverhead = Math.max(0, num(document.getElementById('extraOverhead') && document.getElementById('extraOverhead').value, 0));
    const targetMargin = clamp(num(document.getElementById('targetMargin') && document.getElementById('targetMargin').value, setupValues.defaultMargin), 0, 0.99);
    const ingredients = currentOrderIngredients();
    const errors = [];

    if (!customer) {
      markFieldError('customer');
      errors.push('Customer is required.');
    }
    if (!product) {
      markFieldError('product');
      errors.push('Product/Cake is required.');
    }
    if (qty <= 0) {
      markFieldError('quantity');
      errors.push('Quantity must be greater than 0.');
    }
    if (ingredients.length === 0) {
      errors.push('Add at least one ingredient row with quantity.');
    }
    if (errors.length > 0) {
      setOrderFeedback(errors.join(' '), 'error');
      return { error: errors.join(' ') };
    }

    const totalCost = Math.max(0, num(document.getElementById('totalCost') && document.getElementById('totalCost').textContent, 0));
    const suggestedPrice = Math.max(0, num(document.getElementById('roundedQuote') && document.getElementById('roundedQuote').textContent, 0));
    const profit = num(document.getElementById('profitActual') && document.getElementById('profitActual').textContent, actualPrice - totalCost);
    const margin = num(document.getElementById('marginActual') && document.getElementById('marginActual').textContent, actualPrice > 0 ? (profit / actualPrice) * 100 : 0);

    return {
      order: normalizeOrder({
        date,
        id,
        customer,
        product,
        qty,
        packagingCost,
        laborHours,
        deliveryKm,
        extraOverhead,
        targetMargin,
        totalCost,
        suggestedPrice,
        actualPrice,
        profit,
        margin,
        ingredients
      })
    };
  }

  function clearForm() {
    const values = {
      orderId: generateOrderId(),
      customer: '',
      product: '',
      quantity: 1,
      orderDate: new Date().toISOString().slice(0, 10),
      packagingCost: 0,
      laborHours: 0,
      deliveryKm: 0,
      extraOverhead: setupValues.defaultOverhead,
      targetMargin: setupValues.defaultMargin,
      actualPrice: 0
    };

    Object.keys(values).forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.value = values[id];
    });

    populateOrderIngredients([]);
    setEditingMode(null);
    clearFieldErrors();
    setOrderFeedback('', 'success');
    calculateTotals();
  }

  function startEditOrder(orderId) {
    const order = orderLog.find((item) => item.id === orderId);
    if (!order) {
      showToast('Order not found.', 'error');
      return;
    }

    const values = {
      orderId: order.id,
      customer: order.customer,
      product: order.product,
      orderDate: order.date,
      quantity: order.qty,
      packagingCost: order.packagingCost,
      laborHours: order.laborHours,
      deliveryKm: order.deliveryKm,
      extraOverhead: order.extraOverhead,
      targetMargin: order.targetMargin,
      actualPrice: order.actualPrice
    };

    Object.keys(values).forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.value = values[id];
    });

    populateOrderIngredients(order.ingredients || []);
    setEditingMode(order.id);
    activateTab('orderEntryTab');
    calculateTotals();
    showToast(`Editing order ${order.id}.`, 'success');
  }

  function deleteOrder(orderId) {
    const order = orderLog.find((item) => item.id === orderId);
    if (!order) {
      showToast('Order not found.', 'error');
      return;
    }

    if (!window.confirm(`Delete order ${order.id}?`)) return;
    orderLog = orderLog.filter((item) => item.id !== orderId);
    persistOrderLog();
    renderOrderLog();
    if (editingOrderId === orderId) clearForm();
    showToast(`Order ${order.id} deleted.`, 'success');
  }

  function saveOrderToLog() {
    showLoader();
    try {
      const result = collectOrder();
      if (result.error) {
        showToast(result.error, 'error');
        return;
      }

      const order = result.order;

      if (editingOrderId) {
        const index = orderLog.findIndex((item) => item.id === editingOrderId);
        if (index < 0) {
          showToast('Original order not found for update.', 'error');
          return;
        }
        const duplicate = orderLog.findIndex((item, idx) => item.id === order.id && idx !== index);
        if (duplicate >= 0) {
          showToast(`Order ID ${order.id} already exists.`, 'error');
          return;
        }
        orderLog[index] = order;
        persistOrderLog();
        renderOrderLog();
        clearForm();
        setOrderFeedback('Order updated successfully.', 'success');
        showToast('Order updated successfully.', 'success');
        return;
      }

      if (orderLog.some((item) => item.id === order.id)) {
        showToast(`Order ID ${order.id} already exists.`, 'error');
        return;
      }

      orderLog.push(order);
      persistOrderLog();
      renderOrderLog();
      clearForm();
      setOrderFeedback('Order added successfully.', 'success');
      showToast('Order added successfully.', 'success');
    } finally {
      hideLoader();
    }
  }

  function cancelEdit() {
    if (!editingOrderId) return;
    clearForm();
    setOrderFeedback('Edit canceled.', 'success');
    showToast('Edit canceled.', 'success');
  }

  function clearOrderLog() {
    if (orderLog.length === 0) {
      showToast('Order log is already empty.', 'success');
      return;
    }

    if (!window.confirm('Clear the entire order log? This cannot be undone.')) return;
    orderLog = [];
    persistOrderLog();
    renderOrderLog();
    if (editingOrderId) clearForm();
    showToast('Order log cleared.', 'success');
  }

  function renderOrderLog() {
    const tbody = document.getElementById('orderLogBody');
    if (!tbody) return;

    tbody.innerHTML = '';

    if (orderLog.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = 11;
      td.textContent = 'No orders yet.';
      td.style.textAlign = 'center';
      tr.appendChild(td);
      tbody.appendChild(tr);
      return;
    }

    orderLog.forEach((order) => {
      const tr = document.createElement('tr');
      const fields = [
        order.date,
        order.id,
        order.customer,
        order.product,
        money(order.qty),
        money(order.totalCost),
        money(order.suggestedPrice),
        money(order.actualPrice),
        money(order.profit),
        `${money(order.margin)}%`
      ];

      fields.forEach((value) => {
        const td = document.createElement('td');
        td.textContent = value;
        tr.appendChild(td);
      });

      const actions = document.createElement('td');
      actions.style.whiteSpace = 'nowrap';

      const edit = document.createElement('button');
      edit.type = 'button';
      edit.textContent = 'Edit';
      edit.style.marginRight = '6px';
      edit.addEventListener('click', () => startEditOrder(order.id));

      const del = document.createElement('button');
      del.type = 'button';
      del.textContent = 'Delete';
      del.classList.add('btn-danger');
      del.addEventListener('click', () => deleteOrder(order.id));

      actions.appendChild(edit);
      actions.appendChild(del);
      tr.appendChild(actions);
      tbody.appendChild(tr);
    });
  }

  function addIngredientListRow(item = {}) {
    const tbody = document.getElementById('ingredientsListBody');
    if (!tbody) return;

    const tr = document.createElement('tr');

    const nameTd = document.createElement('td');
    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.value = txt(item.name);
    nameTd.appendChild(nameInput);

    const qtyTd = document.createElement('td');
    const qtyInput = document.createElement('input');
    qtyInput.type = 'number';
    qtyInput.min = '0';
    qtyInput.step = '0.0001';
    qtyInput.value = item.purchaseQty == null ? '' : num(item.purchaseQty, 0);
    qtyTd.appendChild(qtyInput);

    const unitTd = document.createElement('td');
    const unitInput = document.createElement('input');
    unitInput.type = 'text';
    unitInput.value = txt(item.baseUnit);
    unitTd.appendChild(unitInput);

    const costTd = document.createElement('td');
    const costInput = document.createElement('input');
    costInput.type = 'number';
    costInput.min = '0';
    costInput.step = '0.0001';
    costInput.value = item.packageCost == null ? '' : num(item.packageCost, 0);
    costTd.appendChild(costInput);

    const wasteTd = document.createElement('td');
    const wasteInput = document.createElement('input');
    wasteInput.type = 'number';
    wasteInput.min = '0';
    wasteInput.max = '0.99';
    wasteInput.step = '0.0001';
    wasteInput.value = item.wastePct == null ? '' : num(item.wastePct, 0);
    wasteTd.appendChild(wasteInput);

    const actionTd = document.createElement('td');
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.textContent = 'Remove';
    remove.classList.add('btn-danger');
    remove.addEventListener('click', () => tr.remove());
    actionTd.appendChild(remove);

    tr.appendChild(nameTd);
    tr.appendChild(qtyTd);
    tr.appendChild(unitTd);
    tr.appendChild(costTd);
    tr.appendChild(wasteTd);
    tr.appendChild(actionTd);
    tbody.appendChild(tr);
  }

  function renderIngredientsList() {
    const tbody = document.getElementById('ingredientsListBody');
    if (!tbody) return;
    tbody.innerHTML = '';
    ingredientsData.forEach((item) => addIngredientListRow(item));
    if (ingredientsData.length === 0) addIngredientListRow();
  }

  function saveIngredientsList() {
    showLoader();
    try {
      const tbody = document.getElementById('ingredientsListBody');
      if (!tbody) return;

      const rows = Array.from(tbody.querySelectorAll('tr'));
      const next = [];
      const seen = new Set();

      for (const row of rows) {
        const cells = row.querySelectorAll('td');
        if (cells.length < 5) continue;

        const name = txt(cells[0].querySelector('input') && cells[0].querySelector('input').value);
        const purchaseQty = Math.max(0, num(cells[1].querySelector('input') && cells[1].querySelector('input').value, 0));
        const baseUnit = txt(cells[2].querySelector('input') && cells[2].querySelector('input').value);
        const packageCost = Math.max(0, num(cells[3].querySelector('input') && cells[3].querySelector('input').value, 0));
        const waste = num(cells[4].querySelector('input') && cells[4].querySelector('input').value, 0);

        const emptyRow = !name && purchaseQty === 0 && !baseUnit && packageCost === 0 && waste === 0;
        if (emptyRow) continue;

        if (!name) {
          showToast('Ingredient name is required for each non-empty row.', 'error');
          return;
        }
        if (waste < 0 || waste >= 1) {
          showToast(`Waste % for ${name} must be between 0 and 0.99.`, 'error');
          return;
        }
        const key = name.toLowerCase();
        if (seen.has(key)) {
          showToast(`Duplicate ingredient name: ${name}.`, 'error');
          return;
        }
        seen.add(key);

        next.push(normalizeIngredient({ name, purchaseQty, baseUnit, packageCost, wastePct: waste }));
      }

      ingredientsData = normalizeIngredients(next);
      persistIngredients();
      renderIngredientsList();
      updateAllIngredientSelects();
      calculateTotals();
      showToast('Ingredients saved successfully.', 'success');
    } finally {
      hideLoader();
    }
  }

  function renderSetup() {
    const values = {
      laborRate: setupValues.laborRate,
      deliveryRate: setupValues.deliveryRate,
      defaultMargin: setupValues.defaultMargin,
      defaultOverhead: setupValues.defaultOverhead,
      roundingIncrement: setupValues.roundingIncrement,
      templatePath: setupValues.templatePath || TEMPLATE_CANDIDATES[0]
    };

    Object.keys(values).forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.value = values[id];
    });
  }

  function saveSetup() {
    showLoader();
    try {
      const raw = {
        laborRate: document.getElementById('laborRate') && document.getElementById('laborRate').value,
        deliveryRate: document.getElementById('deliveryRate') && document.getElementById('deliveryRate').value,
        defaultMargin: document.getElementById('defaultMargin') && document.getElementById('defaultMargin').value,
        defaultOverhead: document.getElementById('defaultOverhead') && document.getElementById('defaultOverhead').value,
        roundingIncrement: document.getElementById('roundingIncrement') && document.getElementById('roundingIncrement').value,
        templatePath: document.getElementById('templatePath') && document.getElementById('templatePath').value
      };

      const previousMargin = num(raw.defaultMargin, setupValues.defaultMargin);
      setupValues = normalizeSetup(raw, setupValues);
      persistSetup();
      renderSetup();

      const target = document.getElementById('targetMargin');
      if (target) target.value = clamp(num(target.value, setupValues.defaultMargin), 0, 0.99);

      const overhead = document.getElementById('extraOverhead');
      if (overhead && txt(overhead.value) === '') overhead.value = setupValues.defaultOverhead;

      calculateTotals();
      showToast('Setup values saved.', 'success');
      if (previousMargin !== setupValues.defaultMargin) {
        showToast('Default margin was clamped to range 0 to 0.99.', 'error');
      }
    } finally {
      hideLoader();
    }
  }

  async function reloadTemplateFromRepo() {
    const preferredPath = txt(document.getElementById('templatePath') && document.getElementById('templatePath').value);
    if (!window.confirm('Reloading will replace setup and ingredient library from workbook. Continue?')) return;
    const loaded = await loadTemplate({ preferredPath, preserveLocal: false, notify: true });
    if (loaded) clearForm();
  }

  function activateTab(targetId) {
    const tabs = document.querySelectorAll('.tabs .tab');
    tabs.forEach((tab) => tab.classList.toggle('active', tab.getAttribute('data-target') === targetId));

    const panes = document.querySelectorAll('.tab-content');
    panes.forEach((pane) => pane.classList.toggle('active', pane.id === targetId));
  }

  function setupTabs() {
    const tabs = document.querySelectorAll('.tabs .tab');
    tabs.forEach((tab) => {
      tab.addEventListener('click', () => activateTab(tab.getAttribute('data-target')));
    });
  }

  function setCellValue(sheet, address, value, forceType = '', options = {}) {
    if (!sheet) return;
    const clearFormula = options && options.clearFormula === true;
    const cell = sheet[address] || {};
    let type = forceType;
    if (!type) {
      if (value instanceof Date) type = 'd';
      else if (typeof value === 'number') type = 'n';
      else type = 's';
    }
    cell.t = type;
    cell.v = value;
    delete cell.w;
    if (clearFormula && cell.f) delete cell.f;
    sheet[address] = cell;
  }

  function buildDownloadName(orderId, orderDate) {
    if (orderId) return `${orderId}.xlsx`;
    if (orderDate) return `Order_${orderDate}.xlsx`;
    return 'BakeryOrder.xlsx';
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  async function exportToExcelWithPopulate(payload) {
    if (!window.XlsxPopulate || typeof window.XlsxPopulate.fromDataAsync !== 'function') return null;
    const popWorkbook = await window.XlsxPopulate.fromDataAsync(getTemplateArrayBufferForExport());
    const orderSheet = popWorkbook.sheet('Order_Costing');
    if (!orderSheet) throw new Error('Sheet "Order_Costing" is missing in workbook.');

    function write(sheet, address, value) {
      if (!sheet) return;
      sheet.cell(address).value(value);
    }

    write(orderSheet, 'B5', payload.orderId);
    write(orderSheet, 'E5', payload.customer);
    write(orderSheet, 'E6', payload.product);
    write(orderSheet, 'B6', payload.orderDate ? new Date(payload.orderDate) : '');
    write(orderSheet, 'B7', payload.quantity);

    write(orderSheet, 'B10', payload.packagingCost);
    write(orderSheet, 'D10', payload.laborHours);
    write(orderSheet, 'F10', payload.deliveryKm);
    write(orderSheet, 'H10', payload.actualPrice);

    if (!payload.overheadIsDefault) write(orderSheet, 'B11', payload.extraOverhead);
    if (!payload.marginIsDefault) write(orderSheet, 'D11', payload.targetMargin);

    for (let row = 15; row <= 39; row += 1) {
      write(orderSheet, `B${row}`, '');
      write(orderSheet, `D${row}`, 0);
    }

    let writeRow = 15;
    payload.orderIngredients.forEach((item) => {
      if (writeRow > 39) return;
      write(orderSheet, `B${writeRow}`, item.name);
      write(orderSheet, `D${writeRow}`, item.qtyPerCake);
      writeRow += 1;
    });

    const setupSheet = popWorkbook.sheet('Setup');
    if (setupSheet) {
      write(setupSheet, 'B6', payload.setupValues.laborRate);
      write(setupSheet, 'B7', payload.setupValues.deliveryRate);
      write(setupSheet, 'B8', payload.setupValues.defaultMargin);
      write(setupSheet, 'B9', payload.setupValues.defaultOverhead);
      write(setupSheet, 'B10', payload.setupValues.roundingIncrement);
    }

    const ingredientsSheet = popWorkbook.sheet('Ingredients');
    if (ingredientsSheet) {
      let row = 5;
      payload.ingredientsData.forEach((item) => {
        write(ingredientsSheet, `A${row}`, item.name);
        write(ingredientsSheet, `D${row}`, item.purchaseQty);
        write(ingredientsSheet, `E${row}`, item.baseUnit);
        write(ingredientsSheet, `F${row}`, item.packageCost);
        write(ingredientsSheet, `G${row}`, item.wastePct);
        row += 1;
      });
    }

    removeOrderLogSheetFromPopulateWorkbook(popWorkbook);

    return popWorkbook.outputAsync({ type: 'blob' });
  }

  async function exportToExcel() {
    showLoader();
    try {
      calculateTotals();

      function value(id) {
        const el = document.getElementById(id);
        return el ? el.value : '';
      }

      const orderId = txt(value('orderId'));
      const orderDate = txt(value('orderDate'));
      const customer = txt(value('customer'));
      const product = txt(value('product'));
      const quantity = Math.max(0, num(value('quantity'), 0));
      const packagingCost = Math.max(0, num(value('packagingCost'), 0));
      const laborHours = Math.max(0, num(value('laborHours'), 0));
      const deliveryKm = Math.max(0, num(value('deliveryKm'), 0));
      const actualPrice = Math.max(0, num(value('actualPrice'), 0));
      const extraOverhead = Math.max(0, num(value('extraOverhead'), 0));
      const targetMargin = clamp(num(value('targetMargin'), setupValues.defaultMargin), 0, 0.99);
      const orderIngredients = currentOrderIngredients();
      const overheadIsDefault = Math.abs(extraOverhead - setupValues.defaultOverhead) <= 1e-9;
      const marginIsDefault = Math.abs(targetMargin - setupValues.defaultMargin) <= 1e-9;

      const payload = {
        orderId,
        orderDate,
        customer,
        product,
        quantity,
        packagingCost,
        laborHours,
        deliveryKm,
        actualPrice,
        extraOverhead,
        targetMargin,
        overheadIsDefault,
        marginIsDefault,
        setupValues: { ...setupValues },
        orderIngredients,
        ingredientsData: ingredientsData.slice(0)
      };

      const fileName = buildDownloadName(orderId, orderDate);

      if (window.XlsxPopulate && typeof window.XlsxPopulate.fromDataAsync === 'function') {
        try {
          const popBlob = await exportToExcelWithPopulate(payload);
          if (popBlob) {
            downloadBlob(popBlob, fileName);
            showToast('Order exported to Excel successfully.', 'success');
            return;
          }
        } catch (populateError) {
          console.warn('Template-preserving export failed. Falling back to SheetJS export.', populateError);
        }
      }

      const wb = cloneWorkbookForExport();
      const orderSheet = wb.Sheets.Order_Costing;
      if (!orderSheet) throw new Error('Sheet "Order_Costing" is missing in workbook.');

      setCellValue(orderSheet, 'B5', orderId, 's');
      setCellValue(orderSheet, 'E5', customer, 's');
      setCellValue(orderSheet, 'E6', product, 's');
      if (orderDate) {
        setCellValue(orderSheet, 'B6', new Date(orderDate), 'd');
      } else {
        setCellValue(orderSheet, 'B6', '', 's');
      }
      setCellValue(orderSheet, 'B7', quantity, 'n');

      setCellValue(orderSheet, 'B10', packagingCost, 'n');
      setCellValue(orderSheet, 'D10', laborHours, 'n');
      setCellValue(orderSheet, 'F10', deliveryKm, 'n');
      setCellValue(orderSheet, 'H10', actualPrice, 'n');

      if (!overheadIsDefault) {
        setCellValue(orderSheet, 'B11', extraOverhead, 'n', { clearFormula: true });
      }

      if (!marginIsDefault) {
        setCellValue(orderSheet, 'D11', targetMargin, 'n', { clearFormula: true });
      }

      for (let row = 15; row <= 39; row += 1) {
        setCellValue(orderSheet, `B${row}`, '', 's');
        setCellValue(orderSheet, `D${row}`, 0, 'n');
      }

      let writeRow = 15;
      orderIngredients.forEach((item) => {
        if (writeRow > 39) return;
        setCellValue(orderSheet, `B${writeRow}`, item.name, 's');
        setCellValue(orderSheet, `D${writeRow}`, item.qtyPerCake, 'n');
        writeRow += 1;
      });

      const setupSheet = wb.Sheets.Setup;
      if (setupSheet) {
        setCellValue(setupSheet, 'B6', setupValues.laborRate, 'n');
        setCellValue(setupSheet, 'B7', setupValues.deliveryRate, 'n');
        setCellValue(setupSheet, 'B8', setupValues.defaultMargin, 'n');
        setCellValue(setupSheet, 'B9', setupValues.defaultOverhead, 'n');
        setCellValue(setupSheet, 'B10', setupValues.roundingIncrement, 'n');
      }

      const ingredientsSheet = wb.Sheets.Ingredients;
      if (ingredientsSheet) {
        let row = 5;
        ingredientsData.forEach((item) => {
          setCellValue(ingredientsSheet, `A${row}`, item.name, 's');
          setCellValue(ingredientsSheet, `D${row}`, item.purchaseQty, 'n');
          setCellValue(ingredientsSheet, `E${row}`, item.baseUnit, 's');
          setCellValue(ingredientsSheet, `F${row}`, item.packageCost, 'n');
          setCellValue(ingredientsSheet, `G${row}`, item.wastePct, 'n');
          row += 1;
        });
      }

      removeOrderLogSheetFromWorkbook(wb);

      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      downloadBlob(blob, fileName);
      showToast('Order exported to Excel successfully.', 'success');
    } catch (error) {
      console.error(error);
      showToast(`Export failed: ${error.message}`, 'error');
    } finally {
      hideLoader();
    }
  }

  function bindRealtimeCalculation() {
    const ids = ['quantity', 'packagingCost', 'laborHours', 'deliveryKm', 'extraOverhead', 'targetMargin', 'actualPrice'];
    ids.forEach((id) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener('input', () => {
        el.classList.remove('is-invalid-input');
        calculateTotals();
      });
      el.addEventListener('change', () => {
        el.classList.remove('is-invalid-input');
        calculateTotals();
      });
    });

    ['customer', 'product'].forEach((id) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener('input', () => {
        el.classList.remove('is-invalid-input');
        updateActionState();
      });
    });
  }

  function bindActions() {
    const addIngredientBtn = document.getElementById('addIngredientBtn');
    const calculateBtn = document.getElementById('calculateBtn');
    const exportBtn = document.getElementById('exportBtn');
    const pushLogBtn = document.getElementById('pushLogBtn');
    const cancelEditBtn = document.getElementById('cancelEditBtn');
    const clearBtn = document.getElementById('clearBtn');
    const addIngredientListBtn = document.getElementById('addIngredientListBtn');
    const saveIngredientsBtn = document.getElementById('saveIngredientsBtn');
    const saveSetupBtn = document.getElementById('saveSetupBtn');
    const reloadTemplateBtn = document.getElementById('reloadTemplateBtn');
    const clearLogBtn = document.getElementById('clearLogBtn');

    if (addIngredientBtn) addIngredientBtn.addEventListener('click', () => addIngredientRow());
    if (calculateBtn) calculateBtn.addEventListener('click', calculateTotals);
    if (exportBtn) exportBtn.addEventListener('click', exportToExcel);
    if (pushLogBtn) pushLogBtn.addEventListener('click', saveOrderToLog);
    if (cancelEditBtn) cancelEditBtn.addEventListener('click', cancelEdit);
    if (clearBtn) clearBtn.addEventListener('click', clearForm);
    if (addIngredientListBtn) addIngredientListBtn.addEventListener('click', () => addIngredientListRow());
    if (saveIngredientsBtn) saveIngredientsBtn.addEventListener('click', saveIngredientsList);
    if (saveSetupBtn) saveSetupBtn.addEventListener('click', saveSetup);
    if (reloadTemplateBtn) reloadTemplateBtn.addEventListener('click', reloadTemplateFromRepo);
    if (clearLogBtn) clearLogBtn.addEventListener('click', clearOrderLog);
  }

  async function init() {
    bindActions();
    bindRealtimeCalculation();
    setupTabs();

    const loaded = await loadTemplate({ preserveLocal: true, notify: false });
    if (!loaded) {
      setupValues = { ...DEFAULT_SETUP };
      ingredientsData = [];
      orderLog = normalizeOrderLog(getJson(STORAGE.orderLog, []));
      renderSetup();
      renderIngredientsList();
      renderOrderLog();
    }

    clearForm();
  }

  document.addEventListener('DOMContentLoaded', () => {
    init();
  });
})();
