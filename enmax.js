(async () => {
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const until = async (fn, {timeout=12000, interval=120} = {}) => {
    const t0 = performance.now();
    let lastErr;
    while (performance.now() - t0 < timeout) {
      try {
        const v = await fn();
        if (v) return v;
      } catch (e) { lastErr = e; }
      await sleep(interval);
    }
    if (lastErr) throw lastErr;
    throw new Error('Timed out');
  };
  const visible = (el) => !!(el && el.isConnected && el.offsetParent !== null);
  const checksum = (el) => {
    if (!el) return 0;
    const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null);
    let h = 0, n;
    while ((n = w.nextNode())) {
      const t = n.nodeValue?.trim();
      if (!t) continue;
      for (let i=0;i<t.length;i++) { h=((h<<5)-h)+t.charCodeAt(i); h|=0; }
    }
    return h>>>0;
  };
  const monthMap = Object.fromEntries("January February March April May June July August September October November December".split(" ").map((m,i)=>[m,i+1]));
  const toISODate = (pretty) => {
    const m = pretty?.match(/([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})/);
    if (!m) return null;
    const [, mon, d, y] = m;
    const mm = String(monthMap[mon]||'').padStart(2,'0');
    const dd = String(d).padStart(2,'0');
    return `${y}-${mm}-${dd}`;
  };

  /** ---------- TARGET THE HEADER ACCOUNT DROPDOWN ---------- */
  const headerDropdown = () => document.querySelector('[data-component="HeaderMyAccountDropdown"]');
  const headerSelect = () => headerDropdown()?.querySelector('.header-myaccount-dropdown_dropdown__select__JMWP2');
  const headerList   = () => headerDropdown()?.querySelector('.header-myaccount-dropdown_dropdown__option__list__YDzIh');
  const headerOpts   = () => [...headerDropdown()?.querySelectorAll('.header-myaccount-dropdown_dropdown__option__Pinvp') ?? []];
  const headerCurrentNumberEl = () => headerSelect()?.querySelector('.body-regular-bold');

  const openHeaderMenu = async () => {
    const list = headerList();
    if (!visible(list) || (list && list.classList.contains('hidden'))) {
      headerSelect()?.click();
      await sleep(50);
    }
  };
  const closeHeaderMenu = async () => {
    const list = headerList();
    if (visible(list) && list && !list.classList.contains('hidden')) {
      headerSelect()?.click();
      await sleep(50);
    }
  };

  const readHeaderOptions = async () => {
    await openHeaderMenu();
    const items = headerOpts();
    return items.map((el, idx) => {
      const num = el.querySelector('.body-regular-bold')?.textContent.trim() ?? '';
      const note = el.querySelector('.body-small')?.textContent.trim() ?? '';
      return {
        idx,
        accountNumber: num,
        label: note ? `${num} — ${note}` : num,
        select: async () => { await openHeaderMenu(); el.click(); }
      };
    });
  };

  /** ---------- BILL HISTORY HELPERS ---------- */
  const billRoot = () => document.querySelector('[data-component="MyBillHistory"]');
  const pickLatestDownloadButton = () => {
    const root = billRoot();
    const all = [...root?.querySelectorAll('button[aria-label="Download"]') ?? []].filter(visible);
    if (!all.length) return null;
    return all[0]; // newest appears first
  };
  const extractBillMeta = (btn) => {
    let scope = btn.closest('.flex') || btn.parentElement;
    for (let i=0;i<6 && scope && !scope.querySelector('p'); i++) scope = scope.parentElement || scope;
    const texts = scope ? [...scope.querySelectorAll('p,span')].map(x => x.textContent.trim()) : [];
    const dates = texts.filter(t => /\b(January|February|March|April|May|June|July|August|September|October|November|December)\b/.test(t) && /\d{4}/.test(t));
    const billDate = dates[0] || null;
    const dueDate  = dates[1] || null;
    const amount   = (texts.find(t => /[\d]+\.\d{2}/.test(t)) || null);
    const billISO  = billDate ? toISODate(billDate) : null;
    return { billDate, dueDate, amount, billISO, suggestedFilename: billISO ? `bill_${billISO}${amount?`_${amount.replace(/[^\d.]/g,'')}`:''}.pdf` : 'bill_latest.pdf' };
  };

  /** ---------- NEW: VIEW-MY-BILL (click once if present) ---------- */
  const getViewMyBillEl = () => {
    // Prefer aria-label; fall back to text match on <a>/<button>
    const byAria = document.querySelector('a[aria-label="View my bill"], button[aria-label="View my bill"]');
    if (visible(byAria)) return byAria;

    const candidates = [...document.querySelectorAll('a,button')];
    for (const el of candidates) {
      if (!visible(el)) continue;
      const txt = el.textContent?.trim() || '';
      if (/^view\s+my\s+bill$/i.test(txt)) return el;
    }
    return null;
  };

  const maybeOpenViewMyBill = async () => {
    // Global guard so we only ever click once
    if (window.__billBotViewMyBillOpened) return true;

    const el = getViewMyBillEl();
    if (!el) return false;

    // Avoid dropdown overlaying / intercepting clicks
    try { await closeHeaderMenu(); } catch {}

    console.log('[bill-bot] Opening "View my bill" window once.');
    el.click();
    window.__billBotViewMyBillOpened = true;

    // Give the panel a moment to mount (non-blocking downstream)
    await sleep(150);
    return true;
  };

  /** ---------- MAIN ---------- */
  const run = async () => {
    const dropdown = await until(() => headerDropdown(), { timeout: 8000 });
    const currentAccount = headerCurrentNumberEl()?.textContent.trim();
    if (!currentAccount) throw new Error('Could not read current account number from header.');

    const options = await readHeaderOptions();
    const ordered = [
      ...options.filter(o => o.accountNumber === currentAccount),
      ...options.filter(o => o.accountNumber !== currentAccount)
    ];

    console.log(`[bill-bot] Header account dropdown detected with ${options.length} accounts.`);

    // Try to open "View my bill" before iterating (if it exists)
    await maybeOpenViewMyBill();

    const results = [];

    for (let i=0;i<ordered.length;i++) {
      const opt = ordered[i];

      const beforeHeaderTxt = headerCurrentNumberEl()?.textContent.trim();
      const beforeBillSig   = checksum(billRoot());

      if (opt.accountNumber !== beforeHeaderTxt) {
        console.log(`[bill-bot] Switching to account ${opt.label}`);
        await opt.select();

        try {
          await until(() => {
            const headerChanged = headerCurrentNumberEl()?.textContent.trim() === opt.accountNumber;
            const billChanged   = checksum(billRoot()) !== beforeBillSig;
            return headerChanged || billChanged;
          }, { timeout: 15000, interval: 150 });
        } catch {
          console.warn(`[bill-bot] UI did not confirm switch to ${opt.accountNumber} within timeout; continuing.`);
        }
      } else {
        await closeHeaderMenu();
      }

      // Try again in case the button appears only after some route/state update.
      await maybeOpenViewMyBill();

      // Now grab the newest bill for THIS account
      const btn = await until(() => pickLatestDownloadButton(), { timeout: 12000, interval: 150 });
      const meta = extractBillMeta(btn);
      const detail = { propertyIndex: i, propertyLabel: opt.label, accountNumber: opt.accountNumber, ...meta };

      if (meta.billISO) btn.setAttribute('data-testid', `download-bill-${opt.accountNumber}-${meta.billISO}`);

      console.log('[bill-bot] Downloading most recent bill:', detail);
      window.dispatchEvent(new CustomEvent('bill:download', { detail, bubbles: true }));
      btn.click();

      results.push(detail);
      await sleep(600);
    }

    console.table(results.map(r => ({
      account: r.accountNumber,
      billDate: r.billDate,
      amount: r.amount,
      file: r.suggestedFilename
    })));

    console.log('[bill-bot] Finished iterating all accounts in the header dropdown.');
  };

  try { await run(); } catch (e) { console.error('[bill-bot] Failed:', e); }
})();
