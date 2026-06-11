// src/search.js
import cfg from '../config.js';

export async function runMemberSearch(page, countryName, categoryName) {
  console.log(`[Search] Country="${countryName}"  Category="${categoryName}"`);

  // ── Always do a hard reload to reset page state between categories ─────
  // Critical fix: after first category, page retains previous filter state.
  // Hard reload guarantees a clean slate every time.
  await page.goto(cfg.SEARCH_URL, {
    waitUntil: 'networkidle',
    timeout:   cfg.NAV_TIMEOUT,
  });
  await page.waitForTimeout(3000);

  // ── Wait for React to fully mount ─────────────────────────────────────
  // In headless mode React takes longer. Wait for #root to have children.
  await page.waitForFunction(() => {
    const root = document.querySelector('#root');
    return root && root.children.length > 0 && document.querySelectorAll('button').length > 0;
  }, { timeout: 30_000 }).catch(() => console.warn('[Search] React mount check timed out'));

  await page.waitForTimeout(2000);

  // ── DEBUG: log what buttons exist ─────────────────────────────────────
  const pageButtons = await page.evaluate(() =>
    [...document.querySelectorAll('button')].map(b => b.textContent.trim()).filter(Boolean)
  );
  console.log('[Search] Buttons on page:', pageButtons.slice(0, 10));

  // ── Click Filter button with multiple selector strategies ─────────────
  const filterClicked = await clickFilterButton(page);
  if (!filterClicked) {
    throw new Error('Could not click Filter button — check debug output above');
  }
  await page.waitForTimeout(1500);

  // ── Country ────────────────────────────────────────────────────────────
  console.log(`[Search] Setting Country = ${countryName}`);
  await setComboBox(page, 'Select Country', countryName.trim().substring(0, 3).toLowerCase());
  await page.waitForTimeout(800);
  await clickOption(page, countryName.trim());
  await page.waitForTimeout(800);

  // ── Category ───────────────────────────────────────────────────────────
  console.log(`[Search] Setting Category = ${categoryName}`);
  const searchTerm = extractSearchTerm(categoryName);
  console.log(`[Search] Typing search term: "${searchTerm}"`);
  await setComboBox(page, 'Search Category', searchTerm);
  await page.waitForTimeout(1200);

  const catClicked = await clickCategoryOption(page, categoryName);
  if (!catClicked) {
    console.warn(`[Search] Category option not found — pressing ArrowDown+Enter`);
    await page.keyboard.press('ArrowDown');
    await page.waitForTimeout(300);
    await page.keyboard.press('Enter');
  }
  await page.waitForTimeout(800);

  // ── Search Members ─────────────────────────────────────────────────────
  await clickSearchMembersButton(page);
  console.log('[Search] Waiting for results...');

  try {
    await page.waitForSelector('text=Search Results', { timeout: 30_000 });
  } catch {
    console.warn('[Search] "Search Results" text not found — continuing anyway');
  }
  await page.waitForTimeout(cfg.WAIT_AFTER_SEARCH || 5000);

  // ── Scroll results container ───────────────────────────────────────────
  await scrollResultsContainer(page);

  // ── Extract members ────────────────────────────────────────────────────
  const members = await extractMembersTable(page);
  console.log(`[Search] Extracted ${members.length} members.`);
  return members;
}

// ─────────────────────────────────────────────────────────────────────────────
//  Click the Filter button — tries 6 different selector strategies
//  because in headless mode getByRole can fail if button text is in an SVG
// ─────────────────────────────────────────────────────────────────────────────
async function clickFilterButton(page) {
  const strategies = [
    // Strategy 1: exact role (works in headed mode)
    async () => {
      const btn = page.getByRole('button', { name: 'Filter' });
      await btn.waitFor({ state: 'visible', timeout: 8000 });
      await btn.click();
      return true;
    },
    // Strategy 2: text content match
    async () => {
      const btn = page.locator('button:has-text("Filter")').first();
      await btn.waitFor({ state: 'visible', timeout: 5000 });
      await btn.click();
      return true;
    },
    // Strategy 3: aria-label
    async () => {
      const btn = page.locator('button[aria-label*="Filter" i]').first();
      await btn.waitFor({ state: 'visible', timeout: 5000 });
      await btn.click();
      return true;
    },
    // Strategy 4: find by position — Filter is always the first major button
    // on the search page (before "Search Members")
    async () => {
      const buttons = page.locator('button').filter({ hasNotText: /sign|login|back/i });
      const count   = await buttons.count();
      for (let i = 0; i < Math.min(count, 5); i++) {
        const text = await buttons.nth(i).textContent().catch(() => '');
        if (text.toLowerCase().includes('filter')) {
          await buttons.nth(i).click();
          return true;
        }
      }
      return false;
    },
    // Strategy 5: click by coordinates — Filter button is typically
    // in the top-right area of the search bar
    async () => {
      const viewport = page.viewportSize() || { width: 1280, height: 720 };
      // Filter button is roughly at 75% width, ~215px from top
      await page.mouse.click(viewport.width * 0.75, 215);
      await page.waitForTimeout(500);
      // Check if filter panel appeared (combobox for Country)
      const comboVisible = await page.locator('[role="combobox"]').first()
        .isVisible({ timeout: 2000 }).catch(() => false);
      return comboVisible;
    },
    // Strategy 6: JavaScript click on any button containing "filter" text
    async () => {
      const clicked = await page.evaluate(() => {
        const btns = [...document.querySelectorAll('button')];
        const filterBtn = btns.find(b =>
          b.textContent.toLowerCase().includes('filter') ||
          b.getAttribute('aria-label')?.toLowerCase().includes('filter')
        );
        if (filterBtn) { filterBtn.click(); return true; }
        return false;
      });
      if (clicked) {
        await page.waitForTimeout(1000);
        return true;
      }
      return false;
    },
  ];

  for (let i = 0; i < strategies.length; i++) {
    try {
      const success = await strategies[i]();
      if (success) {
        console.log(`[Search] Filter clicked via strategy ${i + 1}`);
        return true;
      }
    } catch (err) {
      console.warn(`[Search] Filter strategy ${i + 1} failed: ${err.message.split('\n')[0]}`);
      await page.waitForTimeout(1500);

      // After strategy 3 fails, do a page reload before trying position-based
      if (i === 2) {
        console.log('[Search] Reloading page before trying next strategies...');
        await page.goto(cfg.SEARCH_URL, { waitUntil: 'networkidle', timeout: cfg.NAV_TIMEOUT });
        await page.waitForTimeout(3000);
        await page.waitForFunction(
          () => document.querySelectorAll('button').length > 0,
          { timeout: 15_000 }
        ).catch(() => {});
      }
    }
  }

  return false;
}

// ─────────────────────────────────────────────────────────────────────────────
//  Set a MUI combobox value
// ─────────────────────────────────────────────────────────────────────────────
async function setComboBox(page, labelText, fillValue) {
  const selectors = [
    `[role="combobox"][aria-label*="${labelText}" i]`,
    `input[aria-label*="${labelText}" i]`,
    `input[placeholder*="${labelText}" i]`,
  ];

  for (const sel of selectors) {
    try {
      const el = page.locator(sel).first();
      const vis = await el.isVisible({ timeout: 4000 }).catch(() => false);
      if (vis) {
        await el.click();
        await el.fill(fillValue);
        return;
      }
    } catch { /* try next */ }
  }

  // Fallback: find by role combobox and match placeholder/label
  try {
    const combo = page.getByRole('combobox', { name: new RegExp(labelText, 'i') }).first();
    await combo.click();
    await combo.fill(fillValue);
  } catch (err) {
    throw new Error(`[Search] Could not set combobox "${labelText}": ${err.message}`);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  Click a dropdown option by text
// ─────────────────────────────────────────────────────────────────────────────
async function clickOption(page, text) {
  try {
    await page.getByRole('option', { name: text }).first().click({ timeout: 8000 });
    return;
  } catch { /* try next */ }

  try {
    await page.locator(`[role="option"]:has-text("${text}")`).first().click({ timeout: 5000 });
    return;
  } catch { /* try next */ }

  // Fallback: keyboard select
  await page.keyboard.press('ArrowDown');
  await page.waitForTimeout(200);
  await page.keyboard.press('Enter');
}

// ─────────────────────────────────────────────────────────────────────────────
//  Click Search Members button
// ─────────────────────────────────────────────────────────────────────────────
async function clickSearchMembersButton(page) {
  const strategies = [
    () => page.getByRole('button', { name: 'Search Members' }).click({ timeout: 8000 }),
    () => page.locator('button:has-text("Search Members")').first().click({ timeout: 5000 }),
    () => page.evaluate(() => {
      const btn = [...document.querySelectorAll('button')]
        .find(b => b.textContent.includes('Search Members'));
      if (btn) { btn.click(); return true; }
      return false;
    }),
  ];

  for (const strategy of strategies) {
    try {
      await strategy();
      return;
    } catch { /* try next */ }
  }
  throw new Error('Could not click Search Members button');
}

// ─────────────────────────────────────────────────────────────────────────────
//  Scroll the results container
// ─────────────────────────────────────────────────────────────────────────────
async function scrollResultsContainer(page) {
  console.log('[Search] Scrolling results container...');

  const scrolled = await page.evaluate(async (maxScrolls) => {
    const allEls = [...document.querySelectorAll('*')];
    let container = null;

    for (const el of allEls) {
      const style    = window.getComputedStyle(el);
      const overflow = style.overflowY;
      if (
        (overflow === 'auto' || overflow === 'scroll') &&
        el.scrollHeight > el.clientHeight + 50
      ) {
        const anchors = el.querySelectorAll('a');
        if (anchors.length >= 3) {
          container = el; // keep going — we want the innermost one
        }
      }
    }

    if (!container) return false;

    let lastScrollTop = -1;
    let stable        = 0;

    for (let i = 0; i < maxScrolls; i++) {
      container.scrollTop += 800;
      await new Promise(r => setTimeout(r, 200));
      if (container.scrollTop === lastScrollTop) {
        stable++;
        if (stable >= 3) break;
      } else {
        stable = 0;
      }
      lastScrollTop = container.scrollTop;
    }
    return true;
  }, cfg.PAGE_DOWN_COUNT);

  if (!scrolled) {
    console.log('[Search] Container scroll failed — mouse wheel fallback...');
    try {
      const box = await page.locator('text=Search Results').boundingBox();
      if (box) {
        await page.mouse.move(box.x + box.width / 2, box.y + 200);
        for (let i = 0; i < cfg.PAGE_DOWN_COUNT; i++) {
          await page.mouse.wheel(0, 600);
          await page.waitForTimeout(150);
        }
      }
    } catch {
      for (let i = 0; i < cfg.PAGE_DOWN_COUNT; i++) {
        await page.keyboard.press('PageDown');
        await page.waitForTimeout(150);
      }
    }
  }

  await page.waitForTimeout(1500);
  const count = await page.evaluate(() =>
    document.querySelectorAll('tr td:first-child a, [role="row"] a').length
  );
  console.log(`[Search] After scroll: ~${count} member links visible`);
}

// ─────────────────────────────────────────────────────────────────────────────
//  Extract member rows from the results table
// ─────────────────────────────────────────────────────────────────────────────
async function extractMembersTable(page) {
  const BASE = 'https://www.bniconnectglobal.com';

  return await page.evaluate((base) => {
    const results = [];
    const seen    = new Set();
    const NAV_WORDS = ['Help', 'Home', 'Search', 'Filter', 'Back', 'Sign In',
                       'Connect', 'My BNI', 'Dashboard', 'Settings', 'Support',
                       'zendesk', 'BNI Connect'];

    // Find Search Results section to scope extraction
    let searchSection = null;
    for (const el of document.querySelectorAll('h1,h2,h3,h4,h5,h6,p,div,span')) {
      if (el.childElementCount === 0 && el.textContent.trim() === 'Search Results') {
        let parent = el.parentElement;
        for (let d = 0; d < 5; d++) {
          if (parent && parent.querySelectorAll('a').length >= 3) {
            searchSection = parent;
            break;
          }
          parent = parent?.parentElement;
        }
        if (searchSection) break;
      }
    }
    const scope = searchSection || document.body;

    // Strategy 1: HTML table rows
    scope.querySelectorAll('tr').forEach(row => {
      const cells    = row.querySelectorAll('td');
      if (cells.length < 4) return;
      const nameLink = cells[0].querySelector('a');
      if (!nameLink) return;
      const name = nameLink.innerText.trim();
      if (!name || seen.has(name)) return;
      if (NAV_WORDS.some(w => name.includes(w))) return;
      seen.add(name);
      const href = nameLink.getAttribute('href') || '';
      results.push({
        Name:                   name,
        Chapter:                cells[1]?.innerText.trim() || '',
        Company:                cells[2]?.innerText.trim() || '',
        City:                   cells[3]?.innerText.trim() || '',
        IndustryClassification: cells[4]?.innerText.trim() || '',
        profileUrl: href ? (href.startsWith('http') ? href : base + href) : '',
      });
    });

    if (results.length > 0) return results;

    // Strategy 2: React div-based rows
    scope.querySelectorAll('a').forEach(anchor => {
      const name = anchor.innerText.trim();
      if (!name || name.length > 80 || seen.has(name)) return;
      if (NAV_WORDS.some(w => name.toLowerCase().includes(w.toLowerCase()))) return;
      if (!/[a-zA-Z]/.test(name)) return;

      let rowContainer = anchor.parentElement;
      for (let d = 0; d < 6; d++) {
        if (!rowContainer) break;
        const texts = [...rowContainer.querySelectorAll('p, span, td')]
          .map(el => el.innerText.trim())
          .filter(t => t && t !== name && t.length > 1);
        if (texts.length >= 3) break;
        rowContainer = rowContainer.parentElement;
      }
      if (!rowContainer) return;

      const cols = [...rowContainer.querySelectorAll('p, td')]
        .map(el => el.innerText.trim())
        .filter(t => t && t !== name && !/^[A-Z]{1,3}$/.test(t) && t !== '+');

      const href = anchor.getAttribute('href') || '';
      seen.add(name);
      results.push({
        Name:                   name,
        Chapter:                cols[0] || '',
        Company:                cols[1] || '',
        City:                   cols[2] || '',
        IndustryClassification: cols[3] || '',
        profileUrl: href ? (href.startsWith('http') ? href : base + href) : '',
      });
    });

    return results;
  }, BASE);
}

// ─────────────────────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────────────────────
function extractSearchTerm(categoryName) {
  const match = categoryName.match(/\(([^)]+)\)/);
  if (match) return match[1].split('/')[0].split(' ')[0];
  return categoryName.split(' ')[0];
}

async function clickCategoryOption(page, categoryName) {
  const candidates = [
    categoryName,
    categoryName.replace(/\(([^)]+)\)/, '> $1'),
    categoryName.replace(/.*\(([^)]+)\).*/, '$1'),
    categoryName.replace(/\s*\([^)]*\)/, '').trim(),
  ];

  for (const text of candidates) {
    try {
      const opt   = page.getByRole('option', { name: text.trim() });
      const count = await opt.count();
      if (count > 0) {
        await opt.first().click({ timeout: 3000 });
        console.log(`[Search] Selected: "${text.trim()}"`);
        return true;
      }
    } catch { /* try next */ }

    try {
      const opt = page.locator(`[role="option"]:has-text("${text.trim()}")`).first();
      if (await opt.isVisible({ timeout: 2000 }).catch(() => false)) {
        await opt.click();
        console.log(`[Search] Selected (locator): "${text.trim()}"`);
        return true;
      }
    } catch { /* try next */ }
  }
  return false;
}