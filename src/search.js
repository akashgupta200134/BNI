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
    // CRITICAL: never blind-select with ArrowDown+Enter. If we can't match
    // the category, that previously caused the search to run with NO
    // category filter at all (or a wrong one), polluting the sheet with
    // unrelated members. Throw instead — bot.js already handles search
    // failures per-category and skips that one without breaking the run.
    throw new Error(
      `[Search] Could not find/select category option for "${categoryName}" — ` +
      `aborting this category rather than risk selecting the wrong filter.`
    );
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
//  Scroll the results container — waits for BNI to load each batch before
//  continuing, and only stops when NO new rows appear after multiple retries
// ─────────────────────────────────────────────────────────────────────────────
async function scrollResultsContainer(page) {
  console.log('[Search] Scrolling results container...');

  // ── Step 1: Find the scrollable container (Node context) ─────────────────
  // We do the container-finding outside evaluate so we can use
  // page.waitForFunction between scrolls (evaluate can't await across ticks)
  const containerHandle = await page.evaluateHandle(() => {
    const allEls = [...document.querySelectorAll('*')];
    let container = null;
    for (const el of allEls) {
      const overflow = window.getComputedStyle(el).overflowY;
      if (
        (overflow === 'auto' || overflow === 'scroll') &&
        el.scrollHeight > el.clientHeight + 50 &&
        el.querySelectorAll('a').length >= 3
      ) {
        container = el; // keep iterating — we want the innermost matching el
      }
    }
    return container;
  });

  const isValidContainer = await page.evaluate(el => el !== null, containerHandle);

  if (!isValidContainer) {
    // ── Fallback: mouse wheel scroll ────────────────────────────────────────
    console.log('[Search] No scrollable container found — trying mouse wheel fallback...');
    try {
      const box = await page.locator('text=Search Results').boundingBox();
      if (box) {
        await page.mouse.move(box.x + box.width / 2, box.y + 200);
      } else {
        await page.mouse.move(640, 400);
      }

      let prevCount  = 0;
      let stableRounds = 0;

      while (stableRounds < 6) {
        await page.mouse.wheel(0, 600);
        // Wait up to 5s for new rows to appear
        await page.waitForFunction(
          (prev) => document.querySelectorAll('tr td:first-child a, [role="row"] a').length > prev,
          prevCount,
          { timeout: 5000 }
        ).catch(() => {}); // timeout = no new rows loaded, that's fine

        const curCount = await page.evaluate(() =>
          document.querySelectorAll('tr td:first-child a, [role="row"] a').length
        );
        console.log(`[Scroll-Wheel] Links visible: ${curCount}`);

        if (curCount > prevCount) {
          prevCount    = curCount;
          stableRounds = 0;
        } else {
          stableRounds++;
        }
      }
    } catch (err) {
      console.warn('[Search] Mouse wheel fallback failed:', err.message);
    }

    await page.waitForTimeout(2000);
    const finalCount = await page.evaluate(() =>
      document.querySelectorAll('tr td:first-child a, [role="row"] a').length
    );
    console.log(`[Search] After scroll: ~${finalCount} member links visible`);
    return;
  }

  // ── Step 2: Scroll incrementally, waiting for BNI to load each batch ─────
  console.log('[Search] Found scrollable container — starting incremental scroll...');

  let prevCount    = 0;
  let stableRounds = 0;
  let totalScrolls = 0;

  // Count links using both table and div-row selectors
  const countLinks = () => page.evaluate(el => el.querySelectorAll('a').length, containerHandle);

  while (stableRounds < 6) {   // 6 rounds × up to 8s each = up to 48s of patience
    // Scroll down by one viewport-ish chunk
    await page.evaluate(el => { el.scrollTop += 600; }, containerHandle);
    totalScrolls++;

    // Wait for BNI to load new rows — up to 8 seconds per scroll step.
    // If new links appear before timeout, we continue immediately.
    // If nothing loads in 8s, we treat it as a stable (no-new-data) round.
    try {
      await page.waitForFunction(
        ({ el, prev }) => el.querySelectorAll('a').length > prev,
        { el: containerHandle, prev: prevCount },
        { timeout: 8000 }
      );
    } catch {
      // No new rows loaded within 8s — count as stable
    }

    const curCount = await countLinks();
    console.log(`[Scroll] Step ${totalScrolls} — links in container: ${curCount}`);

    if (curCount > prevCount) {
      prevCount    = curCount;
      stableRounds = 0;             // new data appeared — keep scrolling
    } else {
      stableRounds++;
      console.log(`[Scroll] No new rows (stable round ${stableRounds}/6)`);

      // On stable rounds, do an extra big jump to make sure we're at bottom
      if (stableRounds === 3) {
        console.log('[Scroll] Jumping to bottom to confirm end of list...');
        await page.evaluate(el => { el.scrollTop = el.scrollHeight; }, containerHandle);
        await page.waitForTimeout(5000); // give BNI extra time after jump
        const afterJump = await countLinks();
        console.log(`[Scroll] After bottom-jump: ${afterJump} links`);
        if (afterJump > prevCount) {
          prevCount    = afterJump;
          stableRounds = 0;           // more data loaded — restart
        }
      }
    }
  }

  await containerHandle.dispose();

  console.log(`[Search] Scroll complete after ${totalScrolls} steps.`);
  await page.waitForTimeout(2000);

  const finalCount = await page.evaluate(() =>
    document.querySelectorAll('tr td:first-child a, [role="row"] a').length
  );
  console.log(`[Search] After scroll: ~${finalCount} member links visible`);
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
    // FIX: previously assumed a fixed positional order (Chapter, Company,
    // City, Industry) after filtering blank text nodes. When BNI's DOM
    // genuinely omits a field (e.g. no Company text node rendered), every
    // field after it shifted left by one, corrupting Company/City/Industry.
    // Now we classify each text fragment by its CONTENT pattern instead of
    // its position, so missing fields stay missing instead of causing a shift.
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

      // Classify each fragment by pattern rather than assuming fixed position:
      //  - Industry/Classification always contains " > " (breadcrumb format)
      //  - City is usually short (<= 3-4 words) and doesn't contain " > "
      //  - Chapter is typically the BNI chapter name (often starts with "BNI"
      //    or is a single capitalized word/phrase, no commas)
      //  - Company is whatever's left — usually the longest remaining fragment
      let chapter = '', company = '', city = '', industry = '';
      const remaining = [];

      for (const text of cols) {
        if (text.includes(' > ')) {
          // This is the Industry/Classification breadcrumb
          if (!industry) industry = text;
          else remaining.push(text); // shouldn't normally happen
        } else {
          remaining.push(text);
        }
      }

      // From the non-industry fragments, identify Chapter (usually first,
      // often starts with "BNI" or has no comma/space-heavy city pattern)
      // and City (often contains a comma, or is the fragment right before
      // the industry breadcrumb, or matches common city-name patterns).
      if (remaining.length >= 3) {
        // Standard case: [Chapter, Company, City, ...]
        chapter = remaining[0] || '';
        company = remaining[1] || '';
        city    = remaining[2] || '';
      } else if (remaining.length === 2) {
        // One field is missing — figure out which by pattern.
        // If the 2nd fragment looks like a city (short, capitalized place
        // name, often single word or word+comma) and industry exists,
        // assume Company is the missing one and shift accordingly.
        const looksLikeCity = (t) =>
          t.length < 30 && /^[A-Za-z][A-Za-z\s.,'-]*$/.test(t) && t.split(' ').length <= 4;

        if (industry && looksLikeCity(remaining[1])) {
          // [Chapter, City] — Company missing
          chapter = remaining[0] || '';
          company = '';
          city    = remaining[1] || '';
        } else {
          // [Chapter, Company] — City missing
          chapter = remaining[0] || '';
          company = remaining[1] || '';
          city    = '';
        }
      } else if (remaining.length === 1) {
        chapter = remaining[0] || '';
      }

      const href = anchor.getAttribute('href') || '';
      seen.add(name);
      results.push({
        Name:                   name,
        Chapter:                chapter,
        Company:                company,
        City:                   city,
        IndustryClassification: industry,
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
  // ── Step 1: Try exact / near-exact text candidates first (fast path) ─────
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

  // ── Step 2: Fixed candidates failed — read ALL rendered options from the
  // dropdown and find the best fuzzy match. This handles cases where BNI
  // renders the option differently than our guessed patterns (different
  // separator, different casing, extra words, etc).
  console.warn(`[Search] Fixed candidates failed for "${categoryName}" — reading live dropdown options...`);

  const liveOptions = await page.evaluate(() => {
    return [...document.querySelectorAll('[role="option"]')]
      .map(el => el.textContent.trim())
      .filter(Boolean);
  });

  if (liveOptions.length === 0) {
    console.error('[Search] No dropdown options rendered at all.');
    return false;
  }

  console.log(`[Search] Live options found (${liveOptions.length}):`, liveOptions.slice(0, 15));

  // Extract the "core" identifying text from the requested category
  // e.g. "Legal & Accounting (Administrative Services)" -> "Administrative Services"
  const parenMatch  = categoryName.match(/\(([^)]+)\)/);
  const coreText    = (parenMatch ? parenMatch[1] : categoryName).toLowerCase().trim();
  const fullLower   = categoryName.toLowerCase();

  const normalize = s => s.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const coreNorm   = normalize(coreText);
  const fullNorm   = normalize(categoryName);

  // Score each live option: prefer ones containing the full normalized name,
  // then ones containing the core (parenthesized) part, then word-overlap.
  let bestOption = null;
  let bestScore  = 0;

  for (const opt of liveOptions) {
    const optNorm = normalize(opt);
    let score = 0;

    if (optNorm === fullNorm) score = 1000;                              // exact full match
    else if (optNorm.includes(fullNorm)) score = 900;                    // full contained
    else if (optNorm === coreNorm) score = 800;                          // exact core match
    else if (optNorm.includes(coreNorm) && coreNorm.length > 3) score = 700; // core contained
    else {
      // word overlap score — count shared significant words
      const optWords  = new Set(optNorm.split(' ').filter(w => w.length > 2));
      const coreWords = coreNorm.split(' ').filter(w => w.length > 2);
      const shared    = coreWords.filter(w => optWords.has(w)).length;
      if (shared > 0 && coreWords.length > 0) {
        score = 200 + (shared / coreWords.length) * 400; // up to 600
      }
    }

    if (score > bestScore) {
      bestScore  = score;
      bestOption = opt;
    }
  }

  // Require a reasonably confident match — refuse to guess on weak overlap
  const MIN_CONFIDENT_SCORE = 400;
  if (!bestOption || bestScore < MIN_CONFIDENT_SCORE) {
    console.error(
      `[Search] No confident match for "${categoryName}" among live options ` +
      `(best="${bestOption}" score=${bestScore}). Refusing to select.`
    );
    return false;
  }

  console.log(`[Search] Best fuzzy match: "${bestOption}" (score=${bestScore})`);

  try {
    await page.getByRole('option', { name: bestOption, exact: true }).first()
      .click({ timeout: 5000 });
    console.log(`[Search] Selected via fuzzy match: "${bestOption}"`);
    return true;
  } catch {
    try {
      await page.locator(`[role="option"]:has-text("${bestOption}")`).first()
        .click({ timeout: 5000 });
      console.log(`[Search] Selected via fuzzy match (locator): "${bestOption}"`);
      return true;
    } catch (err) {
      console.error(`[Search] Failed to click matched option: ${err.message}`);
      return false;
    }
  }
}



//updated one



