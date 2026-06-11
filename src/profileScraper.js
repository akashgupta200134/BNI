// src/profileScraper.js
import cfg           from '../config.js';
import * as fileUtils from './fileUtils.js';
import * as loginMod  from './login.js';

export async function openProfile(context, profileUrl) {
  console.log(`[Profile] Opening: ${profileUrl}`);
  const profilePage = await context.newPage();
  profilePage.setDefaultTimeout(cfg.ACTION_TIMEOUT);
  profilePage.setDefaultNavigationTimeout(cfg.NAV_TIMEOUT);
  await profilePage.goto(profileUrl, { waitUntil: 'domcontentloaded', timeout: cfg.NAV_TIMEOUT });
  return profilePage;
}

export async function scrapeProfilePage(profilePage) {
  console.log('[Profile] Scraping...');

  try {
    await profilePage.waitForLoadState('domcontentloaded', { timeout: cfg.NAV_TIMEOUT });
    await profilePage.waitForTimeout(5000);

    // Session expiry check
    const originalUrl = profilePage.url();
    if (originalUrl.includes('/login')) {
      const wasExpired = await loginMod.reAuthIfNeeded(profilePage);
      if (wasExpired) {
        await profilePage.goto(originalUrl, { waitUntil: 'domcontentloaded', timeout: cfg.NAV_TIMEOUT });
        await profilePage.waitForTimeout(3000);
      }
    }

    await profilePage.waitForSelector('#root', { timeout: 10000 }).catch(() => {});
    await profilePage.waitForTimeout(2000);

    const profileText = await extractProfileText(profilePage);
    fileUtils.writeTextFile(cfg.DATA_TXT, profileText);
    console.log(`[Profile] Got ${profileText.length} chars`);

    const contacts = fileUtils.parseContactDetails(profileText);
    console.log(`  Email:   ${contacts.email}`);
    console.log(`  Email2:  ${contacts.email2 || '—'}`);
    console.log(`  Phone:   ${contacts.phone}`);
    console.log(`  Phone2:  ${contacts.phone2 || '—'}`);
    console.log(`  Website: ${contacts.website}`);
    return contacts;

  } catch (err) {
    console.error('[Profile] Scrape failed:', err.message);
    return { email: 'Not Found', email2: '', phone: 'Not Found', phone2: '', website: 'Not Found' };
  }
}

async function extractProfileText(page) {
  return await page.evaluate(() => {
    const collected = [];

    // ── Priority 1: Direct href links (mailto:, tel:, https) ──────────────
    // These are the most reliable — BNI puts actual contact links in the profile
    document.querySelectorAll('a[href]').forEach(a => {
      const href = (a.getAttribute('href') || '').trim();
      if (href.startsWith('mailto:')) {
        const email = href.replace('mailto:', '').split('?')[0].trim();
        if (email) collected.push(`Email: ${email}`);
      }
      if (href.startsWith('tel:')) {
        const phone = href.replace('tel:', '').trim();
        if (phone) collected.push(`Phone: ${phone}`);
      }
      if (href.startsWith('http') &&
          !href.includes('bniconnectglobal') &&
          !href.includes('bni.com') &&
          !href.includes('bni.in') &&
          !href.includes('javascript:')) {
        collected.push(`Website: ${href}`);
      }
    });

    // ── Priority 2: PAD's exact CSS path ──────────────────────────────────
    try {
      const el = document.querySelector(
        'body > div:first-child > div > div > div:nth-child(2) > div:nth-child(2) > div:first-child > div:nth-child(2) > div > div:first-child'
      );
      if (el && el.innerText.trim().length > 20) collected.push(el.innerText);
    } catch {}

    // ── Priority 3: Profile/contact class containers ───────────────────────
    try {
      document.querySelectorAll(
        '[class*="profile"], [class*="contact"], [class*="member"], [class*="networkHome"]'
      ).forEach(s => {
        const t = (s.innerText || '').trim();
        if (t.length > 20) collected.push(t);
      });
    } catch {}

    // ── Priority 4: Full body text (catch-all for regex) ──────────────────
    collected.push(document.body?.innerText || '');

    return collected.join('\n');
  });
}