/**
 * Probe Slack sign-in flow for Varsity Tutors workspace.
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ARTIFACT_DIR = '/opt/cursor/artifacts/autonudge-slack-create';
fs.mkdirSync(ARTIFACT_DIR, { recursive: true });

async function shot(page, name) {
  const file = path.join(ARTIFACT_DIR, `${name}.png`);
  await page.screenshot({ path: file, fullPage: true });
  console.log('shot', file);
}

async function dump(page, name) {
  const info = {
    url: page.url(),
    title: await page.title().catch(() => ''),
    text: (await page.locator('body').innerText().catch(() => '')).slice(0, 4000),
  };
  fs.writeFileSync(path.join(ARTIFACT_DIR, `${name}.json`), JSON.stringify(info, null, 2));
  console.log(name, info.url, info.title);
}

(async () => {
  const browser = await chromium.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-dev-shm-usage'],
  });
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  page.on('framenavigated', (frame) => {
    if (frame === page.mainFrame()) console.log('nav ->', frame.url());
  });

  await page.goto('https://api.slack.com/apps', { waitUntil: 'domcontentloaded', timeout: 90000 });
  await page.waitForTimeout(2000);
  await shot(page, '10-apps-unsigned');
  await dump(page, '10-apps-unsigned');

  // Click sign in link
  const signIn = page.getByRole('link', { name: /sign in to your Slack account/i });
  if (await signIn.count()) {
    console.log('clicking sign in link');
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {}),
      signIn.click(),
    ]);
  } else {
    await page.goto('https://api.slack.com/signin?redir=%2Fapps', { waitUntil: 'domcontentloaded' });
  }
  await page.waitForTimeout(3000);
  await shot(page, '11-signin');
  await dump(page, '11-signin');

  // Try workspace domain entry if present
  const domain = page.locator('input[data-qa="signin_domain_input"], input[name="domain"], #domain');
  if (await domain.count()) {
    await domain.first().fill('varsitytutors');
    await shot(page, '12-domain-filled');
    const cont = page.locator('button[data-qa="submit_team_domain_button"], button:has-text("Continue"), button[type="submit"]');
    if (await cont.count()) {
      await Promise.all([
        page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {}),
        cont.first().click(),
      ]);
      await page.waitForTimeout(3000);
      await shot(page, '13-after-domain');
      await dump(page, '13-after-domain');
    }
  }

  // List all buttons/links for SSO options
  const controls = await page.evaluate(() => {
    const els = [...document.querySelectorAll('a, button, input, [role="button"]')];
    return els.slice(0, 80).map((el) => ({
      tag: el.tagName,
      type: el.getAttribute('type'),
      name: el.getAttribute('name'),
      id: el.id,
      qa: el.getAttribute('data-qa'),
      text: (el.innerText || el.value || '').trim().slice(0, 120),
      href: el.getAttribute('href'),
    }));
  });
  fs.writeFileSync(path.join(ARTIFACT_DIR, '14-controls.json'), JSON.stringify(controls, null, 2));
  console.log('controls', JSON.stringify(controls, null, 2));

  // Click common SSO
  for (const label of [
    /sign in with google/i,
    /google/i,
    /microsoft/i,
    /okta/i,
    /sso/i,
    /sign in with password/i,
    /workspace/i,
  ]) {
    const btn = page.getByRole('button', { name: label }).or(page.getByRole('link', { name: label }));
    if (await btn.count()) {
      console.log('trying control', label);
      await btn.first().click().catch(() => {});
      await page.waitForTimeout(2500);
      await shot(page, `15-try-${String(label)}`);
      await dump(page, `15-try-${String(label)}`);
      break;
    }
  }

  // Direct workspace sign-in URL
  await page.goto('https://varsitytutors.slack.com/sign_in_with_password', {
    waitUntil: 'domcontentloaded',
    timeout: 60000,
  }).catch(async (e) => {
    console.log('workspace password url failed', e.message);
    await page.goto('https://varsitytutors.slack.com/', { waitUntil: 'domcontentloaded', timeout: 60000 });
  });
  await page.waitForTimeout(3000);
  await shot(page, '16-workspace');
  await dump(page, '16-workspace');

  const controls2 = await page.evaluate(() => {
    const els = [...document.querySelectorAll('a, button, input, [role="button"]')];
    return els.slice(0, 100).map((el) => ({
      tag: el.tagName,
      type: el.getAttribute('type'),
      name: el.getAttribute('name'),
      id: el.id,
      qa: el.getAttribute('data-qa'),
      text: (el.innerText || el.value || '').trim().slice(0, 120),
      href: el.getAttribute('href'),
    }));
  });
  fs.writeFileSync(path.join(ARTIFACT_DIR, '17-workspace-controls.json'), JSON.stringify(controls2, null, 2));
  console.log('workspace controls', JSON.stringify(controls2, null, 2));

  await browser.close();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
