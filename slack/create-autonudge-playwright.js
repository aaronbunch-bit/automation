/**
 * Drive Varsity Tutors Slack Google SSO, then create AutoNudge from manifest.
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ARTIFACT_DIR = '/opt/cursor/artifacts/autonudge-slack-create';
const MANIFEST_PATH = path.join(__dirname, 'autonudge.manifest.yaml');
const EMAIL = process.env.SLACK_EMAIL || 'aaron.bunch@varsitytutors.com';
const PASSWORD = process.env.SLACK_PASSWORD || process.env.VT_SSO_PASSWORD || process.env.GOOGLE_PASSWORD || '';
const PROFILE_DIR = '/tmp/autonudge-chrome-profile-sso';

fs.mkdirSync(ARTIFACT_DIR, { recursive: true });
fs.mkdirSync(PROFILE_DIR, { recursive: true });

async function shot(page, name) {
  await page.screenshot({ path: path.join(ARTIFACT_DIR, `${name}.png`), fullPage: true }).catch(() => {});
  console.log('shot', name, page.url());
}

async function dump(page, name) {
  const info = {
    url: page.url(),
    title: await page.title().catch(() => ''),
    text: (await page.locator('body').innerText().catch(() => '')).slice(0, 5000),
  };
  fs.writeFileSync(path.join(ARTIFACT_DIR, `${name}.json`), JSON.stringify(info, null, 2));
  return info;
}

async function acceptCookies(page) {
  const btn = page.locator('#accept-recommended-btn-handler, button:has-text("Allow All")').first();
  if (await btn.isVisible().catch(() => false)) {
    await btn.click().catch(() => {});
    await page.waitForTimeout(800);
  }
}

async function createAppFromManifest(page, manifestYaml, result) {
  await page.goto('https://api.slack.com/apps', { waitUntil: 'domcontentloaded', timeout: 90000 });
  await page.waitForTimeout(2500);
  await shot(page, '70-apps-authed');
  await dump(page, '70-apps-authed');

  const text = ((await page.locator('body').innerText()) || '').toLowerCase();
  if (text.includes('sign in to your slack account') || !text.includes('create new app')) {
    throw new Error('Not authenticated on api.slack.com/apps');
  }

  await page.getByText(/create new app/i).first().click();
  await page.waitForTimeout(1500);
  await shot(page, '71-create-modal');
  await page.getByText(/from a manifest/i).first().click();
  await page.waitForTimeout(1500);
  await shot(page, '72-from-manifest');

  // Pick workspace if needed
  const ws = page.getByText(/varsity tutors/i).first();
  if (await ws.isVisible().catch(() => false)) {
    await ws.click().catch(() => {});
  }
  if (await page.getByRole('button', { name: /^next$/i }).count()) {
    await page.getByRole('button', { name: /^next$/i }).click();
    await page.waitForTimeout(1500);
  }

  await page.getByRole('tab', { name: /yaml/i }).click().catch(() => {});
  await page.waitForTimeout(500);

  const method = await page.evaluate((yaml) => {
    const cm = document.querySelector('.CodeMirror');
    if (cm && cm.CodeMirror) {
      cm.CodeMirror.setValue(yaml);
      return 'codemirror';
    }
    const ta = document.querySelector('textarea');
    if (ta) {
      const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set;
      setter.call(ta, yaml);
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      return 'textarea';
    }
    return null;
  }, manifestYaml);
  console.log('manifest fill', method);
  await shot(page, '73-filled');

  await page.getByRole('button', { name: /^create$/i }).first().click();
  await page.waitForTimeout(8000);
  await shot(page, '74-after-create');
  const created = await dump(page, '74-after-create');
  result.createUrl = created.url;
  result.createPreview = (created.text || '').slice(0, 2500);
  result.status = 'created_or_requested';
  result.message =
    'AutoNudge create submitted via Playwright. Slack will auto-notify workspace admins/IT for approval when required.';
}

async function main() {
  const manifestYaml = fs.readFileSync(MANIFEST_PATH, 'utf8');
  const result = {
    startedAt: new Date().toISOString(),
    email: EMAIL,
    hasPassword: Boolean(PASSWORD),
    steps: [],
  };

  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    executablePath: '/usr/local/bin/google-chrome',
    viewport: { width: 1440, height: 900 },
    args: ['--no-sandbox', '--disable-dev-shm-usage', '--window-size=1440,900'],
    ignoreDefaultArgs: ['--enable-automation'],
    env: { ...process.env, DISPLAY: ':1' },
    userAgent:
      'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36',
  });
  const page = context.pages()[0] || (await context.newPage());

  try {
    await page.goto('https://varsitytutors.slack.com/', {
      waitUntil: 'domcontentloaded',
      timeout: 90000,
    });
    await page.waitForTimeout(2500);
    await acceptCookies(page);
    await shot(page, '60-workspace');
    await dump(page, '60-workspace');
    result.steps.push('workspace');

    // Click Google SSO on workspace login
    const googleBtn = page.locator(
      '[data-qa="base_google_login_button"], #google_login_button, button:has-text("Google")'
    ).first();
    if (!(await googleBtn.isVisible().catch(() => false))) {
      throw new Error('Google SSO button not visible on workspace login');
    }

    const popupPromise = context.waitForEvent('page', { timeout: 15000 }).catch(() => null);
    await googleBtn.click();
    result.steps.push('clicked_google');
    await page.waitForTimeout(2000);

    let authPage = (await popupPromise) || page;
    // If popup opened, use it; also handle same-tab redirect
    if (authPage !== page) {
      await authPage.waitForLoadState('domcontentloaded').catch(() => {});
      console.log('popup url', authPage.url());
    } else {
      await page.waitForTimeout(3000);
      console.log('same-tab url', page.url());
    }

    await shot(authPage, '61-google');
    await dump(authPage, '61-google');

    // Fill Google email
    const emailInput = authPage.locator('#identifierId, input[type="email"]').first();
    if (await emailInput.isVisible().catch(() => false)) {
      await emailInput.fill(EMAIL);
      await authPage.locator('#identifierNext, button:has-text("Next")').first().click();
      result.steps.push('google_email');
      await authPage.waitForTimeout(4000);
      await shot(authPage, '62-google-after-email');
      await dump(authPage, '62-google-after-email');
    }

    if (PASSWORD) {
      const pass = authPage.locator('input[type="password"], input[name="Passwd"]').first();
      if (await pass.isVisible().catch(() => false)) {
        await pass.fill(PASSWORD);
        await authPage.locator('#passwordNext, button:has-text("Next")').first().click();
        result.steps.push('google_password');
        await authPage.waitForTimeout(6000);
      }
    }

    // Wait for redirect back to Slack / apps (up to 3 min for MFA on VNC)
    console.log('Waiting for Slack session after Google SSO...');
    const deadline = Date.now() + 180000;
    let authed = false;
    while (Date.now() < deadline) {
      for (const p of context.pages()) {
        const url = p.url();
        const body = ((await p.locator('body').innerText().catch(() => '')) || '').toLowerCase();
        if (
          (url.includes('varsitytutors.slack.com') && !url.includes('sign_in') && body.includes('slack')) ||
          (url.includes('api.slack.com') && body.includes('create new app') && !body.includes('sign in to your slack account'))
        ) {
          // Heuristic: client loaded or apps page signed in
          if (!body.includes('sign in to varsity tutors') && !body.includes('sign in with email')) {
            authed = true;
            break;
          }
        }
        // Also detect Google password / challenge wall
        if (url.includes('accounts.google.com') && (body.includes('enter your password') || body.includes('verify'))) {
          await shot(p, '63-google-challenge');
          await dump(p, '63-google-challenge');
        }
      }
      if (authed) break;
      await page.waitForTimeout(4000);
    }

    // Explicit check of apps page
    const appsPage = context.pages().find((p) => p.url().includes('api.slack.com')) || page;
    await appsPage.goto('https://api.slack.com/apps', { waitUntil: 'domcontentloaded', timeout: 90000 }).catch(() => {});
    await page.waitForTimeout(2500);
    await shot(appsPage, '64-apps-check');
    const appsState = await dump(appsPage, '64-apps-check');
    const signedIn =
      (appsState.text || '').toLowerCase().includes('create new app') &&
      !(appsState.text || '').toLowerCase().includes('sign in to your slack account');

    if (!signedIn) {
      result.status = 'blocked_auth';
      result.message =
        'Google SSO reached the identity provider, but this environment has no Google/SSO password or MFA factor. AutoNudge app creation cannot complete until SLACK_PASSWORD or VT_SSO_PASSWORD (and MFA if required) is available to the agent.';
      fs.writeFileSync(path.join(ARTIFACT_DIR, 'result.json'), JSON.stringify(result, null, 2));
      console.log(JSON.stringify(result, null, 2));
      await context.close();
      process.exit(2);
    }

    await createAppFromManifest(appsPage, manifestYaml, result);
    fs.writeFileSync(path.join(ARTIFACT_DIR, 'result.json'), JSON.stringify(result, null, 2));
    console.log(JSON.stringify(result, null, 2));
    await context.close();
    process.exit(0);
  } catch (err) {
    result.status = 'error';
    result.error = String(err && err.stack ? err.stack : err);
    await shot(page, '99-error');
    await dump(page, '99-error').catch(() => {});
    fs.writeFileSync(path.join(ARTIFACT_DIR, 'result.json'), JSON.stringify(result, null, 2));
    console.error(result.error);
    await context.close().catch(() => {});
    process.exit(1);
  }
}

main();
