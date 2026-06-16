// INSIGHT M2 DEMO -- NOT PRODUCTION CODE
// app.js — browser-side demo logic
// Uses MSAL Browser (CDN) for Entra auth; calls /api/classlink for ClassLink data.

'use strict';

// ── Credentials ────────────────────────────────────────────────────────────────
const ENTRA_TENANT_ID = '04d60e3b-0fed-4071-a820-e4e87e8f8c98';
const ENTRA_CLIENT_ID = '7eb737a9-32f2-4298-b617-87d35051d95d';
const ENTRA_SCOPES    = ['openid', 'email', 'profile'];

// ── App state ──────────────────────────────────────────────────────────────────
const state = {
    currentScreen: 1,
    teacher: null,   // { email, displayName, sourcedId, classes, selectedClass }
    student: null,   // { email, displayName, sourcedId, enrolledClass }
    sync:    null    // { teacherCount, classCount, syncedAt }
};
// All screens freely navigable — sign-in enriches data but doesn't block navigation
const screenReady = { 1: true, 2: true, 3: true, 4: true, 5: true };

// ── MSAL setup (v2) ────────────────────────────────────────────────────────────
// Uses MSAL Browser v2 CDN (v3 dropped the UMD bundle so CDN use requires v2)
let msalApp = null;

function getMsalApp() {
    if (msalApp) return msalApp;
    if (typeof msal === 'undefined') {
        throw new Error('MSAL library failed to load. Check your network connection and refresh.');
    }
    msalApp = new msal.PublicClientApplication({
        auth: {
            clientId:    ENTRA_CLIENT_ID,
            authority:   `https://login.microsoftonline.com/${ENTRA_TENANT_ID}`,
            redirectUri: window.location.origin
        },
        cache: { cacheLocation: 'sessionStorage' }
    });
    return msalApp;
}

async function signIn() {
    const app    = getMsalApp();
    const result = await app.loginPopup({
        scopes: ENTRA_SCOPES,
        prompt: 'select_account'
    });
    const claims      = result.idTokenClaims || {};
    const email       = (claims.email || claims.preferred_username || result.account?.username || '').toLowerCase();
    const displayName = claims.name || email;
    return { email, displayName };
}

// ── ClassLink API helper (via Vercel serverless function) ──────────────────────
async function clApi(op, email) {
    const url = `/api/classlink?op=${op}${email ? '&email=' + encodeURIComponent(email) : ''}`;
    const res = await fetch(url);
    if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        throw new Error(body.error || `API error: HTTP ${res.status}`);
    }
    return res.json();
}

// ── Screen navigation ──────────────────────────────────────────────────────────
function navigate(delta) {
    const next = state.currentScreen + delta;
    if (next < 1 || next > 5) return;
    showScreen(next);
}

function showScreen(n) {
    document.querySelectorAll('.screen').forEach(el => el.classList.remove('active'));
    document.getElementById(`screen-${n}`).classList.add('active');
    state.currentScreen = n;

    for (let i = 1; i <= 5; i++) {
        const dot = document.getElementById(`dot-${i}`);
        const lbl = document.getElementById(`lbl-${i}`);
        const con = document.getElementById(`conn-${i}`);
        dot.classList.remove('active', 'done');
        lbl.classList.remove('active');
        if (con) con.classList.remove('done');
        if (i < n)  { dot.classList.add('done'); dot.innerHTML = '<i class="fas fa-check" style="font-size:.65rem"></i>'; if (con) con.classList.add('done'); }
        if (i === n) { dot.classList.add('active'); dot.textContent = i; lbl.classList.add('active'); }
        if (i > n)  { dot.textContent = i; }
    }

    if (n === 4) {
        const tMode = document.querySelector('input[name="t-signin-mode"]:checked')?.value || 'standard';
        polTeacherSignInMode(tMode);
    }
    if (n === 5) {
        updateSummary();
        const sMode = document.querySelector('input[name="s-auth-mode"]:checked')?.value || 'legacy';
        polStudentSignInMode(sMode);
    }
    updateNav();
    window.scrollTo(0, 0);
}

function updateNav() {
    const n       = state.currentScreen;
    const backBtn = document.getElementById('btn-back');
    const nextBtn = document.getElementById('btn-next');
    const soBtn   = document.getElementById('btn-start-over');
    const hint    = document.getElementById('footer-hint');
    backBtn.style.display = n > 1 ? '' : 'none';
    soBtn.style.display   = n === 5 ? '' : 'none';
    if (n === 5) { nextBtn.style.display = 'none'; return; }
    nextBtn.style.display = '';
    nextBtn.disabled = false;
    hint.textContent = '';
}

// ── Tab switcher for real TC UI ────────────────────────────────────────────────
function tcSwitchTab(tab) {
    document.getElementById('tab-classids').classList.toggle('active', tab === 'classids');
    document.getElementById('tab-classrooms').classList.toggle('active', tab === 'classrooms');
    document.getElementById('tc-view-classids').style.display   = tab === 'classids'   ? '' : 'none';
    document.getElementById('tc-view-classrooms').style.display = tab === 'classrooms' ? '' : 'none';
}

// ── Teacher sign-in (Scenario 1 — real MSAL + live ClassLink) ─────────────────
// Called from the Scenario 1 step-1 button. Operates entirely inside the sim zone
// so everything stays inside the console window mockup.
async function teacherSignIn() {
    const zone = document.getElementById('tc-sim-zone');

    // Loading state
    if (zone) zone.innerHTML = `
        <div style="text-align:center;padding:28px 0">
            <i class="fas fa-spinner fa-spin" style="color:#1F5C99;font-size:22px"></i>
            <div style="margin-top:10px;font-size:12px;color:#6b7280">Opening Microsoft sign-in…</div>
        </div>`;

    try {
        const msalResult = await signIn();

        if (zone) zone.innerHTML = `
            <div style="text-align:center;padding:28px 0">
                <i class="fas fa-spinner fa-spin" style="color:#1F5C99;font-size:22px"></i>
                <div style="margin-top:10px;font-size:12px;color:#6b7280">Fetching your ClassLink classrooms…</div>
            </div>`;

        const clData = await clApi('teacher', msalResult.email);

        state.teacher = {
            email:        msalResult.email,
            displayName:  msalResult.displayName,
            sourcedId:    clData.sourcedId,
            classes:      clData.classes || [],
            selectedClass: clData.classes && clData.classes.length > 0 ? clData.classes[0] : null
        };

        // Advance Scenario 1 to step 2 — step 2 will use state.teacher.classes for real data
        activeSimStep  = 2;
        simClassChosen = false;
        renderSim(true);
        
        updateNav();
        updateSummary();

    } catch (err) {
        const msg = classifyError(err);
        // Restore step 1 with error banner so user can retry
        if (zone) {
            zone.classList.add('transitioning');
            setTimeout(() => {
                zone.innerHTML = _s1Step1Html() +
                    `<div class="alert alert-danger py-2 px-3 small mt-2">${esc(msg)}</div>`;
                updateSimDots();
                zone.classList.remove('transitioning');
            }, 100);
        }
    }
}

// ── Student sign-in ────────────────────────────────────────────────────────────
async function studentSignIn() {
    const zone = document.getElementById('s2-sim-zone');

    if (zone) zone.innerHTML = `
        <div style="text-align:center;padding:28px 0">
            <i class="fas fa-spinner fa-spin" style="color:#ce4900;font-size:22px"></i>
            <div style="margin-top:10px;font-size:12px;color:#6b7280">Opening Microsoft sign-in…</div>
        </div>`;

    try {
        const msalResult = await signIn();

        if (zone) zone.innerHTML = `
            <div style="text-align:center;padding:28px 0">
                <i class="fas fa-spinner fa-spin" style="color:#ce4900;font-size:22px"></i>
                <div style="margin-top:10px;font-size:12px;color:#6b7280">Fetching your enrolled class…</div>
            </div>`;

        const clData = await clApi('student', msalResult.email);

        state.student = {
            email:        msalResult.email,
            displayName:  msalResult.displayName,
            sourcedId:    clData.sourcedId,
            enrolledClass: clData.enrolledClass
        };

        // Update Screen 5 class ID field if rostering data exists
        var sciField = document.getElementById('s-class-id');
        if (sciField && clData.enrolledClass) sciField.value = 'cl_' + clData.enrolledClass.sourcedId;

        // Advance S1 to step 2 (shows real data)
        activeSimStepS2 = 2;
        renderSimS2(true);

        updateNav();
        updateSummary();
    } catch (err) {
        const msg = classifyError(err);
        if (zone) {
            zone.classList.add('transitioning');
            setTimeout(() => {
                zone.innerHTML = _s2Step1Html() +
                    `<div class="alert alert-danger py-2 px-3 small mt-2">${esc(msg)}</div>`;
                updateSimDotsS2();
                zone.classList.remove('transitioning');
            }, 100);
        }
    }
}

// ── ClassLink sync (Screen 3) ──────────────────────────────────────────────────
async function triggerSync() {
    const btn     = document.getElementById('btn-sync-now');
    const spinner = document.getElementById('sync-spinner');
    const icon    = document.getElementById('sync-icon');
    const msgEl   = document.getElementById('sync-message');
    if (!btn) return;
    btn.disabled = true;
    if (icon)    icon.classList.add('d-none');
    if (spinner) spinner.classList.remove('d-none');
    if (msgEl)   msgEl.textContent = 'Connecting to ClassLink…';

    try {
        const data = await clApi('sync');
        state.sync = data;

        // Update the new DFC-style sync status elements
        const timeEl    = document.getElementById('last-sync-time');
        const badgeEl   = document.getElementById('sync-status-badge');
        const countsEl  = document.getElementById('sync-counts');

        if (timeEl)   timeEl.textContent  = data.syncedAt;
        if (badgeEl)  { badgeEl.style.display = 'inline-flex'; }
        if (countsEl) {
            countsEl.style.display = 'inline';
            const students = data.studentCount != null ? `${data.studentCount} students` : '— students';
            countsEl.textContent = `${data.teacherCount} teachers  ·  ${data.classCount} classes  ·  ${students}`;
        }
        if (msgEl)   msgEl.textContent = '';
    } catch (err) {
        if (msgEl) msgEl.textContent = err.message;
    } finally {
        btn.disabled = false;
        if (icon)    icon.classList.remove('d-none');
        if (spinner) spinner.classList.add('d-none');
    }
}

// ── Screens 4 & 5: Policy page interactions ───────────────────────────────────

function polToggle(headEl) {
    headEl.classList.toggle('collapsed');
    const body = headEl.nextElementSibling;
    if (body) body.style.display = headEl.classList.contains('collapsed') ? 'none' : '';
}

function polToggleInput(cb, inputId) {
    const input = document.getElementById(inputId);
    if (!input) return;
    const row = input.closest('.pol-row');
    input.disabled = !cb.checked;
    if (row) row.style.opacity = cb.checked ? '1' : '0.5';
}

// Teacher sign-in mode change
function polTeacherSignInMode(mode) {
    const rosterCb   = document.getElementById('t-roster-cb');
    const rosterWarn = document.getElementById('t-roster-sso-warn');
    const ssoOrgWarn = document.getElementById('t-sso-org-warn');
    const ssoOpts    = document.getElementById('t-sso-opts');
    const pwdRow         = document.getElementById('t-pwd-row');
    const rosterLabelWrap = document.getElementById('t-roster-label-wrap');
    const isSso          = mode !== 'standard';

    // Show/hide "Use logged-in computer account" sub-option
    if (ssoOpts) ssoOpts.style.display = isSso ? 'block' : 'none';

    // Hide Teacher Password when SSO Required (password is unused in SSO flow)
    if (pwdRow) pwdRow.style.display = isSso ? 'none' : '';

    // Disable rostering checkbox in Standard mode so it can't be clicked
    if (rosterCb) {
        rosterCb.disabled = !isSso;
        rosterCb.style.opacity = isSso ? '' : '0.35';
        rosterCb.style.cursor  = isSso ? '' : 'not-allowed';
    }

    // Show tooltip on the disabled rostering label only when SSO is not active
    if (rosterLabelWrap) rosterLabelWrap.className = isSso ? '' : 'pol-tooltip-wrap';

    // Clear the SSO warning whenever mode changes (it'll re-show if user clicks while blocked)
    if (rosterWarn) rosterWarn.style.display = 'none';

    // If switching back to Standard while rostering was on — uncheck it
    if (!isSso && rosterCb && rosterCb.checked) {
        rosterCb.checked = false;
        polTeacherRostering(rosterCb);
    }

    if (ssoOrgWarn) {
        ssoOrgWarn.style.display = (isSso && !state.teacher) ? 'block' : 'none';
    }
}

// Show/hide password field
function polShowPwd(inputId, cb) {
    const el = document.getElementById(inputId);
    if (el) el.type = cb.checked ? 'text' : 'password';
}

// Private Mode toggle — enables access token field and its Show Password checkbox
function polPrivMode(cb) {
    const token    = document.getElementById('t-access-token');
    const showCb   = document.getElementById('t-show-token');
    const showLbl  = showCb ? showCb.nextElementSibling : null;
    if (token)   { token.disabled = !cb.checked; token.style.background = cb.checked ? '#fff' : '#f0f0f0'; }
    if (showCb)  { showCb.disabled = !cb.checked; }
    if (showLbl) { showLbl.style.color = cb.checked ? '#2E3A40' : '#929292'; }
}

// Teacher Password checkbox — enables field and its Show Password
function polTeacherPwd(cb) {
    const field  = document.getElementById('t-pwd-field');
    const showCb = document.getElementById('t-show-pwd');
    const showLbl = showCb ? showCb.nextElementSibling : null;
    if (field)   { field.disabled = !cb.checked; field.style.background = cb.checked ? '#fff' : '#f0f0f0'; }
    if (showCb)  { showCb.disabled = !cb.checked; }
    if (showLbl) { showLbl.style.color = cb.checked ? '#2E3A40' : '#929292'; }
}

// Student Policy — Private Mode toggle (mirrors polPrivMode for Teacher)
function polStudentPrivMode(cb) {
    const token   = document.getElementById('s-access-token');
    const showCb  = document.getElementById('s-show-token');
    const showLbl = showCb ? showCb.nextElementSibling : null;
    if (token)   { token.disabled = !cb.checked; token.style.background = cb.checked ? '#fff' : '#f0f0f0'; }
    if (showCb)  { showCb.disabled = !cb.checked; }
    if (showLbl) { showLbl.style.color = cb.checked ? '#2E3A40' : '#929292'; }
}

// Student Policy — Automatically exit class: enable select + extend sub-option
function polStudentAutoExit(cb) {
    const mins   = document.getElementById('s-exit-mins');
    const extRow = document.getElementById('s-extend-row');
    const ext    = document.getElementById('s-extend');
    if (mins)   mins.disabled = !cb.checked;
    if (extRow) { extRow.style.opacity = cb.checked ? '1' : '0.5'; extRow.style.pointerEvents = cb.checked ? '' : 'none'; }
    if (ext)    ext.disabled = !cb.checked;
}

// Student Policy — Insight Connector enable/disable
function polStudentConnector(cb) {
    const ip = document.getElementById('s-connector-ip');
    if (ip) { ip.disabled = !cb.checked; ip.style.background = cb.checked ? '#fff' : '#f0f0f0'; }
}

function polToggleConnector(cb) {
    const ip = document.getElementById('t-connector-ip');
    if (ip) {
        ip.disabled = !cb.checked;
        ip.style.background = cb.checked ? '#fff' : '#f0f0f0';
    }
}

// Legacy compat stubs
function polEnableTeacherSSO(cb) {}
function polTeacherAuthMode(mode) {}

function polTeacherRostering(cb) {
    const opts      = document.getElementById('t-roster-opts');
    const ssoWarn   = document.getElementById('t-roster-sso-warn');
    const orgWarn   = document.getElementById('t-roster-org-warn');
    const mode      = document.querySelector('input[name="t-signin-mode"]:checked')?.value || 'standard';
    const isSso     = mode !== 'standard';
    const clConfigured = state.sync !== null; // sync has run = ClassLink configured

    if (cb.checked && !isSso) {
        // Block: SSO not enabled
        cb.checked = false;
        if (ssoWarn) ssoWarn.style.display = 'block';
        if (orgWarn) orgWarn.style.display = 'none';
        if (opts)    opts.style.display    = 'none';
        return;
    }

    if (ssoWarn) ssoWarn.style.display = 'none';

    if (cb.checked && !clConfigured) {
        // Show immediate org-settings warning for ClassLink — but still show fallback options
        if (orgWarn) orgWarn.style.display = 'block';
    } else {
        if (orgWarn) orgWarn.style.display = 'none';
    }

    if (opts) opts.style.display = cb.checked ? 'block' : 'none';
}

// Save validation for teacher policy
function polSaveTeacher() {
    const mode       = document.querySelector('input[name="t-signin-mode"]:checked')?.value || 'standard';
    const rosterOn   = document.getElementById('t-roster-cb')?.checked || false;
    const errContainer = document.getElementById('t-save-errors');
    if (!errContainer) return;

    const errors = [];

    // SSO: block save if non-standard mode and not configured
    if (mode !== 'standard' && !state.teacher) {
        errors.push({
            icon: 'fab fa-microsoft',
            msg: 'Policy cannot be saved until Microsoft Entra ID SSO is configured in Organization Settings.',
            link: true
        });
    }

    // ClassLink: block save if rostering enabled and sync hasn't run
    if (rosterOn && !state.sync) {
        errors.push({
            icon: 'fas fa-sync',
            msg: 'Policy cannot be saved until ClassLink Rostering is configured in Organization Settings.',
            link: true
        });
    }

    if (errors.length === 0) {
        errContainer.style.display = 'none';
        errContainer.innerHTML = '';
        // Show brief success
        errContainer.style.display = 'block';
        errContainer.innerHTML = `<div style="display:flex;align-items:center;gap:8px;padding:7px 10px;background:#d1fae5;border:1px solid #a7f3d0;border-radius:3px;font-size:11px;color:#065f46"><i class="fas fa-check-circle"></i> Policy saved successfully.</div>`;
        setTimeout(() => { errContainer.style.display = 'none'; }, 3000);
        return;
    }

    const html = errors.map(e => `
        <div style="display:flex;align-items:flex-start;gap:7px;padding:7px 10px;background:#fef3c7;border:1px solid #fcd34d;border-radius:3px;font-size:11px;color:#92400e;margin-bottom:4px">
            <i class="${e.icon}" style="margin-top:1px;flex-shrink:0"></i>
            <span>${esc(e.msg)} ${e.link ? `<a href="https://www.deepfreeze.com/NU/Site/OrganizationSettings" target="_blank" rel="noopener" style="color:#1d4ed8;font-weight:700">Go to Organization Settings <i class="fas fa-external-link-alt" style="font-size:9px"></i></a>` : ''}</span>
        </div>`).join('');

    errContainer.style.display = 'block';
    errContainer.innerHTML = html;
}

// Toggle entire group on/off (for on-prem section)
function polToggleGroup(cb, groupId) {
    const group = document.getElementById(groupId);
    if (!group) return;
    group.style.opacity = cb.checked ? '1' : '0.5';
    group.style.pointerEvents = cb.checked ? '' : 'none';
    group.querySelectorAll('select,input').forEach(el => el.disabled = !cb.checked);
}

// Student Sign-In Mode change (mirrors polTeacherSignInMode)
function polStudentSignInMode(mode) {
    const ssoWarn        = document.getElementById('s-sso-org-warn');
    const rosterCb       = document.getElementById('s-roster-cb');
    const rosterSsoWarn  = document.getElementById('s-roster-sso-warn');
    const ssoOpts        = document.getElementById('s-sso-opts');
    const rosterLabelWrap = document.getElementById('s-roster-label-wrap');
    const isSso          = mode !== 'legacy';

    if (ssoWarn) ssoWarn.style.display = (isSso && !state.teacher) ? 'block' : 'none';

    // Show/hide "Auto-assign Class ID via Rostering" sub-option
    if (ssoOpts) ssoOpts.style.display = isSso ? 'block' : 'none';

    // Disable rostering checkbox in Standard mode so it can't be clicked
    if (rosterCb) {
        rosterCb.disabled = !isSso;
        rosterCb.style.opacity = isSso ? '' : '0.35';
        rosterCb.style.cursor  = isSso ? '' : 'not-allowed';
    }

    // Show tooltip on the disabled rostering label only when SSO is not active
    if (rosterLabelWrap) rosterLabelWrap.className = isSso ? '' : 'pol-tooltip-wrap';

    // Clear the SSO warning whenever mode changes
    if (rosterSsoWarn) rosterSsoWarn.style.display = 'none';

    // If switching back to Standard while rostering was on — uncheck it
    if (!isSso && rosterCb && rosterCb.checked) {
        rosterCb.checked = false;
        polStudentRostering(rosterCb);
    }
}

// Student Rostering enable/disable (mirrors polTeacherRostering)
function polStudentRostering(cb) {
    const ssoWarn  = document.getElementById('s-roster-sso-warn');
    const orgWarn  = document.getElementById('s-roster-org-warn');
    const mode     = document.querySelector('input[name="s-auth-mode"]:checked')?.value || 'legacy';
    const isSso    = mode !== 'legacy';
    const clOk     = state.sync !== null;

    if (cb.checked && !isSso) {
        cb.checked = false;
        if (ssoWarn) ssoWarn.style.display = 'block';
        if (orgWarn) orgWarn.style.display = 'none';
        return;
    }
    if (ssoWarn) ssoWarn.style.display = 'none';
    if (cb.checked && !clOk) {
        if (orgWarn) orgWarn.style.display = 'block';
    } else {
        if (orgWarn) orgWarn.style.display = 'none';
    }
}

// Save validation for student policy
function polSaveStudent() {
    const mode        = document.querySelector('input[name="s-auth-mode"]:checked')?.value || 'legacy';
    const rosterOn    = document.getElementById('s-roster-cb')?.checked || false;
    const errContainer = document.getElementById('s-save-errors');
    if (!errContainer) return;

    const errors = [];

    if (mode !== 'legacy' && !state.teacher) {
        errors.push({
            icon: 'fab fa-microsoft',
            msg: 'Policy cannot be saved until Microsoft Entra ID SSO is configured in Organization Settings.',
            link: true
        });
    }
    if (rosterOn && !state.sync) {
        errors.push({
            icon: 'fas fa-sync',
            msg: 'Policy cannot be saved until ClassLink Rostering is configured in Organization Settings.',
            link: true
        });
    }

    if (errors.length === 0) {
        errContainer.style.display = 'block';
        errContainer.innerHTML = `<div style="display:flex;align-items:center;gap:8px;padding:7px 10px;background:#d1fae5;border:1px solid #a7f3d0;border-radius:3px;font-size:11px;color:#065f46"><i class="fas fa-check-circle"></i> Policy saved successfully.</div>`;
        setTimeout(() => { errContainer.style.display = 'none'; }, 3000);
        return;
    }

    const html = errors.map(e => `
        <div style="display:flex;align-items:flex-start;gap:7px;padding:7px 10px;background:#fef3c7;border:1px solid #fcd34d;border-radius:3px;font-size:11px;color:#92400e;margin-bottom:4px">
            <i class="${e.icon}" style="margin-top:1px;flex-shrink:0"></i>
            <span>${esc(e.msg)} ${e.link ? `<a href="https://www.deepfreeze.com/NU/Site/OrganizationSettings" target="_blank" rel="noopener" style="color:#1d4ed8;font-weight:700">Go to Organization Settings <i class="fas fa-external-link-alt" style="font-size:9px"></i></a>` : ''}</span>
        </div>`).join('');

    errContainer.style.display = 'block';
    errContainer.innerHTML = html;
}

// Legacy — no longer called from HTML but kept to avoid errors
function polEnableStudentSSO() {}
function polStudentAuthMode(mode) { polStudentSignInMode(mode); }

function polStudentAuthMode(mode) {
    const autoJoinRow = document.getElementById('s-autojoin-row');
    const classIdRow  = document.getElementById('s-class-id');
    if (autoJoinRow) autoJoinRow.style.display = mode === 'sso_required' ? 'block' : 'none';
    if (classIdRow)  classIdRow.disabled = mode === 'sso_required';
}

// ── Screen 3: DFC Org Settings interactions ───────────────────────────────────

function dfcSwitchTab(tabId, btn) {
    // Tab buttons
    document.querySelectorAll('#dfc-insight-tabs .dfc-tab').forEach(t => t.classList.remove('active'));
    btn.classList.add('active');
    // Tab panes
    document.querySelectorAll('.dfc-tab-pane').forEach(p => p.classList.remove('active'));
    const pane = document.getElementById(tabId);
    if (pane) pane.classList.add('active');
}

function dfcSwitchProvider(provider) {
    // Pills
    document.querySelectorAll('.dfc-provider-pill').forEach(p => p.classList.remove('active'));
    const pill = document.getElementById('pill-' + provider.replace('-sso', '').replace('google', 'google'));
    // Map provider names to pill IDs
    const pillMap = { 'entra': 'pill-entra', 'google-sso': 'pill-google' };
    const pillEl = document.getElementById(pillMap[provider]);
    if (pillEl) pillEl.classList.add('active');
    // Panels
    document.querySelectorAll('.dfc-provider-pane').forEach(p => p.classList.remove('active'));
    const paneMap = { 'entra': 'pane-entra', 'google-sso': 'pane-google-sso' };
    const pane = document.getElementById(paneMap[provider]);
    if (pane) pane.classList.add('active');
}

function dfcToggle(toggleId, statusId, fieldsId) {
    const checked = document.getElementById(toggleId).checked;
    const statusEl = document.getElementById(statusId);
    if (statusEl) {
        statusEl.textContent = checked ? 'On' : 'Off';
        statusEl.style.color = checked ? '#2E78C1' : '#9ca3af';
    }
    const fieldsEl = document.getElementById(fieldsId);
    if (fieldsEl) {
        fieldsEl.style.opacity = checked ? '1' : '0.4';
        fieldsEl.style.pointerEvents = checked ? '' : 'none';
    }
}

function dfcToggleSecret(inputId, btn) {
    const input = document.getElementById(inputId);
    const isPassword = input.type === 'password';
    input.type = isPassword ? 'text' : 'password';
    const icon = btn.querySelector('i');
    if (icon) icon.className = isPassword ? 'fas fa-eye-slash' : 'fas fa-eye';
}

function dfcCopy(text, btn) {
    navigator.clipboard.writeText(text).then(() => {
        const icon = btn.querySelector('i');
        if (icon) { icon.className = 'fas fa-check'; setTimeout(() => icon.className = 'fas fa-copy', 1500); }
    });
}

function dfcTestConnection(btn, resultId) {
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin" style="font-size:9px;margin-right:4px"></i>Testing…';
    const resultEl = document.getElementById(resultId);
    if (resultEl) resultEl.style.display = 'none';
    setTimeout(() => {
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-exchange-alt" style="font-size:9px;margin-right:4px"></i>Test Connection';
        if (resultEl) resultEl.style.display = 'inline-flex';
    }, 1200);
}

function dfcSyncFreqChange(value) {
    const dayEl = document.getElementById('sync-day');
    if (dayEl) dayEl.style.display = value === 'weekly' ? '' : 'none';
}

function dfcRosterProvider(value) {
    const classLinkEl = document.getElementById('roster-classlink');
    const googleEl    = document.getElementById('roster-google-classroom');
    if (classLinkEl) classLinkEl.style.display = value === 'classlink' ? '' : 'none';
    if (googleEl)    googleEl.style.display    = value === 'google-classroom' ? '' : 'none';
}

function dfcSubmitEarlyAccess(formId, successId) {
    const form    = document.getElementById(formId);
    const success = document.getElementById(successId);
    if (form)    form.style.display    = 'none';
    if (success) success.style.display = 'block';
}

// Initialise masked secret values on page load
document.addEventListener('DOMContentLoaded', function() {
    const secrets = [
        { id: 'entra-secret',  masked: '••••••••••••CAZZ' },
        { id: 'roster-secret', masked: '••••••••••••9bee' }
    ];
    secrets.forEach(s => {
        const el = document.getElementById(s.id);
        if (el) { el.setAttribute('data-real', el.value); el.value = s.masked; }
    });
});

// ── Policy interactivity (Screens 4 & 5) ──────────────────────────────────────
function onTeacherAuthChange() {
    const mode  = document.querySelector('input[name="teacherAuthMode"]:checked')?.value;
    const isSso = mode === 'sso_required';
    document.getElementById('sso-info-note').style.display       = isSso ? '' : 'none';
    document.getElementById('teacher-rostering-group').style.display = isSso ? '' : 'none';
}

function onStudentAuthChange() {
    // Screen 5 now uses polStudentAuthMode() — this is kept for legacy compatibility
}

// ── Summary + Start Over ───────────────────────────────────────────────────────
function updateSummary() {
    if (state.teacher) {
        const cls = state.teacher.selectedClass;
        const el  = document.getElementById('summary-teacher-text');
        if (el) el.innerHTML =
            `Teacher: <strong>${esc(state.teacher.email)}</strong>` +
            (cls ? ` — class: <strong>${esc(cls.title)}</strong>` : '');
    }
    if (state.student) {
        const cls = state.student.enrolledClass;
        const el  = document.getElementById('summary-student-text');
        if (el) el.innerHTML =
            `Student: <strong>${esc(state.student.email)}</strong>` +
            (cls ? ` — enrolled in <strong>${esc(cls.title)}</strong>` : ' — no class found');
    }
}

function startOver() {
    state.teacher = null; state.student = null; state.sync = null;

    // Reset Screen 1 scenario switcher
    activeSimScenario = 1; activeSimStep = 1; simClassChosen = false; sim8ClassIdVisible = false;
    document.querySelectorAll('#scenario-pills-row .scenario-pill').forEach((p, i) => p.classList.toggle('active', i === 0));
    const card1 = document.getElementById('scenario-card');
    if (card1) card1.textContent = SIM_SCENARIOS[0].card;
    renderSim(false);

    // Reset Screen 2 scenario switcher
    activeSimScenarioS2 = 1; activeSimStepS2 = 1; simS2ClassIdVisible = false;
    document.querySelectorAll('#s2-pills-row .scenario-pill').forEach((p, i) => p.classList.toggle('active', i === 0));
    const card2 = document.getElementById('s2-scenario-card');
    if (card2) card2.textContent = SIM_SCENARIOS_S2[0].card;
    renderSimS2(false);

    var lst = document.getElementById('last-sync-time');
    var ssb = document.getElementById('sync-status-badge');
    var sc  = document.getElementById('sync-counts');
    if (lst) lst.textContent = '—';
    if (ssb) ssb.style.display = 'none';
    if (sc)  sc.style.display  = 'none';
    var sciField = document.getElementById('s-class-id');
    if (sciField) sciField.value = '';
    showScreen(1);
}

// ── Helpers ────────────────────────────────────────────────────────────────────
function esc(s) {
    if (!s) return '';
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function setStatus(el, type, html) {
    const map = {
        'info':         'alert alert-info py-2 px-3 small mb-0',
        'success':      'alert alert-success py-2 px-3 small mb-0',
        'danger':       'alert alert-danger py-2 px-3 small mb-0',
        'info-light':   'small text-info',
        'success-light':'small text-success',
        'danger-light': 'small text-danger'
    };
    el.innerHTML = `<div class="${map[type] || 'small'}">${html}</div>`;
}

function classifyError(err) {
    const msg = (err.message || '').toLowerCase();
    if (msg.includes('popup') && msg.includes('close')) return 'Sign-in window was closed. Click to try again.';
    if (msg.includes('cancelled') || msg.includes('user_cancelled')) return 'Sign-in was cancelled.';
    if (msg.includes('no classlink mapping')) return err.message;
    if (msg.includes('network') || msg.includes('fetch')) return 'Network error. Check your connection.';
    return 'Sign-in failed: ' + esc(err.message || 'unknown error');
}

// ── Init ───────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    showScreen(1);
    initSimSwitcher();
    initSimSwitcherS2();
});

// ══════════════════════════════════════════════════════════════════════════════
// SCENARIO SWITCHER — Screen 1 only
// activeSimScenario: 1–8   |   activeSimStep: 1–3
// ══════════════════════════════════════════════════════════════════════════════

let activeSimScenario = 1;
let activeSimStep     = 1;
let simClassChosen    = false;
let sim8ClassIdVisible = false;

// ── Screen 2 sim state ────────────────────────────────────────────────────────
let activeSimScenarioS2 = 1;
let activeSimStepS2     = 1;
let simS2ClassIdVisible = false;

const SIM_SCENARIOS = [
    {
        pill: '★ Class roster + sign-in',
        card: 'SSO Required with ClassLink: Sign in with Microsoft and your class list loads automatically — no Class ID needed.',
        steps: 3
    },
    {
        pill: 'Sign-in only',
        card: "SSO Required without ClassLink: Microsoft sign-in is required, but you'll still type your Class ID after signing in.",
        steps: 2
    },
    {
        pill: 'Class ID only',
        card: "Standard mode: No Microsoft sign-in. Enter your Class ID and optional password as usual.",
        steps: 1
    }
];

function initSimSwitcher() {
    const row = document.getElementById('scenario-pills-row');
    if (!row) return;
    row.innerHTML = SIM_SCENARIOS.map((s, i) =>
        `<button class="scenario-pill${i === 0 ? ' active' : ''}" onclick="switchSim(${i + 1})">${s.pill}</button>`
    ).join('');
    const card = document.getElementById('scenario-card');
    if (card) card.textContent = SIM_SCENARIOS[0].card;
    renderSim(false);
}

function switchSim(n) {
    activeSimScenario  = n;
    activeSimStep      = 1;
    simClassChosen     = false;
    sim8ClassIdVisible = false;

    // Pills
    document.querySelectorAll('.scenario-pill').forEach((p, i) =>
        p.classList.toggle('active', i + 1 === n));

    // Description card fade
    const card = document.getElementById('scenario-card');
    if (card) {
        card.classList.add('fading');
        setTimeout(() => {
            card.textContent = SIM_SCENARIOS[n - 1].card;
            card.classList.remove('fading');
        }, 100);
    }

    renderSim(true);
}

function simAdvanceStep() {
    simClassChosen = false;
    const maxSteps = SIM_SCENARIOS[activeSimScenario - 1].steps;
    if (activeSimStep < maxSteps) activeSimStep++;
    renderSim(true);
}

function simGoFallback() {
    simClassChosen = false;
    activeSimStep  = 3;
    renderSim(true);
}

function simRefresh() {
    const sel = document.getElementById('sim-class-sel');
    if (!sel) return;
    sel.style.opacity = '.35';
    setTimeout(() => { sel.style.opacity = ''; }, 700);
}

function openClassListModal() {
    const m = document.getElementById('class-list-modal');
    if (m) { m.style.display = 'flex'; }
}
function closeClassListModal() {
    const m = document.getElementById('class-list-modal');
    if (m) { m.style.display = 'none'; }
}

function simClassChange() {
    const sel = document.getElementById('sim-class-sel');
    if (sel && sel.value === 'class-list') {
        sel.value = '';
        openClassListModal();
        return;
    }
    simClassChosen = !!(sel && sel.value !== '');
    const btn = document.getElementById('sim-start');
    if (btn) btn.disabled = !simClassChosen;
}

function simDynamicStart(inputId, btnId) {
    const inp = document.getElementById(inputId);
    const btn = document.getElementById(btnId);
    if (inp && btn) btn.disabled = !inp.value.trim();
}

function sim8RevealClassId() {
    sim8ClassIdVisible = true;
    renderSim(false);
}

function renderSim(animate) {
    const zone = document.getElementById('tc-sim-zone');
    if (!zone) return;
    const doRender = () => {
        zone.innerHTML = buildSimHtml();
        updateSimDots();
        zone.classList.remove('transitioning');
    };
    if (animate) {
        zone.classList.add('transitioning');
        setTimeout(doRender, 100);
    } else {
        doRender();
    }
}

function updateSimDots() {
    const dots = document.getElementById('tc-step-dots');
    if (!dots) return;
    const total = SIM_SCENARIOS[activeSimScenario - 1].steps;
    if (total <= 1) { dots.style.display = 'none'; return; }
    dots.style.display = 'flex';
    dots.innerHTML = Array.from({ length: total }, (_, i) =>
        `<span class="tc-dot${i + 1 === activeSimStep ? ' active' : ''}"></span>`
    ).join('');
}

// ── HTML fragment helpers ──────────────────────────────────────────────────────

const _MS_SVG = `<svg width="14" height="14" viewBox="0 0 21 21" style="flex-shrink:0"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>`;

// Real sign-in button (Scenario 1 only — no "simulated" label)
function _s1Step1Html() {
    return `<button class="sim-ms-btn" onclick="teacherSignIn()">${_MS_SVG} Sign in with Microsoft</button>
            <div class="sim-grey-text">Your school requires Microsoft sign-in</div>`;
}

// Simulated sign-in button (all other scenarios)
function _msBtn(label, onclick) {
    return `<button class="sim-ms-btn" onclick="${onclick}">${_MS_SVG} ${esc(label)}</button>`;
}

// Class dropdown using real ClassLink data (Scenario 1 step 2 after real sign-in)
function _realClassDropdown(classes) {
    let opts = '<option value="">— Choose a class —</option>';
    classes.forEach((cls, i) => {
        const cc = cls.courseCode ? ` (${cls.courseCode})` : '';
        opts += `<option value="${i}">⟳ ${esc(cls.title)}${esc(cc)} — ${cls.studentCount} student${cls.studentCount !== 1 ? 's' : ''}</option>`;
    });
    opts += `<option value="class-list" style="color:#2E78C1">+ Add students using Class List</option>`;
    return `
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
          <span class="sim-label" style="margin:0">Select your class:</span>
          <button class="sim-link" onclick="simRefresh()">⟳ Refresh list</button>
        </div>
        <select id="sim-class-sel" class="sim-select" onchange="simClassChange()">${opts}</select>
        <div style="text-align:center;margin:6px 0">
          <button class="sim-link" style="color:#d97706" onclick="simGoFallback()">⚠ What if ClassLink is down?</button>
        </div>
        <div style="text-align:right;margin-top:10px">
          <button class="sim-start-btn" id="sim-start" ${simClassChosen ? '' : 'disabled'}>Start Class</button>
        </div>`;
}

function _signedInPill() {
    return `<div class="sim-signed-pill"><span style="color:#16a34a">●</span> Signed in as teacher1@faronicsna.onmicrosoft.com</div>`;
}

function _classDropdown(showFallback) {
    const sel = `<select id="sim-class-sel" class="sim-select" onchange="simClassChange()">
      <option value="">— Choose a class —</option>
      <option>Science Lab · Period 2 · 14 students</option>
      <option>Mathematics 101 · Period 4 · 22 students</option>
      <option>English Literature · Period 6 · 18 students</option>
      <option value="class-list" style="color:#2E78C1">+ Add students using Class List</option>
    </select>`;
    const bar = `<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
      <span class="sim-label" style="margin:0">Select your class:</span>
      <button class="sim-link" onclick="simRefresh()">⟳ Refresh list</button>
    </div>`;
    const fallback = showFallback
        ? `<div style="text-align:center;margin:6px 0">
             <button class="sim-link" style="color:#d97706" onclick="simGoFallback()">⚠ What if ClassLink is down?</button>
           </div>`
        : '';
    return `${bar}${sel}${fallback}
      <div style="text-align:right;margin-top:10px">
        <button class="sim-start-btn" id="sim-start" ${simClassChosen ? '' : 'disabled'}>Start Class</button>
      </div>`;
}

function _classIdEntry(label, inputId, btnId) {
    return `<label class="sim-label">${esc(label)}</label>
      <input id="${inputId}" type="text" class="sim-input" placeholder="e.g. Science101"
             oninput="simDynamicStart('${inputId}','${btnId}')">
      <div style="text-align:right">
        <button class="sim-start-btn" id="${btnId}" disabled>Start Class</button>
      </div>`;
}

function _orDivider() {
    return `<div class="sim-or">— or —</div>`;
}

function _msOrClassId(msOnclick, inputId, btnId) {
    return `${_msBtn('Sign in with Microsoft', msOnclick)}
      ${_orDivider()}
      ${_classIdEntry('Enter Class ID', inputId, btnId)}`;
}

// ── Main scenario HTML builder ─────────────────────────────────────────────────

function buildSimHtml() {
    const s    = activeSimScenario;
    const step = activeSimStep;

    // ── S1: SSO Required + Rostering (REAL auth) ────────────────────────────
    if (s === 1) {
        if (step === 1) return _s1Step1Html();
        if (step === 2) {
            const pill = _signedInPill();
            const dropdown = (state.teacher && state.teacher.classes && state.teacher.classes.length > 0)
                ? _realClassDropdown(state.teacher.classes)
                : _classDropdown(true);
            return pill + dropdown;
        }
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-warning py-2 px-3 small mb-3">⚠ Couldn't load your class list right now. Enter your Class ID to continue.</div>` +
            _classIdEntry('Enter your Class ID to continue', 'sim-cid-1', 'sim-sb-1');
    }

    // ── S2: SSO Required, no Rostering ──────────────────────────────────────
    if (s === 2) {
        if (step === 1) return _msBtn('Sign in with Microsoft', 'simAdvanceStep()') +
            `<div class="sim-grey-text">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return _signedInPill() +
            _classIdEntry('Enter your Class ID to begin', 'sim-cid-2', 'sim-sb-2');
    }

    // ── S3: Standard — Class ID only ────────────────────────────────────────
    if (s === 3) {
        return `<label class="sim-label">Enter your Class ID</label>
          <input id="sim-cid-3" type="text" class="sim-input" placeholder="e.g. Science101"
                 oninput="simDynamicStart('sim-cid-3','sim-sb-3')">
          <label class="sim-label">Password <span style="font-weight:400;color:#9ca3af">(optional)</span></label>
          <input type="password" class="sim-input" placeholder="Password">
          <div style="text-align:right">
            <button class="sim-start-btn" id="sim-sb-3" disabled>Start Class</button>
          </div>`;
    }

    return '';
}

// ══════════════════════════════════════════════════════════════════════════════
// SCENARIO SWITCHER — Screen 2 (Student Client)
// activeSimScenarioS2: 1–5  |  activeSimStepS2: 1–3
// ══════════════════════════════════════════════════════════════════════════════

const SIM_SCENARIOS_S2 = [
    {
        pill: '★ Class roster + sign-in',
        card: 'Your school requires Microsoft sign-in. Once signed in, your class is assigned automatically from the school roster — no Class ID needed.',
        steps: 2
    },
    {
        pill: 'Sign-in only',
        card: "Microsoft sign-in is required, but class lookup is off. After signing in, you still type your Class ID to join.",
        steps: 2
    },
    {
        pill: 'Class ID only',
        card: "No Microsoft sign-in. Enter your Class ID and optional password — this is how Insight works today without SSO.",
        steps: 1
    },
    {
        pill: 'What if sign-in fails?',
        card: "See what the student experiences if Microsoft sign-in fails — and whether a Class ID fallback is available.",
        steps: 3
    }
];

function initSimSwitcherS2() {
    const row = document.getElementById('s2-pills-row');
    if (!row) return;
    row.innerHTML = SIM_SCENARIOS_S2.map((s, i) =>
        `<button class="scenario-pill${i === 0 ? ' active' : ''}" onclick="switchSimS2(${i + 1})">${s.pill}</button>`
    ).join('');
    const card = document.getElementById('s2-scenario-card');
    if (card) card.textContent = SIM_SCENARIOS_S2[0].card;
    renderSimS2(false);
}

function switchSimS2(n) {
    activeSimScenarioS2 = n;
    activeSimStepS2     = 1;
    simS2ClassIdVisible = false;

    document.querySelectorAll('#s2-pills-row .scenario-pill').forEach((p, i) =>
        p.classList.toggle('active', i + 1 === n));

    const card = document.getElementById('s2-scenario-card');
    if (card) {
        card.classList.add('fading');
        setTimeout(() => {
            card.textContent = SIM_SCENARIOS_S2[n - 1].card;
            card.classList.remove('fading');
        }, 100);
    }
    renderSimS2(true);
}

function simAdvanceStepS2() {
    simS2ClassIdVisible = false;
    const maxSteps = SIM_SCENARIOS_S2[activeSimScenarioS2 - 1].steps;
    if (activeSimStepS2 < maxSteps) activeSimStepS2++;
    renderSimS2(true);
}

function simS2RevealClassId() {
    simS2ClassIdVisible = true;
    renderSimS2(false);
}

function renderSimS2(animate) {
    const zone = document.getElementById('s2-sim-zone');
    if (!zone) return;
    const doRender = () => {
        zone.innerHTML = buildSimS2Html();
        updateSimDotsS2();
        zone.classList.remove('transitioning');
    };
    if (animate) {
        zone.classList.add('transitioning');
        setTimeout(doRender, 100);
    } else {
        doRender();
    }
}

function updateSimDotsS2() {
    const dots = document.getElementById('s2-step-dots');
    if (!dots) return;
    const total = SIM_SCENARIOS_S2[activeSimScenarioS2 - 1].steps;
    if (total <= 1) { dots.style.display = 'none'; return; }
    dots.style.display = 'flex';
    dots.innerHTML = Array.from({ length: total }, (_, i) =>
        `<span class="tc-dot${i + 1 === activeSimStepS2 ? ' active' : ''}"></span>`
    ).join('');
}

// ── S2 HTML fragment helpers ───────────────────────────────────────────────────

const _MS_SVG_16 = `<svg width="16" height="16" viewBox="0 0 21 21" style="flex-shrink:0"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>`;

// Real sign-in button (Scenario 1 only)
function _s2Step1Html() {
    return `<button class="sc-ms-btn" onclick="studentSignIn()">${_MS_SVG_16} Sign in with Microsoft</button>
            <div class="sim-grey-text" style="margin-top:8px">Your school requires Microsoft sign-in</div>`;
}

// Simulated orange MS button
function _s2MsBtn(onclick) {
    return `<button class="sc-ms-btn" onclick="${onclick}">${_MS_SVG_16} Sign in with Microsoft</button>`;
}

// Class ID form styled to match student client
function _s2ClassIdForm(inputId, btnId, label) {
    return `<label class="sim-label">${esc(label || 'Enter your Class ID')}</label>
            <input id="${inputId}" type="text" class="sc-classid-input" placeholder="e.g. Science101"
                   oninput="simDynamicStart('${inputId}','${btnId}')">
            <div style="text-align:center">
              <button class="sc-join-btn" id="${btnId}" disabled>Join Class</button>
            </div>`;
}

// Signed-in pill (green, uses real email if available)
function _s2SignedInPill(email) {
    const display = email || (state.student && state.student.email) || 'student1@faronicsna.onmicrosoft.com';
    return `<div class="sim-signed-pill"><span style="color:#16a34a">●</span> Signed in as ${esc(display)}</div>`;
}

// Enrolled class result with GO TO CLASS button
function _s2ClassResult(title) {
    return `<div class="small mb-3" style="color:#333">You are enrolled in:<br><strong>${esc(title)}</strong></div>
            <button class="sc-join-btn">GO TO CLASS</button>`;
}

// ── Main S2 scenario HTML builder ─────────────────────────────────────────────

function buildSimS2Html() {
    const s    = activeSimScenarioS2;
    const step = activeSimStepS2;
    const DEMO_CLASS = 'Science Lab · Period 2';
    const DEMO_EMAIL = 'student1@faronicsna.onmicrosoft.com';

    // ── S1: Class roster + sign-in (REAL auth) ───────────────────────────────
    if (s === 1) {
        if (step === 1) return _s2Step1Html();
        if (step === 2) {
            const pill = _s2SignedInPill(state.student && state.student.email);
            if (state.student && state.student.enrolledClass) {
                return pill + _s2ClassResult(state.student.enrolledClass.title);
            } else if (state.student) {
                return pill + `<div class="small" style="color:#c0392b">No class assignment found.<br>Contact your teacher.</div>`;
            }
            return pill + _s2ClassResult(DEMO_CLASS);
        }
    }

    // ── S2: SSO Required + manual Class ID ──────────────────────────────────
    if (s === 2) {
        if (step === 1) return _s2MsBtn('simAdvanceStepS2()') +
            `<div class="sim-grey-text" style="margin-top:8px">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return _s2SignedInPill(DEMO_EMAIL) +
            _s2ClassIdForm('s2-cid-2', 's2-sb-2', 'Enter your Class ID to join');
    }

    // ── S3: Standard — Class ID only ────────────────────────────────────────
    if (s === 3) {
        return `<label class="sim-label">Enter your Class ID</label>
                <input id="s2-cid-3" type="text" class="sc-classid-input" placeholder="e.g. Science101"
                       oninput="simDynamicStart('s2-cid-3','s2-sb-3')">
                <label class="sim-label">Password <span style="font-weight:400;color:#9ca3af">(optional)</span></label>
                <input type="password" class="sc-classid-input" placeholder="Password">
                <div style="text-align:center">
                  <button class="sc-join-btn" id="s2-sb-3" disabled>Join Class</button>
                </div>`;
    }

    // ── S4: Sign-in failure ──────────────────────────────────────────────────
    if (s === 4) {
        if (step === 1) return _s2MsBtn('simAdvanceStepS2()') +
            `<div class="sim-grey-text" style="margin-top:8px">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return `
            <div class="alert alert-danger py-2 px-3 small mb-3">✕ Couldn't sign in with Microsoft. Check your internet connection and try again.</div>
            ${_s2MsBtn('simAdvanceStepS2()')}
            <div style="text-align:center;margin-top:8px">
              <button class="sim-link" onclick="simS2RevealClassId()">Enter Class ID instead →</button>
            </div>
            ${simS2ClassIdVisible ? `<div class="mt-3">${_s2ClassIdForm('s2-cid-4', 's2-sb-4', 'Enter your Class ID to continue')}</div>` : ''}`;
        if (step === 3) return _s2SignedInPill(DEMO_EMAIL) +
            `<div class="alert alert-success py-2 px-3 mb-3" style="font-size:11px">✓ Sign-in succeeded on retry.</div>` +
            _s2ClassResult(DEMO_CLASS);
    }

    return '';
}
