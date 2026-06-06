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

    if (n === 5) updateSummary();
    if (n === 3 && !state.sync) triggerSync();
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
    const btn      = document.getElementById('btn-student-signin');
    const statusEl = document.getElementById('student-signin-status');
    btn.disabled   = true;
    document.getElementById('student-class-section').style.display = 'none';

    try {
        setStatus(statusEl, 'info-light', 'Opening Microsoft sign-in…');
        const msalResult = await signIn();

        setStatus(statusEl, 'info-light', `Signed in as ${esc(msalResult.email)} — looking up enrolled class…`);
        const clData = await clApi('student', msalResult.email);

        state.student = {
            email:        msalResult.email,
            displayName:  msalResult.displayName,
            sourcedId:    clData.sourcedId,
            enrolledClass: clData.enrolledClass
        };

        setStatus(statusEl, 'success-light', `Signed in as <strong>${esc(msalResult.email)}</strong>`);

        const section   = document.getElementById('student-class-section');
        const classInfo = document.getElementById('student-class-info');
        const joinBtn   = document.getElementById('btn-student-join');

        if (clData.enrolledClass) {
            classInfo.innerHTML = `<div class="small text-white">You are enrolled in:<br><strong>${esc(clData.enrolledClass.title)}</strong></div>`;
            // Update Screen 5 student class ID field if it exists
            var sciField = document.getElementById('s-class-id');
            if (sciField) sciField.value = 'cl_' + clData.enrolledClass.sourcedId;
            joinBtn.disabled = false;
        } else {
            classInfo.innerHTML = `<div class="small text-warning">No class assignment found.<br>Contact your teacher.</div>`;
            joinBtn.disabled = true;
        }
        section.style.display = 'block';

        
        updateNav();
        updateSummary();
    } catch (err) {
        const msg = classifyError(err);
        setStatus(statusEl, 'danger-light', msg);
        btn.disabled = false;
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
    const standardFields = document.getElementById('t-standard-fields');
    const rosterCb       = document.getElementById('t-roster-cb');
    const rosterWarn     = document.getElementById('t-roster-sso-warn');

    // Show/hide Class ID field under Standard mode
    if (standardFields) standardFields.style.display = mode === 'standard' ? '' : 'none';

    // If switching back to Standard while rostering was on — uncheck + hide fallback options
    const isSso = mode !== 'standard';
    if (!isSso && rosterCb && rosterCb.checked) {
        rosterCb.checked = false;
        polTeacherRostering(rosterCb);
    }
    // Hide the SSO warning if mode changes away from Standard
    if (isSso && rosterWarn) rosterWarn.style.display = 'none';
}

// Legacy compat stubs
function polEnableTeacherSSO(cb) {}
function polTeacherAuthMode(mode) {}

function polTeacherRostering(cb) {
    const opts     = document.getElementById('t-roster-opts');
    const warn     = document.getElementById('t-roster-sso-warn');
    const mode     = document.querySelector('input[name="t-signin-mode"]:checked')?.value || 'standard';
    const isSso    = mode !== 'standard';

    if (cb.checked && !isSso) {
        // Tried to enable rostering while in Standard mode — block it and show warning
        cb.checked = false;
        if (warn) warn.style.display = 'block';
        if (opts) opts.style.display = 'none';
        return;
    }

    // Hide warning once rostering is allowed to proceed
    if (warn) warn.style.display = 'none';
    if (opts) opts.style.display = cb.checked ? 'block' : 'none';
}

// Save validation for teacher policy
function polSaveTeacher() {
    const mode       = document.querySelector('input[name="t-signin-mode"]:checked')?.value || 'standard';
    const rosterOn   = document.getElementById('t-roster-cb')?.checked || false;
    const errContainer = document.getElementById('t-save-errors');
    if (!errContainer) return;

    const errors = [];

    // Check SSO configured (use state.teacher as proxy — sign-in on Screen 1 proves SSO works)
    if (mode !== 'standard' && !state.teacher) {
        errors.push({
            icon: 'fab fa-microsoft',
            msg: 'Microsoft Entra ID SSO is not configured in Organization Settings.',
            link: true
        });
    }

    // Check ClassLink configured (state.sync proves a sync has run)
    if (rosterOn && !state.sync) {
        errors.push({
            icon: 'fas fa-sync',
            msg: 'ClassLink Rostering is not configured in Organization Settings.',
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

// Student SSO enable/disable
function polEnableStudentSSO(cb) {
    const fields = document.getElementById('s-sso-fields');
    if (fields) {
        fields.style.opacity = cb.checked ? '1' : '0.45';
        fields.style.pointerEvents = cb.checked ? '' : 'none';
    }
}

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
    const isSso = mode === 'sso_preferred' || mode === 'sso_required';
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
        document.getElementById('summary-teacher-text').innerHTML =
            `Teacher: <strong>${esc(state.teacher.email)}</strong>` +
            (cls ? ` — class: <strong>${esc(cls.title)}</strong>` : '');
    }
    if (state.student) {
        const cls = state.student.enrolledClass;
        document.getElementById('summary-student-text').innerHTML =
            `Student: <strong>${esc(state.student.email)}</strong>` +
            (cls ? ` — enrolled in <strong>${esc(cls.title)}</strong>` : ' — no class found');
    }
}

function startOver() {
    state.teacher = null; state.student = null; state.sync = null;
    
    ['teacher-signin-status','student-signin-status'].forEach(id => document.getElementById(id).innerHTML = '');
    // Reset scenario switcher + teacher sign-in state
    activeSimScenario = 1; activeSimStep = 1; simClassChosen = false; sim8ClassIdVisible = false;
    
    document.querySelectorAll('.scenario-pill').forEach((p, i) => p.classList.toggle('active', i === 0));
    const card = document.getElementById('scenario-card');
    if (card) card.textContent = SIM_SCENARIOS[0].card;
    renderSim(false);
    document.getElementById('teacher-class-section').style.display = 'none';
    document.getElementById('student-class-section').style.display = 'none';
    document.getElementById('btn-teacher-signin').disabled = false;
    document.getElementById('btn-student-signin').disabled = false;
    var lst = document.getElementById('last-sync-time');
    var ssb = document.getElementById('sync-status-badge');
    var sc  = document.getElementById('sync-counts');
    if (lst) lst.textContent = '—';
    if (ssb) ssb.style.display = 'none';
    if (sc)  sc.style.display  = 'none';
    var sciField = document.getElementById('s-class-id');
    if (sciField) sciField.value = '';
    document.getElementById('summary-teacher-text').innerHTML  = 'Teacher: <em>sign in on Screen 1 to populate</em>';
    document.getElementById('summary-student-text').innerHTML  = 'Student: <em>sign in on Screen 2 to populate</em>';
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
});

// ══════════════════════════════════════════════════════════════════════════════
// SCENARIO SWITCHER — Screen 1 only
// activeSimScenario: 1–8   |   activeSimStep: 1–3
// ══════════════════════════════════════════════════════════════════════════════

let activeSimScenario = 1;
let activeSimStep     = 1;
let simClassChosen    = false;
let sim8ClassIdVisible = false;

const SIM_SCENARIOS = [
    {
        pill: '★ Class roster + sign-in',
        card: 'Your IT admin has set up Microsoft sign-in. Sign in and your classes load automatically — no need to type a Class ID.',
        steps: 3
    },
    {
        pill: 'Class roster required',
        card: "Microsoft sign-in is required and your class list must load from ClassLink to start. If ClassLink is unavailable, you can't start class.",
        steps: 3
    },
    {
        pill: 'Sign-in only',
        card: "Microsoft sign-in is required, but you'll still type your Class ID after signing in.",
        steps: 2
    },
    {
        pill: 'Sign-in or Class ID',
        card: "You can sign in with Microsoft to get your class list, or skip it and type your Class ID as usual.",
        steps: 3
    },
    {
        pill: 'Sign-in or Class ID (strict)',
        card: "Same as above, but if ClassLink is unavailable after sign-in, you won't be able to start class.",
        steps: 3
    },
    {
        pill: 'Sign-in or Class ID (no roster)',
        card: "You can sign in with Microsoft, but your class list won't load automatically — you'll still type your Class ID.",
        steps: 2
    },
    {
        pill: 'Class ID only',
        card: "No Microsoft sign-in. Enter your Class ID and optional password as usual. This is how Insight works today.",
        steps: 1
    },
    {
        pill: 'What if sign-in fails?',
        card: "See what the teacher experiences if Microsoft sign-in fails — and what recovery options are available.",
        steps: 3
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

function simClassChange() {
    simClassChosen = true;
    const btn = document.getElementById('sim-start');
    if (btn) btn.disabled = false;
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
    return `<button class="sim-ms-btn" onclick="${onclick}">${_MS_SVG} ${esc(label)}</button>
            <div class="sim-simulated">(simulated — no real sign-in happens here)</div>`;
}

// Class dropdown using real ClassLink data (Scenario 1 step 2 after real sign-in)
function _realClassDropdown(classes) {
    let opts = '<option value="">— Choose a class —</option>';
    classes.forEach((cls, i) => {
        const cc = cls.courseCode ? ` (${cls.courseCode})` : '';
        opts += `<option value="${i}">⟳ ${esc(cls.title)}${esc(cc)} — ${cls.studentCount} student${cls.studentCount !== 1 ? 's' : ''}</option>`;
    });
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

    // ── S1: SSO Required + Rostering + Allow Manual  (REAL auth) ────────────
    if (s === 1) {
        if (step === 1) return _s1Step1Html();
        if (step === 2) {
            // Use live ClassLink data after real sign-in, else demo data
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

    // ── S2: SSO Required + Rostering + Block ────────────────────────────────
    if (s === 2) {
        if (step === 1) return _msBtn('Sign in with Microsoft', 'simAdvanceStep()') +
            `<div class="sim-grey-text">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return _signedInPill() + _classDropdown(true);
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-danger py-2 px-3 small mb-3">✕ Your class list couldn't be loaded and your school requires it to start class. Please contact your IT administrator.</div>
             <div style="text-align:right"><button class="sim-start-btn" disabled>Start Class</button></div>`;
    }

    // ── S3: SSO Required + No Rostering ─────────────────────────────────────
    if (s === 3) {
        if (step === 1) return _msBtn('Sign in with Microsoft', 'simAdvanceStep()') +
            `<div class="sim-grey-text">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return _signedInPill() +
            _classIdEntry('Enter your Class ID to begin', 'sim-cid-3', 'sim-sb-3');
    }

    // ── S4: SSO Preferred + Rostering + Allow Manual ────────────────────────
    if (s === 4) {
        if (step === 1) return _msOrClassId('simAdvanceStep()', 'sim-cid-4a', 'sim-sb-4a');
        if (step === 2) return _signedInPill() + _classDropdown(true);
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-warning py-2 px-3 small mb-3">⚠ Couldn't load your class list right now. Enter your Class ID to continue.</div>` +
            _classIdEntry('Enter your Class ID to continue', 'sim-cid-4b', 'sim-sb-4b');
    }

    // ── S5: SSO Preferred + Rostering + Block ───────────────────────────────
    if (s === 5) {
        if (step === 1) return _msOrClassId('simAdvanceStep()', 'sim-cid-5a', 'sim-sb-5a');
        if (step === 2) return _signedInPill() + _classDropdown(true);
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-danger py-2 px-3 small mb-3">✕ Your class list couldn't be loaded and your school requires it to start class. Please contact your IT administrator.</div>
             <div style="text-align:right"><button class="sim-start-btn" disabled>Start Class</button></div>`;
    }

    // ── S6: SSO Preferred + No Rostering ────────────────────────────────────
    if (s === 6) {
        if (step === 1) return _msOrClassId('simAdvanceStep()', 'sim-cid-6a', 'sim-sb-6a');
        if (step === 2) return _signedInPill() +
            _classIdEntry('Enter your Class ID to begin', 'sim-cid-6b', 'sim-sb-6b');
    }

    // ── S7: Standard — Class ID only ────────────────────────────────────────
    if (s === 7) {
        return `<label class="sim-label">Enter your Class ID</label>
          <input id="sim-cid-7" type="text" class="sim-input" placeholder="e.g. Science101"
                 oninput="simDynamicStart('sim-cid-7','sim-sb-7')">
          <label class="sim-label">Password <span style="font-weight:400;color:#9ca3af">(optional)</span></label>
          <input type="password" class="sim-input" placeholder="Password">
          <div style="text-align:right">
            <button class="sim-start-btn" id="sim-sb-7" disabled>Start Class</button>
          </div>`;
    }

    // ── S8: Sign-in failure ──────────────────────────────────────────────────
    if (s === 8) {
        if (step === 1) return _msBtn('Sign in with Microsoft', 'simAdvanceStep()') +
            `<div class="sim-grey-text">Your school requires Microsoft sign-in</div>`;
        if (step === 2) return `
            <div class="alert alert-danger py-2 px-3 small mb-3">✕ Couldn't sign in with Microsoft. Check your internet connection and try again.</div>
            ${_msBtn('Try again with Microsoft', 'simAdvanceStep()')}
            <div style="text-align:center;margin-top:8px">
              <button class="sim-link" onclick="sim8RevealClassId()">Use Class ID instead →</button>
            </div>
            ${sim8ClassIdVisible ? `<div class="mt-3">${_classIdEntry('Enter your Class ID to continue', 'sim-cid-8', 'sim-sb-8')}</div>` : ''}`;
        if (step === 3) return _signedInPill() +
            `<div class="alert alert-success py-2 px-3 mb-3" style="font-size:11px">✓ Sign-in succeeded on retry.</div>` +
            _classDropdown(false);
    }

    return '';
}
