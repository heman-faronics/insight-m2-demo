// INSIGHT M2 DEMO -- NOT PRODUCTION CODE
// api/google.js — Vercel serverless function
// Server-side Google Classroom rostering, mirroring api/classlink.js.
// Called by the browser via fetch('/api/google?op=...')
//
// TWO MODES
//   live      — real Google APIs. Active when the env vars below are present.
//   simulated — canned fixtures. Active when they are not.
// The response shape is identical in both modes, so the browser never branches
// on it; only the `mode` field differs. Dropping the credentials in flips it.
//
// REQUIRED ENV VARS FOR LIVE MODE (set in Vercel, never in source):
//   GOOGLE_SA_EMAIL              service account address
//   GOOGLE_SA_PRIVATE_KEY        service account private key (PEM, \n escaped)
//   GOOGLE_IMPERSONATE_SUBJECT   Workspace admin the sync job acts as
//   GOOGLE_WORKSPACE_DOMAIN      primary domain, e.g. school.example.org
//
// WHY DOMAIN-WIDE DELEGATION
// The Classroom API has no "list every course in the domain" call — courses.list
// returns only what the *authenticated* user can see, super admin included. So a
// full sync enumerates teachers through the Admin SDK Directory API and then
// impersonates each one. That is the reason admin.directory.user.readonly is in
// the scope list below and is the main structural difference from ClassLink's
// single-tenant OneRoster pull.

const SCOPES = [
    'https://www.googleapis.com/auth/classroom.courses.readonly',
    'https://www.googleapis.com/auth/classroom.rosters.readonly',
    'https://www.googleapis.com/auth/classroom.profile.emails',
    'https://www.googleapis.com/auth/admin.directory.user.readonly'
].join(' ');

const TOKEN_URL     = 'https://oauth2.googleapis.com/token';
const CLASSROOM_API = 'https://classroom.googleapis.com/v1';
const DIRECTORY_API = 'https://admin.googleapis.com/admin/directory/v1';

const SA_EMAIL   = process.env.GOOGLE_SA_EMAIL;
const SA_KEY     = (process.env.GOOGLE_SA_PRIVATE_KEY || '').replace(/\\n/g, '\n');
const SUBJECT    = process.env.GOOGLE_IMPERSONATE_SUBJECT;
const GW_DOMAIN  = process.env.GOOGLE_WORKSPACE_DOMAIN || 'demo.faronics.org';

const LIVE = Boolean(SA_EMAIL && SA_KEY && SUBJECT);

// ══════════════════════════════════════════════════════════════════════════════
// SIMULATED FIXTURES — shaped to mirror the ClassLink cert fixture (one teacher,
// two classes, three students each) so the two providers demo side by side.
// ══════════════════════════════════════════════════════════════════════════════

const SIM_TEACHER = {
    email:       `teacher1@${GW_DOMAIN}`,
    displayName: 'Marian Lucas',
    userId:      '114285730192847561093'
};

const SIM_COURSES = [
    {
        id:           '652841190377',
        name:         'Grade 6 Science',
        section:      'Period 2',
        courseState:  'ACTIVE',
        enrollmentCode: 'qr7x2mv',
        students: [
            { userId: '108372910284756102938', email: `student1@${GW_DOMAIN}`, name: 'Ava Nguyen' },
            { userId: '117482910384756102947', email: `student2@${GW_DOMAIN}`, name: 'Liam Patel' },
            { userId: '102938471029384710293', email: `student3@${GW_DOMAIN}`, name: 'Noah Kim'  }
        ]
    },
    {
        id:           '652841190412',
        name:         'Grade 6 Mathematics',
        section:      'Period 4',
        courseState:  'ACTIVE',
        enrollmentCode: 'bk4n8dz',
        students: [
            { userId: '108372910284756102938', email: `student1@${GW_DOMAIN}`, name: 'Ava Nguyen'    },
            { userId: '129384710293847102938', email: `student4@${GW_DOMAIN}`, name: 'Mia Okonkwo'   },
            { userId: '138475610293847561029', email: `student5@${GW_DOMAIN}`, name: 'Marcus Reyes'  }
        ]
    }
];

// Cross-provider identity map. An Entra-authenticated teacher and a Google
// Classroom course owner will not share an email address, so a pairing that
// mixes providers needs an explicit map — the same manual_map strategy
// api/classlink.js uses for Entra → ClassLink.
const IDENTITY_MAP = {
    'teacher1@faronicsna.onmicrosoft.com': SIM_TEACHER.email,
    'student1@faronicsna.onmicrosoft.com': `student1@${GW_DOMAIN}`,
    'student2@faronicsna.onmicrosoft.com': `student2@${GW_DOMAIN}`
};

// ══════════════════════════════════════════════════════════════════════════════
// LIVE MODE — service account JWT → access token → Classroom / Directory
// ══════════════════════════════════════════════════════════════════════════════

const crypto = require('crypto');

function b64url(input) {
    return Buffer.from(input).toString('base64')
        .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// Signed JWT assertion carrying the impersonation subject (RFC 7523)
function buildAssertion(subject) {
    const now    = Math.floor(Date.now() / 1000);
    const header = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
    const claims = b64url(JSON.stringify({
        iss:   SA_EMAIL,
        sub:   subject,          // the user being impersonated
        scope: SCOPES,
        aud:   TOKEN_URL,
        iat:   now,
        exp:   now + 3600
    }));
    const signer = crypto.createSign('RSA-SHA256');
    signer.update(`${header}.${claims}`);
    return `${header}.${claims}.${b64url(signer.sign(SA_KEY))}`;
}

const tokenCache = new Map();   // subject → { token, expiresAt }

async function getToken(subject) {
    const hit = tokenCache.get(subject);
    if (hit && hit.expiresAt > Date.now() + 60000) return hit.token;

    const res = await fetch(TOKEN_URL, {
        method:  'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body:    new URLSearchParams({
            grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
            assertion:  buildAssertion(subject)
        })
    });
    if (!res.ok) {
        const body = await res.text();
        throw new Error(`Google token failed for ${subject}: HTTP ${res.status} ${body.slice(0, 200)}`);
    }
    const data = await res.json();
    tokenCache.set(subject, {
        token:     data.access_token,
        expiresAt: Date.now() + (data.expires_in || 3600) * 1000
    });
    return data.access_token;
}

async function gGet(base, path, subject, params) {
    const token = await getToken(subject);
    const url   = new URL(base + path);
    Object.entries(params || {}).forEach(([k, v]) => v != null && url.searchParams.set(k, v));
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
        const body = await res.text();
        throw new Error(`Google GET ${path} failed: HTTP ${res.status} ${body.slice(0, 200)}`);
    }
    return res.json();
}

// Courses the given teacher owns or co-teaches
async function liveCoursesForTeacher(teacherEmail) {
    const data = await gGet(CLASSROOM_API, '/courses', teacherEmail, {
        teacherId:   'me',
        courseStates: 'ACTIVE',
        pageSize:    100
    });
    const courses = data.courses || [];
    return Promise.all(courses.map(async c => {
        let students = [];
        try {
            const sd = await gGet(CLASSROOM_API, `/courses/${c.id}/students`, teacherEmail, { pageSize: 200 });
            students = (sd.students || []).map(s => ({
                userId: s.userId,
                email:  s.profile && s.profile.emailAddress,
                name:   s.profile && s.profile.name && s.profile.name.fullName
            }));
        } catch (_) { /* teacher may lack roster scope on a co-taught course */ }
        return {
            id:             c.id,
            name:           c.name,
            section:        c.section || '',
            courseState:    c.courseState,
            enrollmentCode: c.enrollmentCode,
            students
        };
    }));
}

// Every teacher in the domain, via the Admin SDK — see the DWD note at the top
async function liveTeachers() {
    const out = [];
    let pageToken;
    do {
        const data = await gGet(DIRECTORY_API, '/users', SUBJECT, {
            domain:     GW_DOMAIN,
            maxResults: 200,
            projection: 'basic',
            pageToken
        });
        (data.users || []).forEach(u => {
            if (!u.suspended) out.push({ email: u.primaryEmail, displayName: u.name && u.name.fullName });
        });
        pageToken = data.nextPageToken;
    } while (pageToken);
    return out;
}

// ══════════════════════════════════════════════════════════════════════════════
// Shared helpers
// ══════════════════════════════════════════════════════════════════════════════

// Resolve an authenticated identity to a Google Workspace address. Same-domain
// sign-ins pass straight through; a cross-provider sign-in (Entra SSO paired
// with Google Classroom rostering) falls back to the manual map.
function resolveGoogleEmail(email) {
    const lower = (email || '').toLowerCase();
    if (!lower) return null;
    if (lower.endsWith(`@${GW_DOMAIN}`)) return lower;
    return IDENTITY_MAP[lower] || null;
}

function toClassShape(course) {
    return {
        sourcedId:    course.id,
        title:        course.section ? `${course.name} · ${course.section}` : course.name,
        courseCode:   course.enrollmentCode || '',
        studentCount: (course.students || []).length
    };
}

// ══════════════════════════════════════════════════════════════════════════════
// Handler
// ══════════════════════════════════════════════════════════════════════════════

module.exports = async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    if (req.method === 'OPTIONS') { res.status(200).end(); return; }

    const { op, email } = req.query;
    const mode = LIVE ? 'live' : 'simulated';

    try {
        // ── op=config: what the admin UI needs to render its state ────────────
        if (op === 'config') {
            return res.status(200).json({
                mode,
                domain:       GW_DOMAIN,
                serviceAccount: LIVE ? SA_EMAIL : null,
                subject:      LIVE ? SUBJECT : null,
                scopes:       SCOPES.split(' ')
            });
        }

        // ── op=teacher: a teacher's Classroom courses ─────────────────────────
        if (op === 'teacher') {
            if (!email) return res.status(400).json({ error: 'email required' });
            const resolved = resolveGoogleEmail(email);
            if (!resolved) {
                return res.status(404).json({
                    error: `No Google Workspace identity mapped for ${email}`,
                    code:  'IDENTITY_UNMAPPED'
                });
            }

            const courses = LIVE
                ? await liveCoursesForTeacher(resolved)
                : SIM_COURSES;

            return res.status(200).json({
                mode,
                provider:    'google_classroom',
                sourcedId:   resolved,
                displayName: LIVE ? resolved : SIM_TEACHER.displayName,
                classes:     courses.map(toClassShape)
            });
        }

        // ── op=student: the classes a student is enrolled in ──────────────────
        if (op === 'student') {
            if (!email) return res.status(400).json({ error: 'email required' });
            const resolved = resolveGoogleEmail(email);
            if (!resolved) {
                return res.status(404).json({
                    error: `No Google Workspace identity mapped for ${email}`,
                    code:  'IDENTITY_UNMAPPED'
                });
            }

            let courses;
            if (LIVE) {
                // A student can list their own courses once impersonated
                const data = await gGet(CLASSROOM_API, '/courses', resolved, {
                    studentId:    'me',
                    courseStates: 'ACTIVE',
                    pageSize:     100
                });
                courses = (data.courses || []).map(c => ({ ...c, students: [] }));
            } else {
                courses = SIM_COURSES.filter(c => c.students.some(s => s.email === resolved));
            }

            const classes = courses.map(toClassShape);
            return res.status(200).json({
                mode,
                provider:      'google_classroom',
                sourcedId:     resolved,
                enrolledClass: classes[0] || null,
                classes                                   // all enrolments, for conflict handling
            });
        }

        // ── op=sync: the nightly server-side pull the admin screen triggers ───
        if (op === 'sync') {
            let teacherCount, classCount, studentCount;

            if (LIVE) {
                const teachers = await liveTeachers();
                const seenStudents = new Set();
                let classes = 0, owning = 0;
                for (const t of teachers) {
                    let courses = [];
                    try { courses = await liveCoursesForTeacher(t.email); } catch (_) { continue; }
                    if (!courses.length) continue;
                    owning++;
                    classes += courses.length;
                    courses.forEach(c => (c.students || []).forEach(s => s.userId && seenStudents.add(s.userId)));
                }
                teacherCount = owning;
                classCount   = classes;
                studentCount = seenStudents.size;
            } else {
                teacherCount = 1;
                classCount   = SIM_COURSES.length;
                studentCount = new Set(SIM_COURSES.flatMap(c => c.students.map(s => s.userId))).size;
            }

            return res.status(200).json({
                mode,
                provider: 'google_classroom',
                syncedAt: new Date().toISOString().replace('T', ' ').slice(0, 19) + ' UTC',
                teacherCount,
                classCount,
                studentCount
            });
        }

        // ── op=test: the admin "Test Connection" button ───────────────────────
        if (op === 'test') {
            if (!LIVE) {
                return res.status(200).json({
                    mode,
                    ok:      true,
                    detail:  'Simulated — no Google credentials configured on this deployment.',
                    domain:  GW_DOMAIN
                });
            }
            await getToken(SUBJECT);
            const probe = await gGet(DIRECTORY_API, '/users', SUBJECT, { domain: GW_DOMAIN, maxResults: 1 });
            return res.status(200).json({
                mode,
                ok:     true,
                detail: `Domain-wide delegation verified for ${SUBJECT}.`,
                domain: GW_DOMAIN,
                probe:  (probe.users || []).length
            });
        }

        return res.status(400).json({ error: `Unknown op: ${op}` });

    } catch (err) {
        return res.status(502).json({ error: err.message, mode });
    }
};
