const ExcelJS = require('exceljs');
const sgMail = require('@sendgrid/mail');

// ---- Config (all secrets come from GitHub Actions secrets, passed in as env vars) ----
const JOTFORM_API_KEY = process.env.JOTFORM_API_KEY;
const JOTFORM_FORM_ID = '260626585194162'; // "State of Values Report"
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const SENDGRID_API_KEY = process.env.SENDGRID_API_KEY;
const FROM_EMAIL = process.env.FROM_EMAIL; // must be a verified SendGrid sender
const TO_EMAIL = process.env.TO_EMAIL || 'asegal@think2perform.com';

async function main() {
  const submissions = await getRecentSubmissions();

  if (submissions.length === 0) {
    await sendEmail(null, 0);
    console.log('No new submissions in the last 24 hours. Notification sent.');
    return;
  }

  const leads = [];
  for (const sub of submissions) {
    const { name, email } = extractNameEmail(sub);
    if (!email) continue;
    const research = await researchLead(name, email);
    leads.push({ name, email, ...research });
  }

  const excelBuffer = await buildExcel(leads);
  await sendEmail(excelBuffer, leads.length);

  console.log(`Processed and emailed ${leads.length} lead(s).`);
}

// ---- Step 1: Pull yesterday's submissions from Jotform ----
async function getRecentSubmissions() {
  const since = new Date(Date.now() - 24 * 60 * 60 * 1000);
  const sinceStr = since.toISOString().slice(0, 19).replace('T', ' ');

  const filter = encodeURIComponent(JSON.stringify({ 'created_at:gt': sinceStr }));
  const url = `https://api.jotform.com/form/${JOTFORM_FORM_ID}/submissions?apiKey=${JOTFORM_API_KEY}&limit=1000&filter=${filter}&orderby=created_at`;

  const resp = await fetch(url);
  if (!resp.ok) {
    throw new Error(`Jotform API error: ${resp.status} ${await resp.text()}`);
  }
  const data = await resp.json();
  return data.content || [];
}

// ---- Step 2: Pull the name + email out of a submission's answers ----
function extractNameEmail(submission) {
  let name = '';
  let email = '';
  const answers = submission.answers || {};

  for (const key in answers) {
    const field = answers[key];
    const type = (field.type || '').toLowerCase();
    const label = (field.text || '').toLowerCase();

    if (type.includes('email') || label.includes('e-mail') || label.includes('email')) {
      if (typeof field.answer === 'string') email = field.answer.trim();
    }

    if (type.includes('name') || label.includes('name')) {
      if (field.answer && typeof field.answer === 'object') {
        const first = field.answer.first || '';
        const last = field.answer.last || '';
        name = `${first} ${last}`.trim();
      } else if (typeof field.answer === 'string') {
        name = field.answer.trim();
      }
    }
  }

  return { name, email };
}

// ---- Step 3: Research the lead using the Anthropic API with web search ----
async function researchLead(name, email) {
  const domain = (email.split('@')[1] || '').trim();

  const prompt = `Research this lead using web search.
Name: "${name || 'Unknown'}"
Email: "${email}"
Email domain: "${domain}"

Figure out:
1. company - the company they work for (use the email domain as the primary clue; confirm/refine with search)
2. industry - the industry that company is in
3. employee_count - an approximate employee count for that company (e.g. "50-200", "1000+", or a specific number if you find one)
4. title - their job title, if you can find it

If a field can't be confidently determined, use "Unknown" for that field. Do not guess wildly.

Respond with ONLY a JSON object and nothing else - no markdown code fences, no preamble, no explanation. Exact shape:
{"company": "...", "industry": "...", "employee_count": "...", "title": "..."}`;

  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 1024,
      messages: [{ role: 'user', content: prompt }],
      tools: [{ type: 'web_search_20250305', name: 'web_search' }]
    })
  });

  if (!resp.ok) {
    console.error('Anthropic API error:', resp.status, await resp.text());
    return { company: 'Unknown', industry: 'Unknown', employee_count: 'Unknown', title: 'Unknown' };
  }

  const data = await resp.json();
  const textBlocks = (data.content || [])
    .filter((b) => b.type === 'text')
    .map((b) => b.text)
    .join('\n');

  const jsonMatch = textBlocks.match(/\{[\s\S]*\}/);
  let parsed = { company: 'Unknown', industry: 'Unknown', employee_count: 'Unknown', title: 'Unknown' };
  if (jsonMatch) {
    try {
      parsed = { ...parsed, ...JSON.parse(jsonMatch[0]) };
    } catch (e) {
      console.error('Failed to parse research JSON for', email, e);
    }
  }
  return parsed;
}

// ---- Step 4: Build the Excel workbook ----
async function buildExcel(leads) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Leads');

  sheet.columns = [
    { header: 'Name', key: 'name', width: 25 },
    { header: 'Email', key: 'email', width: 32 },
    { header: 'Company', key: 'company', width: 28 },
    { header: 'Industry', key: 'industry', width: 22 },
    { header: 'Employee Count', key: 'employee_count', width: 18 },
    { header: 'Title', key: 'title', width: 28 }
  ];

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B2A4A' } }; // navy

  leads.forEach((lead) => sheet.addRow(lead));
  sheet.autoFilter = { from: 'A1', to: 'F1' };

  return workbook.xlsx.writeBuffer();
}

// ---- Step 5: Email the report via SendGrid ----
async function sendEmail(excelBuffer, count) {
  sgMail.setApiKey(SENDGRID_API_KEY);

  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: '2-digit', day: '2-digit' });

  const msg = {
    to: TO_EMAIL,
    from: FROM_EMAIL,
    subject: `State of Values Report - Daily Lead Research (${today})`,
    text: excelBuffer
      ? `Attached is today's lead research report: ${count} new submission(s) in the last 24 hours.`
      : 'No new "State of Values Report" submissions in the last 24 hours - nothing to report today.'
  };

  if (excelBuffer) {
    msg.attachments = [
      {
        content: Buffer.from(excelBuffer).toString('base64'),
        filename: `state-of-values-leads-${today.replace(/\//g, '-')}.xlsx`,
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        disposition: 'attachment'
      }
    ];
  }

  await sgMail.send(msg);
}

main().catch((err) => {
  console.error('daily-report failed:', err);
  process.exit(1);
});
