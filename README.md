# Daily Lead Report

Every day at ~6:00 AM Central, a GitHub Actions workflow runs `daily-report.js` directly
(no hosting provider involved). It pulls the last 24 hours of "State of Values Report"
Jotform submissions, researches each lead (company, industry, employee count, title)
using the Anthropic API with web search, builds an Excel file, and emails it to
asegal@think2perform.com via SendGrid.

## How it runs

`.github/workflows/daily-report.yml` triggers on a schedule, checks out the repo, installs
dependencies, and runs `node daily-report.js` with your secrets passed in as environment
variables. There's no server, no deployment, and nothing to keep "awake" - GitHub's
runner spins up, runs the script once, and shuts down.

## Repository secrets (Settings > Secrets and variables > Actions)

| Secret | Where to get it |
|---|---|
| `JOTFORM_API_KEY` | Jotform account settings -> API |
| `ANTHROPIC_API_KEY` | Anthropic Console -> API Keys |
| `SENDGRID_API_KEY` | SendGrid dashboard -> Settings -> API Keys |
| `FROM_EMAIL` | The sender address you verified in SendGrid |
| `TO_EMAIL` | Optional - defaults to asegal@think2perform.com if omitted |

## Testing

Go to the repo's Actions tab, open "Daily lead report" in the left sidebar, and click
"Run workflow" to trigger it by hand anytime - you don't have to wait for 6am. Check the
run's logs for errors, and check the inbox for the email.

## Notes

- The cron schedule is in UTC. `11:00 UTC` = 6:00 AM Central Daylight Time. When clocks
  fall back to Standard Time in November, this will fire at 5:00 AM Central until the
  schedule in the workflow file is updated to `0 12 * * *` (GitHub Actions doesn't
  auto-adjust for daylight saving).
- GitHub Actions' `schedule` trigger isn't second-precise - it can run a few minutes late
  during high load, but it's reliable about running.
- No `CRON_SECRET` or public endpoint is needed anymore, since nothing is exposed to the
  internet - the script only ever runs inside GitHub's own runner.
