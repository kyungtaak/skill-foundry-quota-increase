"""Generate a Playwright-agent smoke test that intercepts every network request."""

import argparse
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
COMMON = {
    "First Name": "Alex",
    "Last Name": "Example",
    "Company Email": "alex@example.invalid",
    "Company Name": "Example Corporation",
    "Company Address": "123 Example Street",
    "Company City": "Example City",
    "Company Postal Code": "00001",
    "Company Country": "Example Country",
    "Subscription Id": "11111111-1111-4111-8111-111111111111",
    "Justification": "Synthetic local fixture only. No application will be transmitted.",
}
REQUEST = {
    "request_id": "LOCAL-TEST-001",
    "execute": "Y",
    "quota_type": "Model Deployment (PTU/RPM/TPM)",
    "model_type": "Azure Direct Model (ADM)",
    "deployment_type": "Global Standard",
    "model_name": "Example-Model",
    "region": "Korea Central",
    "requested_tpm": 1000000,
    "quota_input": "1000",
    "subscription_id": COMMON["Subscription Id"],
    "payload_hash": "a" * 64,
    "status": "pending",
}

HARNESS = r"""
async (page) => {
  const browser = page.context().browser();
  const context = await browser.newContext({viewport:{width:1280,height:900}});
  const blocked = [];
  const served = [];
  const url = 'https://customervoice.microsoft.com/Pages/ResponsePage.aspx?local-skill-fixture=1';
  const html = __FIXTURE_HTML__;
  await context.route('**/*', async route => {
    const request = route.request();
    if (request.url() === url && request.method() === 'GET') {
      served.push(request.url());
      await route.fulfill({status:200, contentType:'text/html', body:html});
    } else {
      blocked.push(request.url());
      await route.abort();
    }
  });
  const target = await context.newPage();
  const execute = async (target, config) => await (__FLOW__)(target);
  const common = __COMMON__;
  const request = __REQUEST__;
  const root = __CAPTURES__;
  const base = {action:'fill', common, request, mode:'capture', ticket:null, screenshot_dir:root};
  const cases = [];
  try {
    await target.goto(url);
    const filled = await execute(target, base);
    if (filled.status !== 'filled' || filled.region_selected !== 'Korea Central') {
      throw new Error('Fill fixture failed: ' + JSON.stringify(filled));
    }
    if ((await target.evaluate(()=>window.syntheticSubmitClicks)) !== 0) throw new Error('Fill clicked Submit');
    cases.push('fill_and_capture_no_submit');

    let missingTicketBlocked = false;
    try {
      await execute(target, {...base, action:'submit', mode:'submit',
        request:{...request,status:'submitting'}});
    } catch (error) {
      if (!String(error.message).includes('submission ticket')) throw error;
      missingTicketBlocked = true;
    }
    if (!missingTicketBlocked) throw new Error('Missing ticket was not blocked');
    cases.push('submit_without_ticket_blocked');

    const ticket = {approved:true,request_id:request.request_id,payload_hash:request.payload_hash,
      attempt_id:'local-fixture-attempt',approval_fingerprint:'local-fixture-approval'};
    await target.getByRole('textbox',{name:/Company City/}).fill('Changed City');
    const mismatch = await execute(target, {...base,action:'submit',mode:'submit',
      request:{...request,status:'submitting'},ticket});
    if (mismatch.submission_status !== 'not_clicked' || mismatch.submit_clicked) throw new Error('Changed field was not blocked');
    cases.push('changed_common_field_blocks_submit');
    await execute(target, base);

    const screenshot = target.screenshot.bind(target);
    target.screenshot = async options => {
      const result = await screenshot(options);
      if (String(options.path).includes('ready-to-submit')) {
        await target.getByRole('textbox',{name:/Company City/}).fill('Changed after capture');
      }
      return result;
    };
    let lateMismatch;
    try {
      lateMismatch = await execute(target, {...base,action:'submit',mode:'submit',
        request:{...request,status:'submitting'},ticket});
    } finally {
      target.screenshot = screenshot;
    }
    if (lateMismatch.submission_status !== 'not_clicked' || lateMismatch.submit_clicked ||
        (await target.evaluate(()=>window.syntheticSubmitClicks)) !== 0) {
      throw new Error('Field change during screenshot was not blocked');
    }
    cases.push('field_changed_after_capture_blocks_submit');
    await execute(target, base);
    await target.evaluate(() => {
      document.querySelector('#confirmation h1').textContent = "Thanks! You've completed the request!";
    });
    const submitted = await execute(target, {...base,action:'submit',mode:'submit',
      request:{...request,status:'submitting'},ticket});
    if (submitted.submission_status !== 'submitted' || !submitted.submit_clicked ||
        submitted.confirmation_text !== "Thanks! You've completed the request!" ||
        (await target.evaluate(()=>window.syntheticSubmitClicks)) !== 1) {
      throw new Error('Synthetic submit confirmation failed');
    }
    cases.push('one_mock_click_explicit_confirmation');

    await target.goto(url);
    const missing = await execute(target, {...base,request:{...request,model_name:'Missing-Example-Model'}});
    if (missing.status !== 'model_missing' ||
        (await target.evaluate(()=>window.syntheticSubmitClicks)) !== 0) throw new Error('Missing model test failed');
    cases.push('missing_model_no_substitution');

    await target.goto(url);
    const noRegion = await execute(target, {...base,request:{...request,model_name:'Example-No-Region',region:null}});
    if (noRegion.status !== 'filled' || noRegion.region_asked) throw new Error('Conditional region test failed');
    cases.push('conditional_region_not_forced');

    await target.goto(url);
    await target.getByRole('textbox',{name:/Subscription Id/}).evaluate(element => {
      element.addEventListener('blur', () => {
        if (window.syntheticFocusRaceDone) return;
        window.syntheticFocusRaceDone = true;
        document.querySelector('[aria-label="1. First Name *"]').value += element.value;
        element.value = '';
      });
    });
    const raced = await execute(target, base);
    if (raced.status !== 'validation_error' || raced.submitted ||
        !raced.input_corrections.some(item => item.field === 'Subscription Id') ||
        (await target.evaluate(()=>window.syntheticSubmitClicks)) !== 0) {
      throw new Error('Cross-field focus race was not detected');
    }
    const refilled = await execute(target, base);
    if (refilled.status !== 'filled') throw new Error('Pre-click refill failed');
    cases.push('focus_race_detected_then_explicit_pre_click_refill');

    await target.goto(url);
    await target.getByRole('textbox',{name:/First Name/}).evaluate(element => {
      element.addEventListener('blur', () => { element.value = ''; });
    });
    const persistent = await execute(target, base);
    if (persistent.status !== 'validation_error' || persistent.input_corrections.length !== 2 ||
        (await target.evaluate(()=>window.syntheticSubmitClicks)) !== 0) {
      throw new Error('Persistent input failure was not bounded and blocked');
    }
    cases.push('persistent_input_failure_stops_after_two_attempts');
    return {status:'success',cases,real_submissions:0,network_intercepted:true,
      fixture_page_loads:served.length,blocked_requests:blocked.length};
  } finally {
    await context.close();
  }
}
"""


def make(output, captures):
    output = Path(output).resolve()
    captures = Path(captures).resolve()
    if output.exists():
        raise FileExistsError(output)
    captures.mkdir(parents=True, exist_ok=True)
    flow = (ROOT / "scripts" / "browser_flow.js").read_text(encoding="utf-8")
    flow = flow.replace("__QUOTA_CONFIG__", "config")
    code = HARNESS.replace("__FLOW__", flow)
    code = code.replace("__FIXTURE_HTML__", json.dumps(
        (ROOT / "tests" / "fixtures" / "form.html").read_text(encoding="utf-8")
    ))
    code = code.replace("__COMMON__", json.dumps(COMMON))
    code = code.replace("__REQUEST__", json.dumps(REQUEST))
    code = code.replace("__CAPTURES__", json.dumps(str(captures)))
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(code, encoding="utf-8")
    return {"script": str(output), "network_policy": "intercept all; no real submission"}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True)
    parser.add_argument("--captures", required=True)
    args = parser.parse_args()
    print(json.dumps(make(args.output, args.captures)))
