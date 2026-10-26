async (page) => {
  const cfg = __QUOTA_CONFIG__;
  const request = cfg.request;
  const expectedHost = 'customervoice.microsoft.com';
  const expectedTitle = 'Microsoft Foundry Service: Request for Quota Increase';
  const commonFields = [
    'First Name', 'Last Name', 'Company Email', 'Company Name', 'Company Address',
    'Company City', 'Company Postal Code', 'Company Country', 'Subscription Id', 'Justification'
  ];
  const submitNames = /^(?:Submit|제출)$/i;
  const successText = /^\s*(?:Your response (?:was|has been) submitted|응답이 제출되었습니다|응답을 제출했습니다|Thanks! You've completed the request!)[.!]?\s*$/i;
  const escape = text => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const label = text => new RegExp('^\\s*(?:\\d+\\.\\s*)?' + escape(text) + '\\s*\\*?\\s*$', 'i');
  const textbox = text => page.getByRole('textbox', {name:label(text)});
  const modelBox = () => page.getByRole('listbox', {name:/Global Standard Model/i});
  const regionBox = () => page.getByRole('listbox', {name:/Global Standard Region/i});
  const quotaBox = () => page.getByRole('textbox', {name:/Global Standard Quota/i});
  const joinPath = (directory, filename) =>
    directory.replace(/[\\/]+$/, '') + (directory.includes('\\') ? '\\' : '/') + filename;
  const filePrefix = request.request_id.replace(/[^a-zA-Z0-9._-]/g, '_') +
    '.' + request.payload_hash.slice(0, 12);
  const screenshotPath = suffix => joinPath(cfg.screenshot_dir, filePrefix + '.' + suffix + '.png');
  const evidence = {
    request_id:request.request_id, payload_hash:request.payload_hash,
    quota_type:null, model_type:null, deployment_type:null, model_name:null,
    quota_input:null, region_asked:false, region_selected:null,
    common_verified:false, subscription_id_verified:false,
    screenshot_paths:[], visible_errors:[], submitted:false,
    model_select_attempts:0, region_select_attempts:0,
    input_corrections:[],
    tested_at:new Date().toISOString()
  };
  if (!['fill', 'submit'].includes(cfg.action)) throw new Error('Unsupported browser action');
  const expectedUrl = new RegExp(
    '^https://' + escape(expectedHost) + '/Pages/ResponsePage\\.aspx(?:\\?|$)', 'i'
  );
  if (!expectedUrl.test(page.url()) ||
      (await page.title()) !== expectedTitle) {
    throw new Error('Navigate to the official quota increase form and verify its title before filling');
  }
  if (request.quota_type !== 'Model Deployment (PTU/RPM/TPM)' ||
      !['Azure OpenAI', 'Azure Direct Model (ADM)'].includes(request.model_type) ||
      request.deployment_type !== 'Global Standard') {
    throw new Error('This skill version supports only Azure OpenAI/ADM Global Standard TPM requests');
  }
  if (request.model_type === 'Azure OpenAI' &&
      /audio|realtime|real-time|whisper|(?:^|-)tts|transcrib|translate|(?:^|-)live(?:-|$)/i.test(request.model_name)) {
    throw new Error('Audio/realtime/RPM models cannot use the TPM-only workbook conversion');
  }
  if (!Number.isSafeInteger(request.requested_tpm) || request.requested_tpm <= 0 ||
      request.requested_tpm % 1000 !== 0 ||
      request.quota_input !== String(request.requested_tpm / 1000)) {
    throw new Error('Invalid TPM total or quota conversion');
  }

  async function fillCommitted(control, value, name) {
    const expected = String(value);
    for (let attempt = 1; attempt <= 2; attempt++) {
      await control.waitFor({state:'visible'});
      await control.fill(expected);
      await control.blur();
      await page.waitForTimeout(150);
      if ((await control.inputValue()) === expected) return null;
      evidence.input_corrections.push({field:name, attempt});
    }
    return 'Input did not persist after blur: ' + name;
  }

  async function capture(suffix) {
    const original = page.viewportSize();
    const metrics = await page.evaluate(() => ({
      width:innerWidth,
      height:Math.max(document.documentElement.scrollHeight, ...[...document.querySelectorAll('*')]
        .filter(element => {
          const style = getComputedStyle(element);
          return element.getBoundingClientRect().width > 500 &&
            ['auto', 'scroll'].includes(style.overflowY);
        }).map(element => element.scrollHeight))
    }));
    if (metrics.height > 12000) {
      throw new Error('Form exceeds safe screenshot height; capture and review scroll segments manually');
    }
    try {
      await page.setViewportSize({
        width:original?.width || metrics.width,
        height:Math.max(original?.height || 1000, metrics.height + 120)
      });
      await page.evaluate(() => {
        for (const element of document.querySelectorAll('*')) {
          if (element.scrollTop > 0 && element.getBoundingClientRect().width > 500) {
            element.scrollTop = 0;
          }
        }
        window.scrollTo(0, 0);
      });
      const path = screenshotPath(suffix);
      await page.screenshot({path, fullPage:true});
      return path;
    } finally {
      if (original) await page.setViewportSize(original);
    }
  }

  async function verifyFields() {
    for (const name of commonFields) {
      const control = textbox(name);
      if ((await control.count()) !== 1) return 'Missing or ambiguous common field: ' + name;
      const expected = name === 'Subscription Id' ? request.subscription_id : String(cfg.common[name]);
      if ((await control.inputValue()) !== expected) return 'Common field mismatch: ' + name;
    }
    for (const value of [request.quota_type, request.model_type, request.deployment_type]) {
      const radio = page.getByRole('radio', {name:value, exact:true});
      if ((await radio.count()) !== 1 || !(await radio.isChecked())) return 'Radio selection mismatch: ' + value;
    }
    if ((await modelBox().count()) !== 1 ||
        (await modelBox().innerText()).trim() !== request.model_name) return 'Model selection mismatch';
    if ((await quotaBox().count()) !== 1 ||
        (await quotaBox().inputValue()) !== request.quota_input) return 'Quota input mismatch';
    const asksRegion = (await regionBox().count()) > 0;
    if (asksRegion && (!request.region ||
        (await regionBox().innerText()).trim() !== request.region)) return 'Required region mismatch';
    evidence.region_asked = asksRegion;
    evidence.region_selected = asksRegion ? (await regionBox().innerText()).trim() : null;
    const invalid = await page.locator('[aria-invalid="true"]:visible').count();
    const alerts = (await page.getByRole('alert').allTextContents()).filter(text => text.trim());
    const nativeErrors = await page.evaluate(() =>
      [...document.querySelectorAll('input,textarea,select')]
        .filter(element => element.getClientRects().length && element.validity && !element.validity.valid)
        .map(element => element.validationMessage)
    );
    evidence.visible_errors = alerts.concat(nativeErrors);
    if (invalid || evidence.visible_errors.length) return 'Visible client-side validation error';
    evidence.quota_type = request.quota_type;
    evidence.model_type = request.model_type;
    evidence.deployment_type = request.deployment_type;
    evidence.model_name = request.model_name;
    evidence.quota_input = request.quota_input;
    evidence.common_verified = true;
    evidence.subscription_id_verified = true;
    return null;
  }

  async function selectExactOption(box, optionName, selectionType, maxAttempts = 4) {
    let lastError = null;
    let observedOptions = [];
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      evidence[selectionType + '_select_attempts'] = attempt;
      await box.waitFor({state:'visible'});
      if ((await box.getAttribute('aria-expanded')) !== 'true') await box.click();
      const choices = box.getByRole('listbox');
      const options = choices.getByRole('option');
      observedOptions = await options.allTextContents();
      const choice = options.filter({hasText:new RegExp('^\\s*' + escape(optionName) + '\\s*$')});
      if ((await choice.count()) === 1) {
        try {
          await choice.first().scrollIntoViewIfNeeded();
          await choice.first().click();
          await page.waitForTimeout(120);
          if ((await box.innerText()).trim() === optionName) {
            return {ok:true, attempts:attempt};
          }
          await page.keyboard.press('Enter');
          await page.waitForTimeout(120);
          if ((await box.innerText()).trim() === optionName) {
            return {ok:true, attempts:attempt};
          }
          lastError = selectionType + ' selection did not persist after click.';
        } catch (error) {
          lastError = selectionType + ' selection click failed: ' + String(error);
        }
      } else {
        lastError = selectionType + ' option count was ' + (await choice.count()) + '.';
      }
      if ((await box.getAttribute('aria-expanded')) === 'true') await box.click();
      await page.waitForTimeout(180 * attempt);
    }
    return {ok:false, attempts:maxAttempts, lastError, observedOptions};
  }

  if (cfg.action === 'fill') {
    if (['submitted', 'submitting', 'unknown'].includes(request.status)) {
      throw new Error('Submission state must be reconciled before refilling this request');
    }
    await textbox('Subscription Id').waitFor({state:'visible'});
    await page.waitForTimeout(1000);
    for (const name of commonFields) {
      const value = name === 'Subscription Id' ? request.subscription_id : cfg.common[name];
      if (value === undefined || value === null || value === '') {
        throw new Error('Required common field is missing: ' + name);
      }
      const error = await fillCommitted(textbox(name), value, name);
      if (error) {
        evidence.status = 'validation_error';
        evidence.detail = error;
        evidence.screenshot_paths.push(await capture('input-not-persisted'));
        return evidence;
      }
    }
    await textbox('Justification').blur();
    await page.getByRole('radio', {name:request.quota_type, exact:true}).check();
    await page.getByRole('radio', {name:request.model_type, exact:true}).check();
    await page.getByRole('radio', {name:request.deployment_type, exact:true}).check();
    const modelSelection = await selectExactOption(modelBox(), request.model_name, 'model');
    if (!modelSelection.ok) {
      evidence.status = 'model_missing';
      evidence.detail = 'Exact model is absent/unstable in the live dropdown after retries; no substitution attempted.';
      evidence.live_model_options = modelSelection.observedOptions || [];
      evidence.last_error = modelSelection.lastError || null;
      evidence.screenshot_paths.push(await capture('model-missing'));
      if ((await modelBox().getAttribute('aria-expanded')) === 'true') await modelBox().click();
      return evidence;
    }
    await quotaBox().waitFor({state:'visible'});
    const helpText = await quotaBox().evaluate(element => {
      let node = element;
      for (let index = 0; node && index < 6; index++, node = node.parentElement) {
        if (/Global Standard Quota/i.test(node.textContent)) return node.textContent;
      }
      return '';
    });
    const unitConfirmed = request.model_type === 'Azure Direct Model (ADM)'
      ? /1\s*:\s*1,?000\s*TPM/i.test(helpText)
      : /kTPM/i.test(helpText);
    if (!unitConfirmed) {
      evidence.status = 'unsupported_form';
      evidence.detail = 'The live form does not confirm the expected kTPM/1:1000 unit rule. Stop for review.';
      evidence.screenshot_paths.push(await capture('unit-review'));
      return evidence;
    }
    if ((await regionBox().count()) > 0) {
      evidence.region_asked = true;
      if (!request.region) {
        evidence.status = 'region_missing';
        evidence.detail = 'The selected model asks for a region, but the request has none.';
        evidence.screenshot_paths.push(await capture('region-missing'));
        return evidence;
      }
      const regionSelection = await selectExactOption(regionBox(), request.region, 'region');
      if (!regionSelection.ok) {
        evidence.status = 'region_missing';
        evidence.detail = 'Requested region is absent/unstable in the live dropdown after retries.';
        evidence.live_region_options = regionSelection.observedOptions || [];
        evidence.last_error = regionSelection.lastError || null;
        evidence.screenshot_paths.push(await capture('region-unavailable'));
        if ((await regionBox().getAttribute('aria-expanded')) === 'true') await regionBox().click();
        return evidence;
      }
    }
    const error = await fillCommitted(quotaBox(), request.quota_input, 'Global Standard Quota') ||
      await verifyFields();
    evidence.status = error ? 'validation_error' : 'filled';
    evidence.detail = error || 'All common fields and row values match; input-only capture, never submitted.';
    evidence.screenshot_paths.push(await capture(error ? 'validation-error' : 'filled'));
    return evidence;
  }

  const ticket = cfg.ticket;
  if (cfg.mode !== 'submit' || request.execute !== 'Y' || request.status !== 'submitting' ||
      !ticket || ticket.approved !== true || ticket.request_id !== request.request_id ||
      ticket.payload_hash !== request.payload_hash || !ticket.attempt_id ||
      !ticket.approval_fingerprint) {
    throw new Error('Missing matching approved single-attempt submission ticket');
  }
  const submission = {
    request_id:request.request_id, payload_hash:request.payload_hash,
    attempt_id:ticket.attempt_id, submit_clicked:false, submission_status:'not_clicked',
    confirmation_text:null, confirmation_screenshot:null, detail:null
  };
  const mismatch = await verifyFields();
  if (mismatch) {
    submission.detail = mismatch + '; Submit was not clicked.';
    submission.confirmation_screenshot = await capture('pre-submit-mismatch');
    return submission;
  }
  const submit = page.getByRole('button', {name:submitNames});
  if ((await submit.count()) !== 1 || !(await submit.isEnabled())) {
    submission.detail = 'Submit button missing, ambiguous or disabled; no click attempted.';
    submission.confirmation_screenshot = await capture('submit-unavailable');
    return submission;
  }
  await capture('ready-to-submit');
  const finalMismatch = await verifyFields();
  if (finalMismatch) {
    submission.detail = finalMismatch + '; Submit was not clicked.';
    submission.confirmation_screenshot = await capture('final-pre-submit-mismatch');
    return submission;
  }
  // Any timeout or browser failure from this point is an unknown attempt, never an automatic retry.
  submission.submit_clicked = true;
  await submit.click({timeout:15000, noWaitAfter:true});
  const confirmation = page.getByText(successText).first();
  try {
    await confirmation.waitFor({state:'visible', timeout:15000});
  } catch (error) {
    if (error.name !== 'TimeoutError') throw error;
    submission.submission_status = 'unknown';
    submission.detail = 'No explicit success confirmation appeared after the single click; reconcile before retry.';
    submission.confirmation_screenshot = await capture('submission-unknown');
    return submission;
  }
  if ((await modelBox().count()) > 0 && (await modelBox().isVisible())) {
    submission.submission_status = 'unknown';
    submission.detail = 'Success-like text is present but the input form remains visible; reconcile manually.';
    submission.confirmation_screenshot = await capture('ambiguous-confirmation');
    return submission;
  }
  submission.submission_status = 'submitted';
  submission.confirmation_text = (await confirmation.innerText()).trim();
  submission.confirmation_screenshot = await capture('submitted');
  submission.detail = 'One submit click followed by explicit visible submission confirmation.';
  return submission;
}
