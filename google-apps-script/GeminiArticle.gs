/**
 * Calls Gemini API to turn Q/A form content into a newsletter article and title in one shot.
 * Store your API key in Script Properties: GEMINI_API_KEY (Project Settings → Script properties).
 */

/**
 * Generates both the article body and a matching title from Q/A text using a single Gemini call.
 * The title is derived from the finished article so they are always coherent with each other.
 *
 * @param {string} qaText - Q/A format text from buildArticleFromResponse
 * @return {Object} { title {string}, body {string} } — both fall back to empty string on error
 */
function generateArticle(qaText) {
  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    Logger.log('GeminiArticle: GEMINI_API_KEY not set in Script Properties. Skipping AI.');
    return { title: '', body: '' };
  }

  var model = Config.GEMINI_MODEL || 'gemini-2.5-flash-lite';
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';
  Logger.log('GeminiArticle: calling ' + model + ' with Q/A length ' + (qaText ? qaText.length : 0));

  var prompt = 'You are a professional newsletter writer and editor.\n'
    + 'Below is a form submission with raw details and responses.\n'
    + 'Analyze the details: if it highlights a person (e.g. name, role, background), write it as an engaging community feature / spotlight; if it describes an event, project milestone, launch, or partnership, write it as an exciting announcement or recap.\n'
    + 'Turn this raw information into a beautifully polished, highly engaging newsletter article suitable for publishing (e.g. on Substack).\n'
    + 'Write in a warm, professional, and appealing tone. Use the answers to build the narrative organically; do not repeat the "Q:" and "A:" labels, template references, or consent checkboxes in the finished article.\n\n'
    + 'Return a JSON object with exactly two fields:\n'
    + '  "title": a short, punchy headline for the finished article (maximum 10 words, no trailing punctuation)\n'
    + '  "body": the full polished article text (no title line, clean paragraph breaks separated by \\n\\n)\n\n'
    + qaText;

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 2100,
      responseMimeType: 'application/json'
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var responseBody = response.getContentText();

    if (code === 429 || code === 503) {
      var waitMs = code === 429 ? 36000 : 10000;
      Logger.log('GeminiArticle: error ' + code + '. Retrying once in ' + (waitMs / 1000) + 's...');
      Utilities.sleep(waitMs);
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
      responseBody = response.getContentText();
    }

    if (code !== 200) {
      Logger.log('GeminiArticle: API error HTTP ' + code);
      if (code === 429) {
        Logger.log('GeminiArticle: still over quota. Check https://ai.google.dev/gemini-api/docs/rate-limits or try again later.');
      } else if (code === 503) {
        Logger.log('GeminiArticle: still unavailable (503). Gemini is experiencing high demand. Try again in a few minutes.');
      }
      Logger.log('GeminiArticle: response body: ' + responseBody);
      return { title: '', body: '' };
    }

    var data = JSON.parse(responseBody);
    var rawText = data.candidates && data.candidates[0] && data.candidates[0].content
      && data.candidates[0].content.parts && data.candidates[0].content.parts[0]
      ? data.candidates[0].content.parts[0].text
      : '';

    if (!rawText) {
      Logger.log('GeminiArticle: no text in response. candidates: ' + (data.candidates ? data.candidates.length : 0));
      if (data.candidates && data.candidates[0] && data.candidates[0].finishReason) {
        Logger.log('GeminiArticle: finishReason: ' + data.candidates[0].finishReason);
      }
      if (data.error) {
        Logger.log('GeminiArticle: data.error: ' + JSON.stringify(data.error));
      }
      return { title: '', body: '' };
    }

    // Parse the JSON returned by Gemini
    var result;
    try {
      result = JSON.parse(rawText);
    } catch (parseErr) {
      Logger.log('GeminiArticle: JSON parse failed (' + parseErr.message + '). Using raw text as body.');
      return { title: '', body: rawText.trim() };
    }

    var title = (result.title || '').trim().replace(/[.!?]+$/, ''); // strip trailing punctuation
    var body  = (result.body  || '').trim();

    Logger.log('GeminiArticle: success — title: "' + title + '", body: ' + body.length + ' chars');
    return { title: title, body: body };

  } catch (err) {
    Logger.log('GeminiArticle: catch ' + err.message);
    Logger.log('GeminiArticle: stack ' + (err.stack || ''));
    return { title: '', body: '' };
  }
}

function getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}
