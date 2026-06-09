module.exports = async function (context, req) {
  context.log('ClassifyDocument function processed a request.');

  if (req.method === 'OPTIONS') {
    context.res = {
      status: 204,
      headers: getCorsHeaders(),
      body: null
    };
    return;
  }

  if (!req.body || typeof req.body !== 'object') {
    context.res = {
      status: 400,
      headers: getCorsHeaders(),
      body: { error: 'Request body must contain JSON with modelId and fileBase64.' }
    };
    return;
  }

  const modelId = String(req.body.modelId || '').trim();
  const fileBase64 = String(req.body.fileBase64 || '').trim();
  const endpoint = String(process.env.DOCUMENT_INTELLIGENCE_ENDPOINT || '').trim();
  const apiKey = String(process.env.DOCUMENT_INTELLIGENCE_API_KEY || '').trim();

  if (!endpoint || !apiKey) {
    context.res = {
      status: 500,
      headers: getCorsHeaders(),
      body: { error: 'Azure Document Intelligence endpoint and API key must be configured in environment variables.' }
    };
    return;
  }

  if (!modelId || !fileBase64) {
    context.res = {
      status: 400,
      headers: getCorsHeaders(),
      body: { error: 'Request must include modelId and fileBase64.' }
    };
    return;
  }

  try {
    const fileBytes = Buffer.from(fileBase64, 'base64');
    const analyzeUrl = `${endpoint.replace(/\/+$/g, '')}/documentintelligence/documentClassifiers/${encodeURIComponent(modelId)}:analyze?api-version=2024-11-30&split=auto`;
    //const analyzeUrl = `${endpoint.replace(/\/+$/g, '')}/documentintelligence/documentModels/${encodeURIComponent(modelId)}:analyze?api-version=2024-11-30`;
    //const analyzeUrl = `${endpoint.replace(/\/+$/g, '')}/formrecognizer/v2.1/custom/models/${encodeURIComponent(modelId)}/analyze?api-version=2022-08-31`;

    const response = await fetch(analyzeUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/pdf',
        'Ocp-Apim-Subscription-Key': apiKey,
        Accept: 'application/json'
      },
      body: fileBytes
    });

    if (!response.ok) {
      const body = await response.text();
      context.res = {
        status: 502,
        headers: getCorsHeaders(),
        body: { error: `Azure Document Intelligence request failed: ${response.status} ${response.statusText}`, details: body }
      };
      return;
    }

    const operationLocation = response.headers.get('operation-location') || response.headers.get('Operation-Location');
    const resultJson = operationLocation
      ? await pollAzureOperation(operationLocation, apiKey)
      : await response.json();

    context.res = {
      status: 200,
      headers: getCorsHeaders(),
      body: resultJson
    };
  } catch (error) {
    context.log.error('Error processing classification request:', error);
    context.res = {
      status: 500,
      headers: getCorsHeaders(),
      body: { error: 'Classification function error', details: String(error) }
    };
  }
};

async function pollAzureOperation(operationUrl, apiKey) {
  const maxAttempts = 120;
  const delayMs = 1000;

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const pollResponse = await fetch(operationUrl, {
      method: 'GET',
      headers: {
        'Ocp-Apim-Subscription-Key': apiKey,
        Accept: 'application/json'
      }
    });

    if (!pollResponse.ok) {
      const body = await pollResponse.text();
      throw new Error(`Azure polling request failed: ${pollResponse.status} ${pollResponse.statusText} - ${body}`);
    }

    const pollResult = await pollResponse.json();
    const status = String(pollResult.status || '').toLowerCase();
    if (status === 'succeeded') {
      return pollResult;
    }
    if (status === 'failed' || status === 'canceled') {
      throw new Error(`Azure operation did not succeed: ${pollResult.status}`);
    }

    await new Promise(resolve => setTimeout(resolve, delayMs));
  }

  throw new Error('Azure Document Intelligence polling timed out.');
}

function getCorsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'OPTIONS, POST',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}
