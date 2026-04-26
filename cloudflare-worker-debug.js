/**
 * Cloudflare Workers 反代脚本（调试版）
 *
 * 使用前请将 UPSTREAM_URL 改为自己的上游 API 基础地址。
 * 调试版只输出路径信息，不输出请求头、API Key 或完整私有地址。
 */

const UPSTREAM_URL = 'https://example.invalid/v1';
const ALLOWED_ORIGINS = ['*'];
const DEBUG = true;

addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request));
});

async function handleRequest(request) {
  if (request.method === 'OPTIONS') {
    return handleCORS(request);
  }

  if (!['GET', 'POST'].includes(request.method)) {
    return new Response('Method Not Allowed', {
      status: 405,
      headers: getCORSHeaders(request),
    });
  }

  try {
    const url = new URL(request.url);
    const pathname = url.pathname.endsWith('/models') ? '/models' : '/chat/completions';
    const upstreamUrl = `${UPSTREAM_URL}${pathname}`;

    if (DEBUG) {
      console.log('Response Refiner Worker debug:', JSON.stringify({
        method: request.method,
        pathname,
      }));
    }

    const upstreamRequest = new Request(upstreamUrl, {
      method: request.method,
      headers: request.headers,
      body: request.method === 'GET' ? undefined : request.body,
    });

    const upstreamResponse = await fetch(upstreamRequest);
    const response = new Response(upstreamResponse.body, {
      status: upstreamResponse.status,
      statusText: upstreamResponse.statusText,
      headers: upstreamResponse.headers,
    });

    for (const [key, value] of Object.entries(getCORSHeaders(request))) {
      response.headers.set(key, value);
    }

    return response;
  } catch (error) {
    return new Response(JSON.stringify({
      error: true,
      message: error.message || 'Internal Server Error',
    }), {
      status: 500,
      headers: {
        'Content-Type': 'application/json',
        ...getCORSHeaders(request),
      },
    });
  }
}

function handleCORS(request) {
  return new Response(null, {
    status: 204,
    headers: getCORSHeaders(request),
  });
}

function getCORSHeaders(request) {
  const origin = request.headers.get('Origin');
  const headers = {
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Requested-With',
    'Access-Control-Max-Age': '86400',
  };

  if (ALLOWED_ORIGINS.includes('*')) {
    headers['Access-Control-Allow-Origin'] = '*';
  } else if (origin && ALLOWED_ORIGINS.includes(origin)) {
    headers['Access-Control-Allow-Origin'] = origin;
    headers['Vary'] = 'Origin';
  }

  return headers;
}
