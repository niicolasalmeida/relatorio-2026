import { timingSafeEqual } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const currentDir = dirname(fileURLToPath(import.meta.url));
const protectedFiles = new Map([
  [
    '',
    {
      filePath: resolve(currentDir, '..', 'Relatório - Em contrução 2026 v4.backup.html'),
      contentType: 'text/html; charset=utf-8',
    },
  ],
  [
    'Relatório - Em contrução 2026 v4.backup.html',
    {
      filePath: resolve(currentDir, '..', 'Relatório - Em contrução 2026 v4.backup.html'),
      contentType: 'text/html; charset=utf-8',
    },
  ],
  [
    'despesas.inline.js',
    {
      filePath: resolve(currentDir, '..', 'despesas.inline.js'),
      contentType: 'application/javascript; charset=utf-8',
    },
  ],
  [
    'recebimentos.inline.js',
    {
      filePath: resolve(currentDir, '..', 'recebimentos.inline.js'),
      contentType: 'application/javascript; charset=utf-8',
    },
  ],
  [
    'projecao.inline.js',
    {
      filePath: resolve(currentDir, '..', 'projecao.inline.js'),
      contentType: 'application/javascript; charset=utf-8',
    },
  ],
]);

function constantTimeEqual(left, right) {
  const leftBuffer = Buffer.from(left);
  const rightBuffer = Buffer.from(right);

  if (leftBuffer.length !== rightBuffer.length) {
    return false;
  }

  return timingSafeEqual(leftBuffer, rightBuffer);
}

function parseBasicAuth(headerValue) {
  if (!headerValue || !headerValue.startsWith('Basic ')) {
    return null;
  }

  try {
    const decoded = Buffer.from(headerValue.slice(6), 'base64').toString('utf8');
    const separatorIndex = decoded.indexOf(':');

    if (separatorIndex === -1) {
      return null;
    }

    return {
      username: decoded.slice(0, separatorIndex),
      password: decoded.slice(separatorIndex + 1),
    };
  } catch {
    return null;
  }
}

function unauthorizedResponse() {
  return new Response('Acesso restrito.', {
    status: 401,
    headers: {
      'Content-Type': 'text/plain; charset=utf-8',
      'Cache-Control': 'no-store',
      'WWW-Authenticate': 'Basic realm="Relatorio", charset="UTF-8"',
    },
  });
}

function missingConfigurationResponse() {
  return new Response(
    'Configure SITE_USERNAME e SITE_PASSWORD nas variaveis de ambiente da Vercel.',
    {
      status: 500,
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Cache-Control': 'no-store',
      },
    },
  );
}

function normalizeRequestedPath(request) {
  const url = new URL(request.url);
  return decodeURIComponent(url.searchParams.get('path') ?? '').replace(/^\/+/, '');
}

async function buildProtectedFileResponse(requestedPath) {
  const fileConfig = protectedFiles.get(requestedPath);
  const body = await readFile(fileConfig.filePath, 'utf8');

  return new Response(body, {
    status: 200,
    headers: {
      'Content-Type': fileConfig.contentType,
      'Cache-Control': 'no-store',
      'X-Robots-Tag': 'noindex, nofollow',
    },
  });
}

async function handleRequest(request, includeBody = true) {
  const requestedPath = normalizeRequestedPath(request);

  if (!protectedFiles.has(requestedPath)) {
    return new Response('Not found', {
      status: 404,
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Cache-Control': 'no-store',
      },
    });
  }

  const expectedUsername = process.env.SITE_USERNAME;
  const expectedPassword = process.env.SITE_PASSWORD;

  if (!expectedUsername || !expectedPassword) {
    return missingConfigurationResponse();
  }

  const credentials = parseBasicAuth(request.headers.get('authorization'));

  if (!credentials) {
    return unauthorizedResponse();
  }

  const usernameMatches = constantTimeEqual(credentials.username, expectedUsername);
  const passwordMatches = constantTimeEqual(credentials.password, expectedPassword);

  if (!usernameMatches || !passwordMatches) {
    return unauthorizedResponse();
  }

  try {
    const response = await buildProtectedFileResponse(requestedPath);

    if (includeBody) {
      return response;
    }

    return new Response(null, {
      status: response.status,
      headers: response.headers,
    });
  } catch (error) {
    console.error('Erro ao carregar o arquivo protegido:', error);
    return new Response('Erro ao carregar o relatorio.', {
      status: 500,
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Cache-Control': 'no-store',
      },
    },
    );
  }
}

export async function GET(request) {
  return handleRequest(request, true);
}

export async function HEAD(request) {
  return handleRequest(request, false);
}
