interface TokenCache {
  token: string;
  expiresAt: number;
}

const tokenCache = new Map<string, TokenCache>();

export class GraphClient {
  private tenantId: string;
  private clientId: string;
  private clientSecret: string;
  private cacheKey: string;

  constructor(tenantId: string, clientId: string, clientSecret: string) {
    this.tenantId = tenantId;
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.cacheKey = `${tenantId}:${clientId}`;
  }

  async getAccessToken(): Promise<string> {
    const cached = tokenCache.get(this.cacheKey);
    if (cached && cached.expiresAt > Date.now() + 60000) {
      return cached.token;
    }

    const tokenUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    });

    const res = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });

    if (!res.ok) {
      const errorData = await res.json().catch(() => ({}));
      throw new Error(`Auth failed: ${(errorData as any).error_description || (errorData as any).error || res.statusText}`);
    }

    const data = await res.json() as { access_token: string; expires_in: number };
    tokenCache.set(this.cacheKey, {
      token: data.access_token,
      expiresAt: Date.now() + data.expires_in * 1000,
    });

    return data.access_token;
  }

  async request(path: string, options: RequestInit = {}): Promise<Response> {
    const token = await this.getAccessToken();
    const url = path.startsWith('http') ? path : `https://graph.microsoft.com/v1.0${path}`;
    return fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...options.headers,
      },
    });
  }

  async get(path: string): Promise<any> {
    const res = await this.request(path);
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Graph API GET ${path} failed (${res.status}): ${err}`);
    }
    return res.json();
  }

  async post(path: string, body: any): Promise<any> {
    const res = await this.request(path, {
      method: 'POST',
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Graph API POST ${path} failed (${res.status}): ${err}`);
    }
    if (res.status === 204) return null;
    return res.json();
  }

  async patch(path: string, body: any): Promise<any> {
    const res = await this.request(path, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Graph API PATCH ${path} failed (${res.status}): ${err}`);
    }
    if (res.status === 204) return null;
    return res.json().catch(() => null);
  }

  async put(path: string, body: Buffer | string, contentType: string): Promise<any> {
    const token = await this.getAccessToken();
    const url = path.startsWith('http') ? path : `https://graph.microsoft.com/v1.0${path}`;
    const res = await fetch(url, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': contentType,
      },
      body,
    });
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Graph API PUT ${path} failed (${res.status}): ${err}`);
    }
    if (res.status === 204) return null;
    return res.json();
  }

  async getBuffer(path: string): Promise<Buffer> {
    const res = await this.request(path);
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Graph API GET ${path} failed (${res.status}): ${err}`);
    }
    const arrayBuffer = await res.arrayBuffer();
    return Buffer.from(arrayBuffer);
  }

  async getAllPages<T>(path: string): Promise<T[]> {
    const results: T[] = [];
    let nextLink: string | undefined = path;

    while (nextLink) {
      const data = await this.get(nextLink);
      if (data.value) {
        results.push(...data.value);
      }
      nextLink = data['@odata.nextLink'];
    }

    return results;
  }
}
