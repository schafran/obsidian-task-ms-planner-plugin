import { requestUrl } from 'obsidian';

interface NetworkRequestOptions {
	headers?: Record<string, string>;
	body?: string;
}

interface NetworkResponse<T> {
	headers: Record<string, string>;
	body: T;
	status: number;
}

/**
 * msal-node's default HttpClient uses global fetch, which Obsidian's Electron
 * renderer subjects to CORS and blocks against Microsoft identity endpoints.
 * requestUrl bypasses that, same as TodoClient does for Graph calls.
 */
export class ObsidianNetworkClient {
	async sendGetRequestAsync<T>(url: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
		return this.send<T>(url, 'GET', options);
	}

	async sendPostRequestAsync<T>(url: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
		return this.send<T>(url, 'POST', options);
	}

	private async send<T>(url: string, method: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
		const res = await requestUrl({
			url,
			method,
			headers: options?.headers,
			body: options?.body,
			throw: false,
		});
		if (res.status >= 400) {
			console.error('[todo-sync] auth request failed', url, res.status, res.text);
		}
		return { headers: res.headers, body: res.json as T, status: res.status };
	}
}
