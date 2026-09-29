import * as http from 'node:http';

export interface AuthCodeResult {
	code: string;
	state: string;
}

const SUCCESS_HTML = '<html><body>Signed in. You can close this tab and return to Obsidian.</body></html>';
const ERROR_HTML = '<html><body>Sign-in failed. You can close this tab and return to Obsidian.</body></html>';

/**
 * Listens for the single OAuth redirect on the app's registered
 * http://localhost:8080/callback redirect URI, then closes.
 */
export function waitForAuthCode(port: number, timeoutMs: number): Promise<AuthCodeResult> {
	return new Promise((resolve, reject) => {
		const server = http.createServer((req, res) => {
			const url = new URL(req.url ?? '/', `http://localhost:${port}`);
			if (url.pathname !== '/callback') {
				res.writeHead(404).end();
				return;
			}

			const error = url.searchParams.get('error');
			const errorDescription = url.searchParams.get('error_description');
			const code = url.searchParams.get('code');
			const state = url.searchParams.get('state');

			if (error) {
				res.writeHead(200, { 'Content-Type': 'text/html' }).end(ERROR_HTML);
				finish(() => reject(new Error(errorDescription ?? error)));
				return;
			}
			if (!code || !state) {
				res.writeHead(200, { 'Content-Type': 'text/html' }).end(ERROR_HTML);
				finish(() => reject(new Error('Redirect was missing code or state')));
				return;
			}

			res.writeHead(200, { 'Content-Type': 'text/html' }).end(SUCCESS_HTML);
			finish(() => resolve({ code, state }));
		});

		let settled = false;
		const finish = (action: () => void) => {
			if (settled) return;
			settled = true;
			window.clearTimeout(timer);
			server.close();
			action();
		};

		const timer = window.setTimeout(() => {
			finish(() => reject(new Error('Timed out waiting for Microsoft sign-in')));
		}, timeoutMs);

		server.on('error', (err) => finish(() => reject(err)));
		server.listen(port);
	});
}
