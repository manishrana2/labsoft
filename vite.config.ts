import { defineConfig } from 'vite'

export default defineConfig({
	server: {
		host: '127.0.0.1',
		allowedHosts: ['.loca.lt', 'localhost', '127.0.0.1'],
		proxy: {
			'/api': {
				target: 'http://localhost:3001',
				changeOrigin: true
			}
		}
	}
})
