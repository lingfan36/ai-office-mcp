# Privacy Policy — Word MCP Live

**Effective Date:** 2026-03-08

## Overview

Word MCP Live is a locally-running MCP (Model Context Protocol) server that enables AI assistants to read and edit Microsoft Word documents. This privacy policy explains how user data is handled.

## Data Collection

**Word MCP Live does not collect, transmit, or store any user data.** Specifically:

- **No telemetry** — no usage analytics or crash reports are sent anywhere
- **No network requests** — the server does not contact external services (unless you explicitly configure SSE or HTTP transport for remote access)
- **No data storage** — no databases, logs, or caches are created beyond what is needed for the current session
- **No third-party services** — no external APIs, CDNs, or tracking services are used

## How It Works

- The server runs entirely on your local machine
- It communicates with your MCP client (e.g., Claude Desktop) via stdio (standard input/output) by default
- Documents are read from and written to your local filesystem only
- On Windows, the server may interact with Microsoft Word via COM automation — this is a direct local process communication, not a network call

## Remote Transport (Optional)

If you configure the server to use SSE or HTTP transport (via `MCP_TRANSPORT` environment variable), the server will listen on a network port. In this case:

- You are responsible for securing the network connection (firewall, TLS, authentication)
- The server itself does not implement authentication or encryption
- Document content will be transmitted over the network connection you configured

## Environment Variables

The server reads the following optional environment variables. These are stored locally and never transmitted:

- `MCP_AUTHOR` — author name for tracked changes and comments
- `MCP_AUTHOR_INITIALS` — author initials for comments
- `MCP_TRANSPORT`, `MCP_HOST`, `MCP_PORT` — transport configuration

## Open Source

This project is open source under the MIT License. You can audit the full source code at:
https://github.com/ykarapazar/word-mcp-live

## Contact

For privacy questions or concerns:
- GitHub Issues: https://github.com/ykarapazar/word-mcp-live/issues
- Maintainer: Yüce Karapazar
