# Session Sushi

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

**Session Sushi** is a zero dependency browser extension for handling cookies, Microsoft 365 OAuth tokens, and Graph API interactions. Build for security professionals.

![Session Sushi](https://phishing.club/img/session-sushi/animated.gif)

## Features

### Cookie Management
- View all cookies in browser
- Import/export cookies as JSON
- Search and filter cookies
- Clear all cookies
- Incognito session support

### Microsoft 365 Sessions
- Acquire OAuth tokens
- Store multiple M365 refresh tokens as sessions
- Manual and automatic token refresh
- Session import/export

### M365 Data Browsers
- **Grapth** - Execute Graph API queries and useful quick action queries
- **Mailbox** - Browse folders, read emails, search messages
- **OneDrive** - Navigate folders, search files, view file metadata
- **SharePoint** - Access sites and documents
- **Teams** - View teams and channels
- **Calendar** - Access calendar events
- **Directory** - Query users and groups

## Usage

Use this extension in a isolated browser that does not interfer with your normal session.
Preferly open a incognito/private window to fully isolate the sessions. 

## Installation

### Browser Extension Stores

**TODO**

### Manual Installation

A manual installation, where you sideload the extension is the most secure way you can use this 
extensions as it will NOT auto update. 

```bash
git clone https://github.com/phishingclub/session-sushi.git
cd session-sushi
```

Load in Chrome/Edge:
1. Navigate to `chrome://extensions/` or `edge://extensions/`
2. Enable "Developer mode"
3. Click "Load unpacked"
4. Select the `session-sushi` directory

## Development

```bash
git clone https://github.com/phishingclub/session-sushi.git
cd session-sushi
```

Load the extension in developer mode and make changes. Reload extension after modifications.

## Contributing

- No dependencies allowed - vanilla JavaScript only
- Feature / Bug fixes in bug-* or feature-* branches
- Rebase branch to a single commit when it is ready to review / merge
- Ensure the last commit is performed signed to agree with CLA.

## License Agreement

**Important**: All contributors must agree to our Contributor License Agreement (CLA).

By contributing to Phishing Club, you agree that your contributions will be licensed under the same dual license terms (AGPL-3.0 and commercial). You confirm that:

- You have the right to contribute the code
- Your contributions are your original work or properly attributed
- You grant Phishing Club the right to license your contributions under both AGPL-3.0 and commercial licenses

## License

GNU Affero General Public License v3.0 (AGPL-3.0). See [LICENSE](LICENSE).

## Support

- **Community**: [Phishing Club Discord](https://discord.gg/Zssps7U8gX)

## Related Projects

- [Phishing Club](https://github.com/phishingclub/phishingclub) - Phishing simulation platform
