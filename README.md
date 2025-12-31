# <img src="icons/icon48.png" alt="Session Sushi" width="42" align="top"> Session Sushi

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
Graph, User, Directory, Mailbox, Calendar, OneDrive, SharePoint and Teams.

## Usage

Use this extension in a isolated browser that does not interfer with your normal session.
Preferly open a incognito/private window to fully isolate the sessions. 

## Installation


### Browser Extension Stores

Download for [Edge](https://microsoftedge.microsoft.com/addons/detail/session-sushi/bbaaafoehpllebaidnjbggniedbhdeok)

Download for [Chrome](https://chromewebstore.google.com/detail/session-sushi/mlfopacnocgoemdlknapgfpdjcjmefgja)

### Manual Installation

The most secure way you can use an extension is by sideloading it manually, this way it does not auto update.

**Option 1: Clone with Git**

```bash
git clone https://github.com/phishingclub/session-sushi.git
```

**Option 2: Download ZIP**

1. Go to [https://github.com/phishingclub/session-sushi](https://github.com/phishingclub/session-sushi)
2. Click the "Code" button
3. Select "Download ZIP"
4. Extract the ZIP file to a location of your choice

**Load in Chrome/Edge:**

1. Navigate to `chrome://extensions/` or `edge://extensions/`
2. Enable "Developer mode"
3. Click "Load unpacked"
4. Select the `session-sushi` directory (or the extracted folder if you downloaded the ZIP)

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
