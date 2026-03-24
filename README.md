# n8n-nodes-onedrivewithdriveid

Microsoft OneDrive node for n8n with Drive ID support.

## Installation

```bash
npm install n8n-nodes-onedrivewithdriveid
```

## Overview

This is a fork of the official Microsoft OneDrive node that adds support for specifying a custom Drive ID in the credentials. This is useful when you need to access a specific OneDrive drive other than the default user drive.

## Credentials

When configuring the credentials, you can optionally specify a **Drive ID**:

1. In n8n, go to **Credentials** and create a new **Microsoft Drive OAuth2 API (With Drive ID)**
2. Enter your Azure AD application credentials (Client ID, Client Secret)
3. Optionally enter a **Drive ID** to target a specific OneDrive drive
4. If no Drive ID is provided, the node will use the authenticated user's default drive

### Finding your Drive ID

To find your Drive ID:
1. Go to [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in and run: `GET /me/drives`
3. The response will list all drives with their IDs

## Operations

### File Operations
- Copy
- Delete
- Download
- Get
- Search
- Share
- Upload
- Rename

### Folder Operations
- Create
- Delete
- Get Children
- Search
- Share
- Rename

## Usage

1. Install the package
2. Restart n8n
3. Create credentials with your Azure AD app and optionally a Drive ID
4. Add the "Microsoft OneDrive (With Drive ID)" node to your workflow

## License

MIT