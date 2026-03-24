import type { ICredentialType, INodeProperties } from 'n8n-workflow';

export class MicrosoftOneDriveWithDriveIdOAuth2Api implements ICredentialType {
	name = 'microsoftOneDriveWithDriveIdOAuth2Api';

	extends = ['microsoftOAuth2Api'];

	displayName = 'Microsoft Drive OAuth2 API (With Drive ID)';

	documentationUrl = 'microsoft';

	properties: INodeProperties[] = [
		{
			displayName: 'Drive ID',
			name: 'driveId',
			type: 'string',
			default: '',
			description: 'The ID of the OneDrive drive to use. If not provided, the default user drive will be used.',
		},
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default: 'openid offline_access Files.ReadWrite.All',
		},
	];
}