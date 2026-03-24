"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MicrosoftOneDriveWithDriveIdOAuth2Api = void 0;
class MicrosoftOneDriveWithDriveIdOAuth2Api {
    constructor() {
        this.name = 'microsoftOneDriveWithDriveIdOAuth2Api';
        this.extends = ['microsoftOAuth2Api'];
        this.displayName = 'Microsoft Drive OAuth2 API (With Drive ID)';
        this.documentationUrl = 'microsoft';
        this.properties = [
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
}
exports.MicrosoftOneDriveWithDriveIdOAuth2Api = MicrosoftOneDriveWithDriveIdOAuth2Api;
//# sourceMappingURL=MicrosoftOneDriveWithDriveIdOAuth2Api.credentials.js.map